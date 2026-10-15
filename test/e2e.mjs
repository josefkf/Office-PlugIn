// End-to-end test of the task pane with office.js and Microsoft Graph mocked.
// Run: node build.mjs --mock-auth && node test/e2e.mjs
import { chromium } from "playwright";
import { createServer } from "node:http";
import { readFileSync, existsSync } from "node:fs";
import { extname, join } from "node:path";
import { deflateSync } from "node:zlib";

const DIST = new URL("../dist/", import.meta.url).pathname;
const MOCK_OFFICE = readFileSync(new URL("./mock-office.js", import.meta.url));
const SHOTS = process.env.SHOTS_DIR;

// --- tiny static server for dist/
const types = { ".html": "text/html", ".js": "text/javascript", ".css": "text/css", ".png": "image/png" };
const server = createServer((req, res) => {
  const p = join(DIST, decodeURIComponent(new URL(req.url, "http://x").pathname));
  if (!existsSync(p)) { res.writeHead(404); return res.end(); }
  res.writeHead(200, { "content-type": types[extname(p)] || "application/octet-stream" });
  res.end(readFileSync(p));
}).listen(0);
const BASE = `http://localhost:${server.address().port}`;

// --- real PNGs of a given size, so naturalWidth-based scaling is exercised
function png(w, h, rgb) {
  const crcTable = Array.from({ length: 256 }, (_, n) => { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; return c >>> 0; });
  const crc = (b) => { let c = ~0; for (const x of b) c = crcTable[(c ^ x) & 255] ^ (c >>> 8); return (~c) >>> 0; };
  const chunk = (t, d) => { const l = Buffer.alloc(4); l.writeUInt32BE(d.length); const td = Buffer.concat([Buffer.from(t), d]); const c = Buffer.alloc(4); c.writeUInt32BE(crc(td)); return Buffer.concat([l, td, c]); };
  const ihdr = Buffer.alloc(13); ihdr.writeUInt32BE(w, 0); ihdr.writeUInt32BE(h, 4); ihdr[8] = 8; ihdr[9] = 2;
  const row = Buffer.concat([Buffer.from([0]), Buffer.alloc(w * 3).fill(Buffer.from(rgb))]);
  const raw = Buffer.concat(Array.from({ length: h }, () => row));
  return Buffer.concat([Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]), chunk("IHDR", ihdr), chunk("IDAT", deflateSync(raw)), chunk("IEND", Buffer.alloc(0))]);
}
const BIG = png(800, 400, [15, 108, 189]);   // 600pt wide -> should be scaled to 200pt
const SMALL = png(100, 100, [200, 30, 30]);  // 75pt wide -> keep own size
const SVG = '<svg xmlns="http://www.w3.org/2000/svg" width="64" height="64"><circle cx="32" cy="32" r="30" fill="#0f6cbd"/></svg>';

const img = (id, name, thumb = true) => ({
  id, name, file: { mimeType: "x" },
  "@microsoft.graph.downloadUrl": `https://download.test/${id}`,
  thumbnails: thumb ? [{ medium: { url: `https://thumbs.test/${id}.png` } }] : [],
});
const graph = {
  "/v1.0/sites/contoso.sharepoint.com:/sites/brand": { id: "site1" },
  "/v1.0/sites/site1/drives": { value: [{ id: "docs", name: "Documents" }, { id: "drv", name: "Brand Images" }] },
  // root is paged: page 1 has a nextLink
  "/v1.0/drives/drv/root/children": {
    value: [{ id: "fB", name: "Company B", folder: {} }, img("logo", "logo-big.png"), { id: "doc", name: "brief.docx", file: {} }],
    "@odata.nextLink": "https://graph.microsoft.com/v1.0/drives/drv/root/children?page=2",
  },
  "/v1.0/drives/drv/root/children?page=2": { value: [{ id: "fA", name: "Company A", folder: {} }, img("ico", "icon.svg", false)] },
  "/v1.0/drives/drv/items/fA/children": { value: [{ id: "fA1", name: "Consulting", folder: {} }, img("small", "badge-small.png")] },
  "/v1.0/drives/drv/items/fA1/children": { value: [] },
  "/v1.0/drives/drv/root/search(q='badge')": { value: [img("small", "badge-small.png"), { id: "fZ", name: "badge folder", folder: {} }] },
};
const downloads = { logo: [BIG, "image/png"], small: [SMALL, "image/png"], ico: [SVG, "image/svg+xml"] };

let failures = 0;
const check = (cond, msg) => { console.log(`${cond ? "PASS" : "FAIL"}  ${msg}`); if (!cond) failures++; };

const browser = await chromium.launch();
async function openPane(host, extra = "") {
  const page = await browser.newPage({ viewport: { width: 340, height: 640 } });
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  const graphCalls = [];
  await page.route("https://appsforoffice.microsoft.com/**", (r) => r.fulfill({ contentType: "text/javascript", body: MOCK_OFFICE }));
  await page.route("https://graph.microsoft.com/**", (r) => {
    const u = new URL(r.request().url());
    graphCalls.push({ url: u.pathname + u.search, auth: r.request().headers()["authorization"] });
    const key = u.searchParams.has("page") ? u.pathname + "?page=" + u.searchParams.get("page") : decodeURIComponent(u.pathname);
    const body = graph[key];
    return body ? r.fulfill({ json: body }) : r.fulfill({ status: 404, json: { error: { message: `no mock for ${key}` } } });
  });
  await page.route("https://download.test/**", (r) => {
    const [b, t] = downloads[r.request().url().split("/").pop()];
    return r.fulfill({ body: b, contentType: t, headers: { "access-control-allow-origin": "*" } });
  });
  await page.route("https://thumbs.test/**", (r) => r.fulfill({ body: SMALL, contentType: "image/png" }));
  await page.goto(`${BASE}/taskpane.html?host=${host}${extra}`);
  await page.waitForSelector(".tile");
  return { page, errors, graphCalls };
}
const calls = (page) => page.evaluate(() => window.__calls.map(({ pic, shape, ...c }) => ({ ...c, width: (pic || shape)?.width, lock: (pic || shape)?.lockAspectRatio, left: shape?.left, top: shape?.top })));
const insert = async (page, name) => {
  await page.locator(".tile", { hasText: name }).click();
  await page.waitForFunction((n) => document.getElementById("status").textContent.includes(`Inserted ${n}`) || document.getElementById("status").classList.contains("error"), name);
  return page.locator("#status").textContent();
};

// ---- PowerPoint: browsing, paging, filtering, folders, breadcrumb, search, insert
{
  const { page, errors, graphCalls } = await openPane("PowerPoint");
  const folders = await page.locator(".folder").allTextContents();
  const tiles = await page.locator(".tile .name").allTextContents();
  check(JSON.stringify(folders) === '["Company A","Company B"]', `folders from both pages, sorted: ${folders}`);
  check(JSON.stringify(tiles) === '["icon.svg","logo-big.png"]', `images only (docx filtered), sorted: ${tiles}`);
  check(await page.locator(".tile", { hasText: "icon.svg" }).locator(".thumb").textContent() === "SVG", "no-thumbnail item shows extension label");
  check(graphCalls.every((c) => c.auth === "Bearer test-token"), "every Graph call carries the bearer token");
  check(graphCalls.some((c) => c.url.includes("$expand=thumbnails")), "children requested with thumbnails");
  if (SHOTS) await page.screenshot({ path: `${SHOTS}/1-root.png` });

  let status = await insert(page, "logo-big.png");
  let c = (await calls(page)).at(-1);
  check(status === "Inserted logo-big.png" && c.api === "setSelectedDataAsync" && c.opts.coercionType === "image", `PowerPoint raster via Image coercion (${status})`);
  check(c.head.startsWith("iVBORw0KGgo"), "payload is base64 PNG");
  check(c.opts.imageWidth === 200, `800px image capped to 200pt (got ${c.opts.imageWidth})`);

  status = await insert(page, "icon.svg");
  c = (await calls(page)).at(-1);
  check(c.opts.coercionType === "xmlSvg" && c.head.startsWith("<svg"), `PowerPoint SVG via XmlSvg coercion (${status})`);

  await page.locator(".folder", { hasText: "Company A" }).click();
  await page.waitForFunction(() => document.querySelector(".crumb.current")?.textContent === "Company A");
  check(JSON.stringify(await page.locator(".folder").allTextContents()) === '["Consulting"]', "navigates into company folder");
  await page.locator(".folder", { hasText: "Consulting" }).click();
  await page.waitForSelector("#empty:not([hidden])");
  check(await page.locator(".crumb").count() === 3, "breadcrumb Home › Company A › Consulting");
  if (SHOTS) await page.screenshot({ path: `${SHOTS}/2-empty-subfolder.png` });
  await page.locator("button.crumb", { hasText: "Home" }).click();
  await page.waitForFunction(() => document.querySelector(".crumb.current")?.textContent === "Home" && document.querySelectorAll(".tile").length === 2, null, { timeout: 5000 }).catch(() => {});
  check(await page.locator(".tile").count() === 2, "breadcrumb back to Home");

  await page.fill("#search", "badge");
  await page.waitForFunction(() => document.getElementById("status").textContent === "1 result");
  check(await page.locator(".folder").count() === 0 && await page.locator(".tile").count() === 1, "search shows images only, no folders");
  check(await page.locator("#breadcrumb").isHidden(), "breadcrumb hidden during search");
  if (SHOTS) await page.screenshot({ path: `${SHOTS}/3-search.png` });
  await page.fill("#search", "");
  await page.waitForFunction(() => document.querySelectorAll(".tile").length === 2);
  check(await page.locator("#breadcrumb").isVisible(), "clearing search returns to folder view");
  check(errors.length === 0, `no page errors ${errors.join("; ")}`);
  await page.close();
}

// ---- Word: Word API path, scaling only when needed
{
  const { page, errors } = await openPane("Word");
  await insert(page, "logo-big.png");
  let c = (await calls(page)).at(-1);
  check(c.api === "Word.insertInlinePictureFromBase64" && c.loc === "After", "Word raster via insertInlinePictureFromBase64 (After selection)");
  check(c.width === 200 && c.lock === true, `Word width 200pt with aspect lock (got ${c.width})`);
  await page.locator(".folder", { hasText: "Company A" }).click();
  await page.waitForSelector(".tile >> text=badge-small.png");
  await insert(page, "badge-small.png");
  c = (await calls(page)).at(-1);
  check(c.width === undefined, "small image keeps its own size");
  await page.locator("button.crumb", { hasText: "Home" }).click();
  await page.waitForSelector(".tile >> text=icon.svg");
  await insert(page, "icon.svg");
  c = (await calls(page)).at(-1);
  check(c.api === "setSelectedDataAsync" && c.opts.coercionType === "xmlSvg", "Word SVG falls back to XmlSvg coercion");
  check(errors.length === 0, "no page errors");
  await page.close();
}

// ---- Excel: shapes API, placed at active cell
{
  const { page, errors } = await openPane("Excel");
  await insert(page, "logo-big.png");
  let c = (await calls(page)).at(-1);
  check(c.api === "Excel.addImage" && c.left === 120 && c.top === 45 && c.width === 200, `Excel addImage at active cell, 200pt (left ${c.left}, top ${c.top}, w ${c.width})`);
  await insert(page, "icon.svg");
  c = (await calls(page)).at(-1);
  check(c.api === "Excel.addSvg" && c.head.startsWith("<svg"), "Excel SVG via addSvg");
  if (SHOTS) await page.screenshot({ path: `${SHOTS}/4-inserted.png` });
  check(errors.length === 0, "no page errors");
  await page.close();
}

// ---- Older clients without WordApi/ExcelApi 1.x: common API fallback
{
  const { page } = await openPane("Excel", "&legacy=1");
  await insert(page, "logo-big.png");
  const c = (await calls(page)).at(-1);
  check(c.api === "setSelectedDataAsync" && c.opts.coercionType === "image", "legacy Excel falls back to Image coercion");
  await page.close();
}

// ---- Error surface: library missing
{
  graph["/v1.0/sites/site1/drives"] = { value: [{ id: "docs", name: "Documents" }] };
  const page = await browser.newPage();
  await page.route("https://appsforoffice.microsoft.com/**", (r) => r.fulfill({ contentType: "text/javascript", body: MOCK_OFFICE }));
  await page.route("https://graph.microsoft.com/**", (r) => {
    const u = new URL(r.request().url()); const b = graph[decodeURIComponent(u.pathname)];
    return b ? r.fulfill({ json: b }) : r.fulfill({ status: 404, json: {} });
  });
  await page.goto(`${BASE}/taskpane.html?host=Word`);
  await page.waitForSelector(".status.error");
  const msg = await page.locator("#status").textContent();
  check(msg.includes('Library "Brand Images" not found'), `clear error when library is missing: ${msg}`);
  await page.close();
}

await browser.close();
server.close();
console.log(failures ? `\n${failures} FAILED` : "\nALL PASSED");
process.exit(failures ? 1 : 0);
