import { config } from "./config.js";
import { getToken } from "./auth.js";

const GRAPH = "https://graph.microsoft.com/v1.0";
let driveIdPromise;

async function graphGet(pathOrUrl) {
  const token = await getToken();
  const url = pathOrUrl.startsWith("https://") ? pathOrUrl : GRAPH + pathOrUrl;
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (!res.ok) {
    let detail = "";
    try { detail = (await res.json()).error?.message || ""; } catch { /* ignore */ }
    throw new Error(`Graph ${res.status}${detail ? ": " + detail : ""}`);
  }
  return res.json();
}

async function getAll(path) {
  const items = [];
  let next = path;
  while (next) {
    const page = await graphGet(next);
    items.push(...(page.value || []));
    next = page["@odata.nextLink"];
  }
  return items;
}

export function getDriveId() {
  if (!driveIdPromise) {
    driveIdPromise = (async () => {
      const site = await graphGet(`/sites/${config.siteHostname}:${config.sitePath}?$select=id`);
      const drives = await getAll(`/sites/${site.id}/drives?$select=id,name`);
      const drive = drives.find((d) => d.name.toLowerCase() === config.libraryName.toLowerCase());
      if (!drive) throw new Error(`Library "${config.libraryName}" not found on ${config.sitePath}`);
      return drive.id;
    })();
    driveIdPromise.catch(() => { driveIdPromise = undefined; }); // allow retry
  }
  return driveIdPromise;
}

function ext(name) {
  const i = name.lastIndexOf(".");
  return i < 0 ? "" : name.slice(i + 1).toLowerCase();
}

function normalize(items) {
  const folders = [];
  const images = [];
  for (const it of items) {
    if (it.folder) {
      folders.push({ id: it.id, name: it.name });
    } else if (it.file && config.extensions.includes(ext(it.name))) {
      images.push({
        id: it.id,
        name: it.name,
        ext: ext(it.name),
        thumb: it.thumbnails?.[0]?.medium?.url || it.thumbnails?.[0]?.small?.url || null,
        downloadUrl: it["@microsoft.graph.downloadUrl"] || null,
      });
    }
  }
  const byName = (a, b) => a.name.localeCompare(b.name, undefined, { numeric: true });
  return { folders: folders.sort(byName), images: images.sort(byName) };
}

export async function listFolder(folderId) {
  const driveId = await getDriveId();
  const base = folderId ? `/drives/${driveId}/items/${folderId}` : `/drives/${driveId}/root`;
  return normalize(await getAll(`${base}/children?$expand=thumbnails&$top=200`));
}

export async function searchLibrary(query) {
  const driveId = await getDriveId();
  const q = encodeURIComponent(query.replace(/'/g, "''"));
  const res = normalize(await getAll(`/drives/${driveId}/root/search(q='${q}')?$expand=thumbnails&$top=200`));
  return { folders: [], images: res.images };
}

export async function downloadImage(image) {
  // downloadUrl is pre-authenticated and short-lived; fall back to /content with a token.
  if (image.downloadUrl) {
    try {
      const res = await fetch(image.downloadUrl);
      if (res.ok) return res.blob();
    } catch { /* fall through */ }
  }
  const driveId = await getDriveId();
  const token = await getToken();
  const res = await fetch(`${GRAPH}/drives/${driveId}/items/${image.id}/content`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!res.ok) throw new Error(`Download failed (${res.status})`);
  return res.blob();
}
