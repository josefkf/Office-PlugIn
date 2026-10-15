import { isAuthSupported } from "./auth.js";
import { listFolder, searchLibrary, downloadImage } from "./graph.js";
import { insertImage } from "./insert.js";

const $ = (id) => document.getElementById(id);
const state = { path: [{ id: null, name: "Home" }], searchTimer: null, busy: false, seq: 0 };

function setStatus(msg, kind = "") {
  const el = $("status");
  el.textContent = msg;
  el.className = "status " + kind;
}

function renderBreadcrumb() {
  const bc = $("breadcrumb");
  bc.replaceChildren();
  state.path.forEach((p, i) => {
    if (i > 0) bc.append(Object.assign(document.createElement("span"), { className: "sep", textContent: "›" }));
    const last = i === state.path.length - 1;
    const el = document.createElement(last ? "span" : "button");
    el.textContent = p.name;
    el.className = last ? "crumb current" : "crumb";
    if (!last) el.onclick = () => { state.path = state.path.slice(0, i + 1); load(); };
    bc.append(el);
  });
}

function renderItems({ folders, images }) {
  const fl = $("folders");
  fl.replaceChildren();
  for (const f of folders) {
    const b = document.createElement("button");
    b.className = "folder";
    b.textContent = f.name;
    b.onclick = () => { state.path.push(f); load(); };
    fl.append(b);
  }

  const grid = $("grid");
  grid.replaceChildren();
  for (const img of images) {
    const tile = document.createElement("button");
    tile.className = "tile";
    tile.title = `Insert ${img.name}`;
    tile.dataset.id = img.id;
    const frame = document.createElement("div");
    frame.className = "thumb";
    if (img.thumb) {
      const i = document.createElement("img");
      i.src = img.thumb;
      i.alt = "";
      i.loading = "lazy";
      frame.append(i);
    } else {
      frame.textContent = img.ext.toUpperCase();
    }
    const label = document.createElement("span");
    label.className = "name";
    label.textContent = img.name;
    tile.append(frame, label);
    tile.onclick = () => onInsert(img, tile);
    grid.append(tile);
  }

  const empty = !folders.length && !images.length;
  $("empty").hidden = !empty;
}

// Each view change gets a sequence number; responses for an older view are dropped,
// so a slow folder or search request can't overwrite what the user is looking at now.
async function load() {
  const seq = ++state.seq;
  renderBreadcrumb();
  setStatus("Loading…");
  try {
    const folder = state.path[state.path.length - 1];
    const res = await listFolder(folder.id);
    if (seq !== state.seq) return;
    renderItems(res);
    setStatus("");
  } catch (e) {
    if (seq !== state.seq) return;
    renderItems({ folders: [], images: [] });
    setStatus(e.message, "error");
  }
}

async function runSearch(q) {
  if (!q) { $("breadcrumb").hidden = false; return load(); }
  const seq = ++state.seq;
  $("breadcrumb").hidden = true;
  setStatus("Searching…");
  try {
    const res = await searchLibrary(q);
    if (seq !== state.seq) return;
    renderItems(res);
    setStatus(`${res.images.length} result${res.images.length === 1 ? "" : "s"}`);
  } catch (e) {
    if (seq !== state.seq) return;
    setStatus(e.message, "error");
  }
}

async function onInsert(img, tile) {
  if (state.busy) return;
  state.busy = true;
  tile.classList.add("working");
  setStatus(`Inserting ${img.name}…`);
  try {
    const blob = await downloadImage(img);
    await insertImage(img, blob);
    setStatus(`Inserted ${img.name}`, "ok");
  } catch (e) {
    setStatus(e.message, "error");
  } finally {
    tile.classList.remove("working");
    state.busy = false;
  }
}

Office.onReady(() => {
  if (!isAuthSupported()) {
    setStatus("This Office version doesn't support single sign-on for add-ins (NestedAppAuth). Update Microsoft 365 Apps.", "error");
    return;
  }
  $("search").addEventListener("input", (e) => {
    clearTimeout(state.searchTimer);
    const q = e.target.value.trim();
    state.searchTimer = setTimeout(() => runSearch(q), 350);
  });
  load();
});
