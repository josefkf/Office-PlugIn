// Minimal office.js stand-in. Host and supported sets come from the page URL: ?host=Word|Excel|PowerPoint
(() => {
  const p = new URLSearchParams(location.search);
  const host = p.get("host") || "PowerPoint";
  const legacy = p.get("legacy") === "1";
  const calls = (window.__calls = []);
  const sets = { NestedAppAuth: true, ImageCoercion: true, WordApi: !legacy, ExcelApi: !legacy };
  const run = (ctx) => async (fn) => fn(ctx);
  window.Office = {
    HostType: { Word: "Word", Excel: "Excel", PowerPoint: "PowerPoint" },
    CoercionType: { Image: "image", XmlSvg: "xmlSvg" },
    AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },
    onReady(cb) { setTimeout(() => cb({ host }), 0); return Promise.resolve({ host }); },
    context: {
      host,
      requirements: { isSetSupported: (n) => !!sets[n] },
      document: {
        setSelectedDataAsync(data, opts, cb) {
          calls.push({ api: "setSelectedDataAsync", len: data.length, head: data.slice(0, 12), opts });
          cb({ status: "succeeded" });
        },
      },
    },
  };
  const sync = async () => {};
  window.Word = { run: run({ sync, document: { getSelection: () => ({
    insertInlinePictureFromBase64(b64, loc) {
      const pic = { select() {} };
      calls.push({ api: "Word.insertInlinePictureFromBase64", len: b64.length, head: b64.slice(0, 12), loc, pic });
      return pic;
    } }) } }) };
  window.Excel = { run: run({ sync, workbook: {
    getActiveCell: () => ({ left: 120, top: 45, load() {} }),
    worksheets: { getActiveWorksheet: () => ({ shapes: {
      addImage(b64) { const s = {}; calls.push({ api: "Excel.addImage", len: b64.length, head: b64.slice(0, 12), shape: s }); return s; },
      addSvg(xml) { const s = {}; calls.push({ api: "Excel.addSvg", len: xml.length, head: xml.slice(0, 12), shape: s }); return s; },
    } }) },
  } }) };
})();
