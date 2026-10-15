import { config } from "./config.js";

const PX_TO_PT = 0.75;

function blobToBase64(blob) {
  return new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onload = () => resolve(String(r.result).split(",")[1]);
    r.onerror = () => reject(r.error);
    r.readAsDataURL(blob);
  });
}

function naturalSize(blob) {
  return new Promise((resolve) => {
    const url = URL.createObjectURL(blob);
    const img = new Image();
    img.onload = () => { resolve({ w: img.naturalWidth, h: img.naturalHeight }); URL.revokeObjectURL(url); };
    img.onerror = () => { resolve(null); URL.revokeObjectURL(url); };
    img.src = url;
  });
}

// Target width in points, or null to keep the image's own size.
async function targetWidthPt(blob) {
  const size = await naturalSize(blob);
  if (!size || !size.w) return null;
  const widthPt = size.w * PX_TO_PT;
  return widthPt > config.maxInsertWidthPt ? config.maxInsertWidthPt : null;
}

function setSelectedData(data, options) {
  return new Promise((resolve, reject) => {
    Office.context.document.setSelectedDataAsync(data, options, (r) =>
      r.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(new Error(r.error.message))
    );
  });
}

async function insertWord(base64, widthPt) {
  await Word.run(async (ctx) => {
    const pic = ctx.document.getSelection().insertInlinePictureFromBase64(base64, "After");
    if (widthPt) {
      pic.lockAspectRatio = true;
      pic.width = widthPt;
    }
    pic.select("End");
    await ctx.sync();
  });
}

async function insertExcel(payload, isSvg, widthPt) {
  await Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getActiveWorksheet();
    const cell = ctx.workbook.getActiveCell();
    cell.load("left,top");
    const shape = isSvg ? sheet.shapes.addSvg(payload) : sheet.shapes.addImage(payload);
    await ctx.sync();
    shape.left = cell.left;
    shape.top = cell.top;
    if (widthPt) {
      shape.lockAspectRatio = true;
      shape.width = widthPt;
    }
    await ctx.sync();
  });
}

export async function insertImage(image, blob) {
  const host = Office.context.host;
  const isSvg = image.ext === "svg";
  const widthPt = await targetWidthPt(blob);
  const req = Office.context.requirements;

  if (host === Office.HostType.Excel && req.isSetSupported("ExcelApi", "1.9")) {
    const payload = isSvg ? await blob.text() : await blobToBase64(blob);
    return insertExcel(payload, isSvg, widthPt);
  }

  if (host === Office.HostType.Word && !isSvg && req.isSetSupported("WordApi", "1.2")) {
    return insertWord(await blobToBase64(blob), widthPt);
  }

  // PowerPoint, SVG in Word, and older clients: common API.
  const opts = {};
  if (widthPt) opts.imageWidth = widthPt;
  if (isSvg) {
    if (!req.isSetSupported("ImageCoercion", "1.2")) throw new Error("SVG insert isn't supported in this Office version. Use a PNG.");
    return setSelectedData(await blob.text(), { ...opts, coercionType: Office.CoercionType.XmlSvg });
  }
  return setSelectedData(await blobToBase64(blob), { ...opts, coercionType: Office.CoercionType.Image });
}
