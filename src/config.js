// Edit these values for your tenant, then run `npm run build`.
export const config = {
  // Entra app registration (SPA, with redirect URI brk-multihub://<your add-in host>)
  clientId: "00000000-0000-0000-0000-000000000000",
  tenantId: "00000000-0000-0000-0000-000000000000",

  // SharePoint site + library holding the brand assets (can be the OAL image library)
  siteHostname: "contoso.sharepoint.com",
  sitePath: "/sites/brand",
  libraryName: "Brand Images", // library display name as shown in SharePoint

  // Delegated Graph scopes (read-only)
  scopes: ["Sites.Read.All"],

  // Max inserted width in points. Larger images are scaled down, aspect ratio kept.
  maxInsertWidthPt: 200,

  // File types shown in the picker
  extensions: ["png", "jpg", "jpeg", "gif", "bmp", "svg"],
};
