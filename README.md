# Brand Assets add-in (Word, Excel, PowerPoint)

A task pane that browses a SharePoint library (for example, your OAL image library) through Microsoft Graph and inserts the image on click. It works in Word, Excel and PowerPoint on Windows, Mac and the web.

- **Auth:** Nested App Authentication (NAA). It reuses the Office sign-in, so there's no separate login page.
- **Permissions:** `Sites.Read.All` (delegated, read-only). Users only see what they can already read in SharePoint, so per-company or per-service-line restrictions on folders apply automatically.
- **Browsing:** folders with breadcrumb navigation, library-wide search, and thumbnails from SharePoint.
- **Insert behavior:**

| Host | Raster (PNG/JPG/GIF/BMP) | SVG |
|---|---|---|
| Word | `insertInlinePictureFromBase64`, after the cursor | `setSelectedDataAsync` (XmlSvg) |
| Excel | `shapes.addImage` at the active cell | `shapes.addSvg` at the active cell |
| PowerPoint | `setSelectedDataAsync` (Image) | `setSelectedDataAsync` (XmlSvg) |

Images wider than `maxInsertWidthPt` (default 200pt) are scaled down with the aspect ratio locked. Older clients fall back to the common API.

## 1. App registration (Entra ID)

1. **App registrations → New registration**, name "Brand Assets add-in", single tenant.
2. **Authentication → Add a platform → Single-page application**, redirect URI:
   - `brk-multihub://localhost:3000` (dev)
   - `brk-multihub://assets.contoso.com` (prod: your hosting domain, no path)
3. **API permissions → Microsoft Graph → Delegated → `Sites.Read.All`**, then **Grant admin consent**.
4. Copy the **Application (client) ID** and **Directory (tenant) ID**.

## 2. Configure

Edit `src/config.js`: `clientId`, `tenantId`, `siteHostname`, `sitePath`, `libraryName` (the library's display name).

## 3. Run locally and sideload (Windows)

```powershell
npm install
npm run dev-certs        # trusts a localhost HTTPS cert (one time)
npm run build
npm start                # serves dist/ on https://localhost:3000
```

Sideload `dist/manifest.xml`:
- **Web:** open Word, Excel or PowerPoint on the web → Home → Add-ins → More Settings → Upload My Add-in.
- **Desktop:** `npx office-addin-debugging start dist/manifest.xml desktop --app word` (or `excel` / `powerpoint`).

The **Brand Assets** button appears on the Home tab.

## 4. Deploy to production

1. Host the static files on any HTTPS host. Azure Static Web Apps (free tier) is the easy option. SharePoint can't host it, because it doesn't serve HTML pages as apps.
2. Build for that URL: `BASE_URL=https://assets.contoso.com PROVIDER="Hempel IT" npm run build` (PowerShell: `$env:BASE_URL="https://assets.contoso.com"; npm run build`). Upload `dist/`.
3. Add the matching `brk-multihub://` redirect URI to the app registration.
4. **Microsoft 365 admin center → Settings → Integrated apps → Upload custom apps → Office Add-in**, then upload `dist/manifest.xml` and assign it to users or groups. It can take up to 24h to appear.

To ship an update, upload the new files to the host. Change the manifest (and bump `<Version>`) only if the ribbon, URLs or icons change.

## Tests

```bash
npm test   # builds with mocked auth, runs headless Chromium against mocked office.js + Graph
```

The tests cover folder paging, file filtering, the breadcrumb, search, the insert path for each host, image scaling, the legacy fallback and the missing-library error. They don't replace a real-tenant test.

## Known limits and things to verify on your tenant

- **NAA needs a recent Microsoft 365 Apps build.** Clients that are too old get a clear message in the pane, but there's no fallback sign-in. Microsoft's NAA sample shows a dialog-based fallback if you need one.
- **Image download** uses the item's pre-authenticated `downloadUrl` and falls back to `/content` with a token. Verify both on your tenant, since CORS behavior is the most likely real-world snag.
- **SVG** insertion needs ImageCoercion 1.2 in Word and PowerPoint. In that case, use PNG versions of the icons.
- **Search** uses OneDrive search over the library. New files can take a few minutes to become searchable, but browsing is instant.
