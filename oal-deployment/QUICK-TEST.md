# Quick functionality test: Office template library (manual)

This is a minimal test before the scripted deployment: a test site, one library you upload templates to by hand, and the registration as an Organization Assets Library.

Run the commands in **Windows PowerShell 5.1**, or in PowerShell 7 after:

```powershell
Import-Module Microsoft.Online.SharePoint.PowerShell -UseWindowsPowerShell
```

## 1. Test site (skip if you use an existing site)

```powershell
$tenant = "contoso"   # change
Connect-SPOService -Url "https://$tenant-admin.sharepoint.com"

New-SPOSite -Url "https://$tenant.sharepoint.com/sites/OALTest" `
  -Title "OAL Test" -Owner "you@contoso.com" `
  -Template "SITEPAGEPUBLISHING#0" -StorageQuota 1024
```

## 2. Library and permissions (SharePoint UI)

1. On the site: **New → Document library**, named `Templates`.
2. Upload 2–3 templates into it, all saved as **.dotx / .potx / .xltx**. Put one in a subfolder (for example `Company A/Consulting`) to see how folders appear in Office.
3. **Settings → Site permissions → Advanced permissions settings → OAL Test Visitors → New**, and add **Everyone except external users**. Microsoft's docs require this. Without it, other users won't see the templates.
4. Optional: upload a thumbnail image to **Site Assets** (see [Thumbnails](#thumbnails)).

## 3. Register as an Office template library

```powershell
$site = "https://$tenant.sharepoint.com/sites/OALTest"

Add-SPOOrgAssetsLibrary -LibraryUrl "$site/Templates" `
  -OrgAssetType OfficeTemplateLibrary `
  -ThumbnailUrl "$site/SiteAssets/oal-thumbnail.png"   # drop this line if no thumbnail
# It asks to enable the CDN: answer Y (or add -Confirm:$false)

Get-SPOOrgAssetsLibrary   # check it's listed
```

## 4. Test

- **When:** allow **up to 24h** for templates to appear. It's often faster.
- **Desktop:** Word / PowerPoint / Excel → **File → New**. Look for a tab with your organization's name, and check how the subfolder shows.
- **PowerPoint on the web** (E3/E5 only): look under **Office Template Library** on the start page.
- **Permissions:** test with a **normal user account**, not only an admin.
- **Requirement:** desktop apps need **Microsoft 365 Apps version 2002 or later**.

## Change or clean up

```powershell
# Change the thumbnail
Set-SPOOrgAssetsLibrary -LibraryUrl "$site/Templates" -ThumbnailUrl "$site/SiteAssets/new.png"

# Unregister the library
Remove-SPOOrgAssetsLibrary -LibraryUrl "$site/Templates"

# Removing the library leaves its CDN origin in place
Remove-SPOTenantCdnOrigin -CdnType Private -OriginUrl "/sites/OALTest/Templates"

# Optional: delete the test site (goes to the recycle bin)
Remove-SPOSite -Identity $site
```

## Thumbnails

There are two different "thumbnails".

### 1. Library thumbnail (`-ThumbnailUrl`)

- Per Microsoft, it's the background image of the library's **card**, mainly shown in the SharePoint file picker under "Your organization". Without one, the card is plain gray.
- It **must be on the same site** as the library, for example in Site Assets.
- Microsoft doesn't document a size. Use a small landscape PNG/JPG, a few hundred pixels wide, with the logo on a brand color. It's a background, so keep text minimal.
- It's cosmetic and can be changed any time with `Set-SPOOrgAssetsLibrary`.
- Not confirmed: whether it appears anywhere in the Office template tab. Check during the test.

### 2. Template previews (what users see in File → New)

- Each template's tile shows the **preview picture saved inside the file**. A template saved without one can show a generic icon.
- **To save the preview:**
  - **Word / Excel:** tick **Save Thumbnail** in the Save As dialog, or File → Info → Properties → Advanced Properties → Summary → **Save preview picture**.
  - **PowerPoint:** same Summary tab, **Save preview picture**. The preview is the first slide, so make that slide representative.
- **Name files clearly,** for example "Company A – Proposal". The file name is what users read under the preview.

## Next step

Once the test looks good, use the scripts in this folder for the real deployment (see [README.md](README.md)).
