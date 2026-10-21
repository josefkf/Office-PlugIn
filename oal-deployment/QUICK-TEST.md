# Quick functionality test: Office template library (PnP PowerShell)

This is a minimal test before the scripted deployment: a test site, one library you upload templates to by hand, and the registration as an Organization Assets Library, all with **PnP.PowerShell**.

## 0. Prerequisites

- **PowerShell 7.4+** and the PnP module:
  ```powershell
  Install-Module PnP.PowerShell -Scope CurrentUser
  ```
- **An Entra app registration for PnP.** Since September 2024 every tenant needs its own app. Create it once (needs admin consent) and note the **ClientId**:
  ```powershell
  Register-PnPEntraIDAppForInteractiveLogin -ApplicationName "PnP PowerShell" -Tenant contoso.onmicrosoft.com -Interactive
  ```
- **Your account needs SharePoint Administrator.**

```powershell
$tenant   = "contoso"                                  # change
$clientId = "00000000-0000-0000-0000-000000000000"     # your PnP app
$admin    = "https://$tenant-admin.sharepoint.com"
$site     = "https://$tenant.sharepoint.com/sites/OALTest"
```

## 1. Test site (skip if you use an existing site)

```powershell
Connect-PnPOnline -Url $admin -ClientId $clientId -Interactive

New-PnPSite -Type CommunicationSite -Title "OAL Test" -Url $site -Owner "you@contoso.com" -Wait
```

## 2. Library and permissions

```powershell
Connect-PnPOnline -Url $site -ClientId $clientId -Interactive

# Template library
New-PnPList -Title "Templates" -Url "Templates" -Template DocumentLibrary -OnQuickLaunch

# "Everyone except external users" as Visitors: required for the OAL
$tenantId = Get-PnPTenantId
$visitors = Get-PnPGroup -AssociatedVisitorGroup
Add-PnPGroupMember -Group $visitors -LoginName "c:0-.f|rolemanager|spo-grid-all-users/$tenantId"
```

Then **upload 2–3 templates by hand** into the `Templates` library in the browser, all saved as **.dotx / .potx / .xltx**. Put one in a subfolder (for example `Company A/Consulting`) to see how folders appear in Office.

Optional: upload a thumbnail (see [Thumbnails](#thumbnails)). It must be on the same site:

```powershell
New-PnPList -Title "AssetThumbnails" -Url "AssetThumbnails" -Template DocumentLibrary
Add-PnPFile -Path ".\oal-thumbnail.png" -Folder "AssetThumbnails"
```

## 3. Register as an Office template library

This is a tenant-level setting, so connect to the admin site:

```powershell
Connect-PnPOnline -Url $admin -ClientId $clientId -Interactive

Add-PnPOrgAssetsLibrary -LibraryUrl "$site/Templates" `
  -OrgAssetType OfficeTemplateLibrary `
  -CdnType Private `
  -ThumbnailUrl "$site/AssetThumbnails/oal-thumbnail.png"   # drop this line if no thumbnail

Get-PnPOrgAssetsLibrary   # check it's listed
```

> **Always pass `-CdnType Private`.** PnP's default is **Public**. A public CDN means the files can be fetched by URL without signing in, which you don't want for company templates. (The Microsoft SPO module defaults to Private.)

## 4. Test

- **When:** allow **up to 24h** for templates to appear. It's often faster.
- **Desktop:** Word / PowerPoint / Excel → **File → New**. Look for a tab with your organization's name, and check how the subfolder shows.
- **PowerPoint on the web** (E3/E5 only): look under **Office Template Library** on the start page.
- **Permissions:** test with a **normal user account**, not only an admin.
- **Requirement:** desktop apps need **Microsoft 365 Apps version 2002 or later**.

## Change or clean up

```powershell
Connect-PnPOnline -Url $admin -ClientId $clientId -Interactive

# Change the thumbnail
Set-PnPOrgAssetsLibrary -LibraryUrl "$site/Templates" -ThumbnailUrl "$site/AssetThumbnails/new.png"

# Unregister the library and remove its CDN origin in one go
Remove-PnPOrgAssetsLibrary -LibraryUrl "sites/OALTest/Templates" -ShouldRemoveFromCdn $true -CdnType Private

# Check no origin is left behind
Get-PnPTenantCdnOrigin -CdnType Private

# Optional: delete the test site (goes to the recycle bin)
Remove-PnPTenantSite -Url $site -Force
```

## Thumbnails

There are two different "thumbnails".

### 1. Library thumbnail (`-ThumbnailUrl`)

- Per Microsoft, it's the background image of the library's **card**, mainly shown in the SharePoint file picker under "Your organization". Without one, the card is plain gray.
- It **must be on the same site** as the library.
- Microsoft doesn't document a size. Use a small landscape PNG/JPG, a few hundred pixels wide, with the logo on a brand color. It's a background, so keep text minimal.
- It's cosmetic and can be changed any time with `Set-PnPOrgAssetsLibrary`.
- Not confirmed: whether it appears anywhere in the Office template tab. Check during the test.

### 2. Template previews (what users see in File → New)

- Each template's tile shows the **preview picture saved inside the file**. A template saved without one can show a generic icon.
- **To save the preview:**
  - **Word / Excel:** tick **Save Thumbnail** in the Save As dialog, or File → Info → Properties → Advanced Properties → Summary → **Save preview picture**.
  - **PowerPoint:** same Summary tab, **Save preview picture**. The preview is the first slide, so make that slide representative.
- **Name files clearly,** for example "Company A – Proposal". The file name is what users read under the preview.

## Next step

Once the test looks good, use the scripts in this folder for the real deployment (see [README.md](README.md)).
