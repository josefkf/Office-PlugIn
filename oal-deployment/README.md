# Organization Assets Library: deployment scripts

These scripts deploy Word, PowerPoint and Excel templates from a SharePoint library, so they show up in **File → New → [your organization]** in the Office apps.

| Script | What it does | Run |
|---|---|---|
| `00-Install-Prerequisites.ps1` | Installs PnP.PowerShell and the SharePoint Online module. `-RegisterPnPApp` creates the Entra app PnP needs. | Once |
| `01-New-AssetSite.ps1` | Creates the communication site and libraries (versioning, optional content approval), adds "Everyone except external users" as Visitors and editors as Members, and uploads an optional thumbnail. | Once, then safe to re-run |
| `02-Publish-Templates.ps1` | Uploads `.dotx/.potx/.xltx` from a local folder tree, keeping the folders (company / service line). Supports `-WhatIf` and `-IncludeImages`. | On every template change |
| `03-Register-OrgAssetsLibrary.ps1` | `Add-SPOOrgAssetsLibrary` for the template library (and optional image library), then verifies. | Once |
| `99-Remove-OrgAssetsLibrary.ps1` | Rollback: unregisters, removes CDN origins, and deletes the site with `-DeleteSite`. | If needed |

All settings live in `OalConfig.psd1`.

To try it by hand first, see **[QUICK-TEST.md](QUICK-TEST.md)**: a manual cmdlet-by-cmdlet functionality test, plus notes on thumbnails.

## Requirements
- PowerShell 7.4+ on Windows. The SPO module loads through the Windows PowerShell compatibility layer, which is handled automatically.
- SharePoint Administrator. Registering the PnP app needs admin consent.
- Microsoft 365 Apps 2002+ on clients. PowerPoint on the web needs E3/E5.

## Run order
```powershell
.\00-Install-Prerequisites.ps1 -RegisterPnPApp -Tenant contoso.onmicrosoft.com   # copy the ClientId into OalConfig.psd1
# edit OalConfig.psd1
.\01-New-AssetSite.ps1
.\02-Publish-Templates.ps1 -WhatIf     # preview
.\02-Publish-Templates.ps1
.\03-Register-OrgAssetsLibrary.ps1
```

## Local folder layout (example)
```
C:\OAL\Templates\
├── Company A\
│   ├── Consulting\  Proposal.potx, Report.dotx
│   └── Engineering\ Technical Note.dotx
└── Company B\       Letter.dotx, Budget.xltx
```

## Notes
- **Visibility:** the OAL requires "Everyone except external users" as Visitors, so all users see all templates.
- **Content approval** (on by default): files editors upload in SharePoint stay hidden until approved. Files uploaded by `02` are approved automatically.
- **Updates:** re-run `02`. Changes can take up to 24h to reach the Office apps.
- **Subfolders:** check how company/service-line subfolders appear in the Office picker in your tenant before announcing the structure.
- **Not covered here:** an Intune policy to hide Office's built-in templates, and font deployment.
