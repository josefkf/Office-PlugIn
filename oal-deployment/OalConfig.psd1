@{
    # --- Tenant -------------------------------------------------------------
    TenantName           = 'contoso'                    # contoso.sharepoint.com -> 'contoso'
    PnPClientId          = '00000000-0000-0000-0000-000000000000'  # Entra app for PnP interactive login (see 00-Install-Prerequisites.ps1)

    # --- Asset site ---------------------------------------------------------
    SiteUrl              = 'https://contoso.sharepoint.com/sites/BrandTemplates'
    SiteTitle            = 'Brand Templates'
    SiteOwner            = 'admin@contoso.com'

    # People who maintain templates (site Members = edit). Users by UPN and/or Entra security groups by object ID.
    EditorUpns           = @()                          # e.g. @('marketing.lead@contoso.com')
    EditorGroupObjectIds = @()                          # e.g. @('11111111-2222-3333-4444-555555555555')

    # --- Template library (registered as OfficeTemplateLibrary) -------------
    TemplateLibraryTitle = 'Templates'
    TemplateLibraryUrl   = 'Templates'                  # URL segment: /sites/BrandTemplates/Templates
    LocalTemplateRoot    = 'C:\OAL\Templates'           # Local folder; subfolders (company / service line) are recreated as-is

    # Content approval: editors' uploads stay invisible to users until approved.
    EnableContentApproval = $true
    MajorVersionLimit     = 50

    # Optional card background in the Office template picker. Leave '' to skip.
    ThumbnailPath        = ''                           # e.g. 'C:\OAL\thumbnail.png'

    # Make the library available to Microsoft 365 Copilot search. $null = don't set.
    CopilotSearchable    = $null

    # --- Optional image library (registered as ImageDocumentLibrary) --------
    # Feeds "Brand Images" in PowerPoint and the SharePoint image picker.
    ImageLibrary = @{
        Enabled   = $false
        Title     = 'Brand Images'
        Url       = 'BrandImages'
        LocalRoot = 'C:\OAL\Images'
    }
}
