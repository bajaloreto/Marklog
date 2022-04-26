# Connect and authenticate to SharePoint
$adminSiteUrl = "https://ecoalianzadeloretoac-admin.sharepoint.com/"
Connect-SPOService $adminSiteUrl

# Extract site script from site url
$siteUrl = "https://ecoalianzadeloretoac.sharepoint.com/sites/DesarrolloyOperacion"
$extracted = Get-SPOSiteScriptFromWeb `
-WebUrl $siteUrl `
-IncludeBranding `
-IncludeTheme `
-IncludeRegionalSettings `
-IncludeSiteExternalSharingCapability `
-IncludeLinksToExportedItems `
-IncludedLists ("Lists/Project%20progress%20tracker", "Lists/Project%20Issue%20tracker")

# Add site script to SharePoint
$SiteScript = Add-SPOSiteScript `
-Title "EAL Plantilla Base." `
-Description "Plantilla basica." `
-Content $extracted

# Create site design in SharePoint with the previous site script
$SiteDesign = Add-SPOSiteDesign -Title "Plantilla basica" -WebTemplate 64 -SiteScripts $SiteScript.Id
# </CreateSiteDesignInSharePoint>

# Disconnect from SharePoint services
Disconnect-SPOService
