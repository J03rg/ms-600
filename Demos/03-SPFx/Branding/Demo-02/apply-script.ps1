Connect-SPOService -Url https://integrationsonline-admin.sharepoint.com -credential alexander.pajer@integrations.at
Get-Content '.\site-script.json' -Raw | Add-SPOSiteScript -Title "MS-600-Design"
Add-SPOSiteDesign -Title "MS-600 Example" -WebTemplate "64" -SiteScripts "4991aec5-723b-4d80-bdd3-a059bd0a2e01" -Description "Creates customer list and applies standard theme"