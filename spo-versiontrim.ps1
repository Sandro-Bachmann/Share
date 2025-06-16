# SPO Cleanup, set all sites to automatic trim and remove not used versions
# Connect to SharePoint Online Admin Center
# Make sure you have installed the SharePoint Online Management Shell
# Install-Module -Name Microsoft.Online.SharePoint.PowerShell -Force

# Variables
$AdminCenterURL = "https://YOURM365DOMAIN-admin.sharepoint.com"

# Connect to SharePoint Online
Connect-SPOService -Url $AdminCenterURL

# Activate default versioning to new sites / teams to automatic
Set-SPOTenant-EnableAutoExpirationVersionTrim $true

# Get all site collections
$sites = Get-SPOSite -Limit All

foreach ($site in $sites) {
    Write-Host "Processing site: $($site.Url)"
    
    # Set versioning to automatic
    Set-SPOSite -Identity $site.Url -EnableAutoExpirationVersionTrim $true -Confirm:$false
    
    # Enable automatic site collection trimming cleanup
    New-SPOSiteFileVersionBatchDeleteJob -Identity $site.Url -Automatic -Confirm:$false

    Write-Host "Updated site: $($site.Url)" -foregroundcolor Green
}

# Disconnect from SharePoint Online
Disconnect-SPOService