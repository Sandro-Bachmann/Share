# SharePoint Online Permissions Audit Report - Remodeled with Certificate Authentication

#region ***Parameters***
# Example usage:
$SiteURL="https://domain.sharepoint.com/sites/xy/"
$ReportFile="C:\filepath\site-SitePermission.csv"
#endregion

# --- Azure AD (Entra ID) App Registration Details ---
# 👈 IMPORTANT: Replace these placeholders with YOUR actual values from your App Registration
$ClientID = "CLIENTID"
$TenantID = "TENANTID"
$CertificateThumbprint = "CERTIFICATE_THUMBPRINT"

# --- Global Variable for Collected Permissions ---
# This array will store all permission entries before a single export to CSV.
# Explicitly declare as ArrayList to ensure it remains a mutable collection.
[System.Collections.ArrayList]$GlobalPermissionCollection = @()

# Function to Get Permissions Applied on a particular Object (Web, List, or Folder)
Function Get-PnPPermissions([Microsoft.SharePoint.Client.SecurableObject]$Object, [Microsoft.SharePoint.Client.Web]$CurrentWeb)
{
    # Determine the type of the object and its properties
    $ObjectType = ""
    $ObjectURL = ""
    $ObjectTitle = ""

    Switch($Object.TypedObject.ToString())
    {
        "Microsoft.SharePoint.Client.Web"
        {
            $ObjectType = "Site"
            $ObjectURL = $Object.Url
            $ObjectTitle = $Object.Title
        }
        "Microsoft.SharePoint.Client.ListItem"
        {
            $ObjectType = "Folder"
            # Get the URL of the Folder
            # Ensure the folder object is loaded with its properties
            $Folder = Get-PnPProperty -ClientObject $Object -Property Folder
            $ObjectTitle = $Object.Folder.Name
            # Construct the absolute URL for the folder
            # A more robust way to combine URLs, handling potential trailing slashes
            $ObjectURL = [System.Uri]::new($CurrentWeb.Url, $Object.Folder.ServerRelativeUrl).AbsoluteUri
        }
        Default # Covers Lists, Document Libraries, etc.
        {
            $ObjectType = $Object.BaseType # List, DocumentLibrary, etc.
            $ObjectTitle = $Object.Title
            # Get the URL of the List or Library
            $RootFolder = Get-PnPProperty -ClientObject $Object -Property RootFolder
            # Construct the absolute URL for the list/library
            $ObjectURL = [System.Uri]::new($CurrentWeb.Url, $RootFolder.ServerRelativeUrl).AbsoluteUri
        }
    }

    # Load RoleAssignments and HasUniqueRoleAssignments properties efficiently
    # Only load RoleAssignments and HasUniqueRoleAssignments directly.
    # Member and RoleDefinitionBindings are properties of individual RoleAssignment objects,
    # which will be accessed in the foreach loop.
    Get-PnPProperty -ClientObject $Object -Property HasUniqueRoleAssignments, RoleAssignments

    # Check if Object has unique permissions
    $HasUniquePermissions = $Object.HasUniqueRoleAssignments

    # Loop through each permission assigned and extract details
    Foreach($RoleAssignment in $Object.RoleAssignments)
    {
        # Ensure Member and RoleDefinitionBindings are loaded for the current RoleAssignment
        Get-PnPProperty -ClientObject $RoleAssignment -Property Member, RoleDefinitionBindings

        # Get the Principal Type: User, SP Group, AD Group
        $PermissionType = $RoleAssignment.Member.PrincipalType

        # Get the Permission Levels assigned
        $PermissionLevels = $RoleAssignment.RoleDefinitionBindings | Select -ExpandProperty Name

        # Remove "Limited Access" as it's often an artifact of sharing and not a direct permission level for auditing
        $PermissionLevels = ($PermissionLevels | Where { $_ -ne "Limited Access"}) -join "; "

        # Skip principals with no effective permissions assigned (after filtering "Limited Access")
        If($PermissionLevels.Length -eq 0) {
            Write-Verbose "Skipping $($RoleAssignment.Member.LoginName) on $($ObjectURL) as no effective permissions found."
            Continue
        }

        # Initialize common properties for the permission entry
        $PermissionsEntry = [PSCustomObject]@{
            Object              = $ObjectType
            Title               = $ObjectTitle
            URL                 = $ObjectURL
            HasUniquePermissions= $HasUniquePermissions
            Users               = ""
            Type                = $PermissionType
            Permissions         = $PermissionLevels
            GrantedThrough      = ""
        }

        # Check if the Principal is a SharePoint group
        If($PermissionType -eq "SharePointGroup")
        {
            Write-Verbose "Processing SharePoint Group: $($RoleAssignment.Member.LoginName)"
            # Get Group Members
            # Use -ErrorAction SilentlyContinue to avoid breaking on groups with no members or access issues
            $GroupMembers = Get-PnPGroupMember -Identity $RoleAssignment.Member.LoginName -ErrorAction SilentlyContinue

            # If no members or only "System Account", skip
            If($GroupMembers.count -eq 0 -or ($GroupMembers | Where { $_.Title -ne "System Account"}).Count -eq 0) {
                Write-Verbose "Skipping empty or system-only SharePoint Group: $($RoleAssignment.Member.LoginName)"
                Continue
            }
            $GroupUsers = ($GroupMembers | Select -ExpandProperty Title | Where { $_ -ne "System Account"}) -join "; "

            # Add group-specific data to the object
            $PermissionsEntry.Users = $GroupUsers
            $PermissionsEntry.GrantedThrough = "SharePoint Group: $($RoleAssignment.Member.LoginName)"
        }
        Else # User or Security Group (AD Group)
        {
            Write-Verbose "Processing User/AD Group: $($RoleAssignment.Member.Title)"
            # Add user/AD group specific data to the object
            $PermissionsEntry.Users = $RoleAssignment.Member.Title
            $PermissionsEntry.GrantedThrough = "Direct Permissions"
        }

        # Add the constructed permission entry to the global collection using .Add() method
        $GlobalPermissionCollection.Add($PermissionsEntry) | Out-Null
    }
}

# Function to get SharePoint Online site permissions report
Function Generate-PnPSitePermissionRpt()
{
    [cmdletbinding(DefaultParameterSetName='Default', SupportsShouldProcess=$true, ConfirmImpact='High')]
    Param
    (
        [Parameter(Mandatory=$true, HelpMessage="The URL of the SharePoint site to audit.")]
        [String] $SiteURL,
        [Parameter(Mandatory=$true, HelpMessage="The full path to the output CSV file.")]
        [String] $ReportFile,
        [Parameter(Mandatory=$false, HelpMessage="Scan sub-sites recursively.")]
        [switch] $Recursive,
        [Parameter(Mandatory=$false, HelpMessage="Scan folders within lists and libraries.")]
        [switch] $ScanFolders,
        [Parameter(Mandatory=$false, HelpMessage="Include inherited permissions in the report.")]
        [switch] $IncludeInheritedPermissions
    )

    # Validate parameters
    If (-not ($SiteURL -match "^https?://")) {
        Write-Error "Invalid SiteURL. Please provide a valid URL starting with http:// or https://."
        Return
    }
    Try {
        # Connect to the Site using certificate authentication
        Write-Host -ForegroundColor Cyan "🔗 Connecting to SharePoint site: $($SiteURL)..."
        Connect-PnPOnline -URL $SiteURL -Thumbprint $CertificateThumbprint -Tenant $TenantID -ClientId $ClientID -ErrorAction Stop

        # Get the current web context
        $Web = Get-PnPWeb -Includes Url, Title, ServerRelativeUrl, HasUniqueRoleAssignments, Webs, Lists

        Write-Host -ForegroundColor Yellow "👥 Getting Site Collection Administrators for $($Web.Title)..."
        # Get Site Collection Administrators
        $SiteAdmins = Get-PnPSiteCollectionAdmin -ErrorAction SilentlyContinue

        $SiteCollectionAdmins = ($SiteAdmins | Select -ExpandProperty Title) -join "; "

        # Add Site Collection Administrators to the global collection using .Add() method
        $GlobalPermissionCollection.Add([PSCustomObject]@{
            Object              = "Site Collection"
            Title               = $Web.Title
            URL                 = $Web.URL
            HasUniquePermissions= "TRUE" # Site collection admins always have direct control
            Users               = $SiteCollectionAdmins
            Type                = "Site Collection Administrators"
            Permissions         = "Site Owner"
            GrantedThrough      = "Direct Permissions"
        }) | Out-Null

        # Function to Get Permissions of Folders in a given List
        Function Get-PnPFolderPermission([Microsoft.SharePoint.Client.List]$List, [Microsoft.SharePoint.Client.Web]$CurrentWeb)
        {
            Write-Host -ForegroundColor Yellow "`t `t 📁 Getting Permissions of Folders in the List: $($List.Title)"
            # Get All Folders from List
            # Use -Recursive to get all folders in subfolders as well
            $Folders = Get-PnPListItem -List $List -PageSize 2000 -Query "<View Scope='RecursiveAll'><Query><Where><Eq><FieldRef Name='FSObjType'/><Value Type='Integer'>1</Value></Eq></Where></Query></View>" | Where-Object {
                ($_.FieldValues.FileLeafRef -ne "Forms") -and (-Not($_.FieldValues.FileLeafRef.StartsWith("_")))
            }

            $ItemCounter = 0
            # Loop through each Folder
            ForEach($Folder in $Folders)
            {
                $ItemCounter++
                Write-Progress -PercentComplete ($ItemCounter / ($Folders.Count) * 100) -Activity "Getting Permissions of Folders in List '$($List.Title)'" -Status "Processing Folder '$($Folder.FieldValues.FileLeafRef)' at '$($Folder.FieldValues.FileRef)' ($ItemCounter of $($Folders.Count))" -Id 2 -ParentId 1

                # Get Objects with Unique Permissions or Inherited Permissions based on 'IncludeInheritedPermissions' switch
                If($IncludeInheritedPermissions)
                {
                    Get-PnPPermissions -Object $Folder -CurrentWeb $CurrentWeb
                }
                Else
                {
                    # Check if Folder has unique permissions
                    $HasUniquePermissions = Get-PnPProperty -ClientObject $Folder -Property HasUniqueRoleAssignments
                    If($HasUniquePermissions -eq $True)
                    {
                        # Call the function to generate Permission report
                        Get-PnPPermissions -Object $Folder -CurrentWeb $CurrentWeb
                    }
                }
            }
        }

        # Function to Get Permissions of all lists from the given web
        Function Get-PnPListPermission([Microsoft.SharePoint.Client.Web]$Web)
        {
            # Get All Lists from the web
            $Lists = Get-PnPProperty -ClientObject $Web -Property Lists

            # Exclude common system lists to avoid noise
            $ExcludedLists = @(
                "Access Requests","App Packages","appdata","appfiles","Apps in Testing","Cache Profiles","Composed Looks",
                "Content and Structure Reports","Content type publishing error log","Converted Forms","Device Channels",
                "Form Templates","fpdatasources","Get started with Apps for Office and SharePoint","List Template Gallery",
                "Long Running Operation Status","Maintenance Log Library","Images","site collection images","Master Docs",
                "Master Page Gallery","MicroFeed","NintexFormXml","Quick Deploy Items","Relationships List",
                "Reusable Content","Reporting Metadata","Reporting Templates","Search Config List","Site Assets",
                "Preservation Hold Library","Site Pages","Solution Gallery","Style Library",
                "Suggested Content Browser Locations","Theme Gallery","TaxonomyHiddenList","User Information List",
                "Web Part Gallery","wfpub","wfsvc","Workflow History","Workflow Tasks","Pages"
            )

            $Counter = 0
            # Get all lists from the web
            ForEach($List in $Lists)
            {
                # Exclude System Lists and hidden lists
                If($List.Hidden -eq $False -and $ExcludedLists -notcontains $List.Title)
                {
                    $Counter++
                    Write-Progress -PercentComplete ($Counter / ($Lists.Count) * 100) -Activity "Exporting Permissions from List '$($List.Title)' in $($Web.URL)" -Status "Processing Lists $Counter of $($Lists.Count)" -Id 1

                    # Get Item Level Permissions if 'ScanFolders' switch present
                    If($ScanFolders)
                    {
                        # Get Folder Permissions
                        Get-PnPFolderPermission -List $List -CurrentWeb $Web
                    }

                    # Get Lists with Unique Permissions or Inherited Permissions based on 'IncludeInheritedPermissions' switch
                    If($IncludeInheritedPermissions)
                    {
                        Get-PnPPermissions -Object $List -CurrentWeb $Web
                    }
                    Else
                    {
                        # Check if List has unique permissions
                        $HasUniquePermissions = Get-PnPProperty -ClientObject $List -Property HasUniqueRoleAssignments
                        If($HasUniquePermissions -eq $True)
                        {
                            # Call the function to check permissions
                            Get-PnPPermissions -Object $List -CurrentWeb $Web
                        }
                    }
                }
            }
        }

        # Function to Get Webs's Permissions from given URL
        Function Get-PnPWebPermission([Microsoft.SharePoint.Client.Web]$Web)
        {
            # Call the function to Get permissions of the web
            Write-Host -ForegroundColor Yellow "🌐 Getting Permissions of the Web: $($Web.URL)..."
            Get-PnPPermissions -Object $Web -CurrentWeb $Web

            # Get List Permissions
            Write-Host -ForegroundColor Yellow "`t 📚 Getting Permissions of Lists and Libraries..."
            Get-PnPListPermission($Web)

            # Recursively get permissions from all sub-webs based on the "Recursive" Switch
            If($Recursive)
            {
                # Get Subwebs of the Web
                # Ensure 'Webs' property is loaded for the current web
                $Subwebs = Get-PnPProperty -ClientObject $Web -Property Webs

                # Iterate through each subsite in the current web
                ForEach ($Subweb in $Web.Webs)
                {
                    # Connect to the subweb to get its context for further operations
                    # This is important as PnP cmdlets operate on the current connected context
                    Connect-PnPOnline -URL $Subweb.Url -Thumbprint $CertificateThumbprint -Tenant $TenantID -ClientId $ClientID -ErrorAction Stop

                    # Get Webs with Unique Permissions or Inherited Permissions based on 'IncludeInheritedPermissions' switch
                    If($IncludeInheritedPermissions)
                    {
                        Get-PnPWebPermission($Subweb)
                    }
                    Else
                    {
                        # Check if the Web has unique permissions
                        $HasUniquePermissions = Get-PnPProperty -ClientObject $SubWeb -Property HasUniqueRoleAssignments

                        # Get the Web's Permissions
                        If($HasUniquePermissions -eq $true)
                        {
                            # Call the function recursively
                            Get-PnPWebPermission($Subweb)
                        }
                    }
                    # Disconnect from the subweb to avoid lingering connections, though PnP handles context switching well.
                    # Disconnect-PnPOnline # Not strictly necessary here as Connect-PnPOnline manages context, but good for explicit cleanup if not recursing further.
                }
            }
        }

        # Call the function with RootWeb to get site collection permissions
        Get-PnPWebPermission $Web

        # --- Final Export to CSV ---
        Write-Host -ForegroundColor Green "`n✨ All permissions collected. Exporting to CSV..."
        $GlobalPermissionCollection | Export-CSV $ReportFile -NoTypeInformation -Force

        Write-Host -ForegroundColor Green "`n🎉*** Site Permission Report Generated Successfully! Path: $($ReportFile)***"
    }
    Catch {
        Write-Host -ForegroundColor Red "❌ Error Generating Site Permission Report! Details: $($_.Exception.Message)"
        Write-Host -ForegroundColor Red "Error Line Number: $($_.InvocationInfo.ScriptLineNumber)"
        Write-Host -ForegroundColor Red "Error Script Name: $($_.InvocationInfo.ScriptName)"
        Write-Error $_.Exception.Message -ErrorAction Continue
    }
    Finally {
        # Always disconnect from SharePoint Online for a clean exit
        Write-Host -ForegroundColor Cyan "👋 Disconnecting from SharePoint Online..."
        Disconnect-PnPOnline
    }
}


# Call the function to generate permission report
# Example 1: Scan a single site and its folders for unique permissions
Generate-PnPSitePermissionRpt -SiteURL $SiteURL -ReportFile $ReportFile -ScanFolders

# Example 2: Scan recursively, including folders and inherited permissions
# Generate-PnPSitePermissionRpt -SiteURL $SiteURL -ReportFile $ReportFile -Recursive -ScanFolders -IncludeInheritedPermissions