<#
.SYNOPSIS
Restores the hidden Microsoft 365 Group Owners principal to a group-connected SharePoint site's Owners group.

.DESCRIPTION
On group-connected (Microsoft 365 Group / Teams-connected) SharePoint sites, the SharePoint "Associated Owners" group
normally contains a hidden principal that represents the Microsoft 365 Group Owners.

If this principal is removed, adding it back through the SharePoint UI can be unreliable and may add the Members principal instead.
This script rebuilds the correct claims login name for the M365 Group Owners principal and adds it back to:
  1) Site Collection Administrators (optional)
  2) The site's Associated Owners group

The Microsoft 365 Group claims are commonly represented as:
  - Members: c:0o.c|federateddirectoryclaimprovider|{GroupGuid}
  - Owners : c:0o.c|federateddirectoryclaimprovider|{GroupGuid}_o

.PARAMETER SiteUrl
The URL of the SharePoint Online site (must be group-connected).

.PARAMETER ClientId
ClientId to use with Connect-PnPOnline -Interactive for tenants that require it.

.PARAMETER AddAsSiteCollectionAdmin
If specified, also adds the M365 Group Owners principal as Site Collection Admin.

.EXAMPLE
.\Restore-M365GroupOwnersPrincipal.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/ProjectX" -AddAsSiteCollectionAdmin

.EXAMPLE
.\Restore-M365GroupOwnersPrincipal.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/ProjectX" -ClientId "00000000-0000-0000-0000-000000000000"

.NOTES
Requires: PnP.PowerShell
Author: Dami Onabanjo
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string] $SiteUrl,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string] $ClientId,

    [Parameter(Mandatory = $false)]
    [switch] $AddAsSiteCollectionAdmin
)

try {
    # Connect
    if ($ClientId) {
        Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId
    }
    else {
        Connect-PnPOnline -Url $SiteUrl -Interactive
    }

    # Get the M365 Group Id tied to the site
    $site = Get-PnPSite -Includes RelatedGroupId
    if (-not $site.RelatedGroupId -or $site.RelatedGroupId -eq [Guid]::Empty) {
        throw "This site does not appear to be Microsoft 365 Group-connected (RelatedGroupId is empty)."
    }

    $groupGuid = $site.RelatedGroupId.Guid.ToString()
    $m365GroupOwnersClaim = "c:0o.c|federateddirectoryclaimprovider|{0}_o" -f $groupGuid

    Write-Verbose "RelatedGroupId: $groupGuid"
    Write-Verbose "Owners claim : $m365GroupOwnersClaim"

    # Optionally add as Site Collection Admin
    if ($AddAsSiteCollectionAdmin) {
        if ($PSCmdlet.ShouldProcess($SiteUrl, "Add M365 Group Owners principal as Site Collection Admin")) {
            try {
                Add-PnPSiteCollectionAdmin -Owners $m365GroupOwnersClaim | Out-Null
                Write-Host "Added as Site Collection Admin: $m365GroupOwnersClaim" -ForegroundColor Green
            }
            catch {
                Write-Warning "Could not add as Site Collection Admin (may already exist): $($_.Exception.Message)"
            }
        }
    }

    # Add to the site's associated Owners group
    $ownersGroup = Get-PnPGroup -AssociatedOwnerGroup

    if ($PSCmdlet.ShouldProcess($ownersGroup.Title, "Add M365 Group Owners principal to Owners group")) {
        try {
            Add-PnPGroupMember -Group $ownersGroup -LoginName $m365GroupOwnersClaim | Out-Null
            Write-Host "Added to Owners group '$($ownersGroup.Title)': $m365GroupOwnersClaim" -ForegroundColor Green
        }
        catch {
            Write-Warning "Could not add to Owners group (may already exist): $($_.Exception.Message)"
        }
    }
}
catch {
    Write-Error $_.Exception.Message
}
finally {
    Disconnect-PnPOnline -ErrorAction SilentlyContinue | Out-Null
}
