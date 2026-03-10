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

.PREREQUISITES
- The account running this script must have Site Collection Administrator privileges on the target site collection.
- PnP.PowerShell installed and available in the session.
- A ClientId is required for interactive authentication in this environment.

.PARAMETER SiteUrl
The URL of the SharePoint Online site (must be group-connected).

.PARAMETER ClientId
ClientId to use with Connect-PnPOnline -Interactive.

.PARAMETER AddAsSiteCollectionAdmin
If specified, also adds the M365 Group Owners principal as Site Collection Admin.

.EXAMPLE
.\Restore-M365GroupOwnersPrincipal.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/ProjectX" -ClientId "00000000-0000-0000-0000-000000000000" -AddAsSiteCollectionAdmin

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
    [switch] $AddAsSiteCollectionAdmin,

    [Parameter(Mandatory = $false)]
    [int] $MaxRetries = 3,

    [Parameter(Mandatory = $false)]
    [int] $BaseRetryDelaySeconds = 2,

    [Parameter(Mandatory = $false)]
    [switch] $PassThru
)

function Invoke-WithRetry {
    param(
        [Parameter(Mandatory = $true)][scriptblock] $ScriptBlock,
        [Parameter(Mandatory = $true)][string] $OperationName,
        [int] $maxRetries = $MaxRetries,
        [int] $baseDelaySeconds = $BaseRetryDelaySeconds
    )

    $attempt = 0
    while ($true) {
        try {
            return & $ScriptBlock
        }
        catch {
            $attempt++
            $msg = $_.Exception.Message
            $isThrottle = ($msg -match "429" -or $msg -match "throttl" -or $msg -match "Too Many Requests" -or $msg -match "503" -or $msg -match "temporarily unavailable" -or $msg -match "timeout")

            if (-not $isThrottle -or $attempt -gt $maxRetries) {
                throw
            }

            $delay = [Math]::Min(300, ($baseDelaySeconds * [Math]::Pow(2, ($attempt - 1))))
            Write-Warning "[$OperationName] transient failure (attempt $attempt/$maxRetries). Waiting $delay seconds then retrying... $msg"
            Start-Sleep -Seconds $delay
        }
    }
}

function Get-M365GroupOwnersClaim {
    param(
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()]
        [string] $GroupGuid
    )
    return "c:0o.c|federateddirectoryclaimprovider|{0}_o" -f $GroupGuid
}

try {
    $groupGuid = $null
    $m365GroupOwnersClaim = $null

    # Connect (ClientId is required in this environment)
    Invoke-WithRetry -OperationName "Connect-PnPOnline" -ScriptBlock {
        Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId
    }

    # Get the M365 Group Id tied to the site
    $site = Invoke-WithRetry -OperationName "Get-PnPSite" -ScriptBlock {
        Get-PnPSite -Includes RelatedGroupId
    }

    if (-not $site.RelatedGroupId -or $site.RelatedGroupId -eq [Guid]::Empty) {
        throw "This site does not appear to be Microsoft 365 Group-connected (RelatedGroupId is empty)."
    }

    $groupGuid = $site.RelatedGroupId.Guid.ToString()
    $m365GroupOwnersClaim = Get-M365GroupOwnersClaim -GroupGuid $groupGuid

    Write-Verbose "RelatedGroupId: $groupGuid"
    Write-Verbose "Owners claim : $m365GroupOwnersClaim"

    # Optionally add as Site Collection Admin
    if ($AddAsSiteCollectionAdmin) {
        if ($PSCmdlet.ShouldProcess($SiteUrl, "Add M365 Group Owners principal as Site Collection Admin")) {
            try {
                Invoke-WithRetry -OperationName "Add-PnPSiteCollectionAdmin" -ScriptBlock {
                    Add-PnPSiteCollectionAdmin -Owners $m365GroupOwnersClaim | Out-Null
                }
                Write-Host "Added as Site Collection Admin: $m365GroupOwnersClaim" -ForegroundColor Green
            }
            catch {
                Write-Warning "Could not add as Site Collection Admin (may already exist or cannot be added): $($_.Exception.Message)"
            }
        }
    }

    # Add to the site's associated Owners group
    $ownersGroup = Invoke-WithRetry -OperationName "Get-PnPGroup-AssociatedOwnerGroup" -ScriptBlock {
        Get-PnPGroup -AssociatedOwnerGroup
    }

    if (-not $ownersGroup) {
        throw "Could not resolve associated Owners group for site."
    }

    if ($PSCmdlet.ShouldProcess($ownersGroup.Title, "Add M365 Group Owners principal to Owners group")) {
        try {
            Invoke-WithRetry -OperationName "Add-PnPGroupMember" -ScriptBlock {
                Add-PnPGroupMember -Group $ownersGroup -LoginName $m365GroupOwnersClaim | Out-Null
            }
            Write-Host "Added to Owners group '$($ownersGroup.Title)': $m365GroupOwnersClaim" -ForegroundColor Green
        }
        catch {
            Write-Warning "Could not add to Owners group (may already exist or cannot be added): $($_.Exception.Message)"
        }
    }

    $result = [PSCustomObject]@{
        SiteUrl = $SiteUrl
        GroupGuid = $groupGuid
        OwnersClaim = $m365GroupOwnersClaim
        AddAsSiteCollectionAdmin = $AddAsSiteCollectionAdmin.IsPresent
        OwnersGroupTitle = $ownersGroup.Title
        Status = 'Success'
        Timestamp = (Get-Date)
    }

    if ($PassThru) { return $result }
}
catch {
    $errorMessage = $_.Exception.Message
    Write-Error $errorMessage

    if ($PassThru) {
        return [PSCustomObject]@{
            SiteUrl = $SiteUrl
            GroupGuid = $groupGuid
            OwnersClaim = $m365GroupOwnersClaim
            AddAsSiteCollectionAdmin = $AddAsSiteCollectionAdmin.IsPresent
            OwnersGroupTitle = $null
            Status = 'Failed'
            ErrorMessage = $errorMessage
            Timestamp = (Get-Date)
        }
    }
}
finally {
    Disconnect-PnPOnline -ErrorAction SilentlyContinue | Out-Null
}