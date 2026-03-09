<#
.SYNOPSIS
Resets a user entry in the SharePoint User Information List across SharePoint Online sites (and optionally OneDrive).

.DESCRIPTION
This script loops through SharePoint Online site collections and removes/re-adds a specified user from the site
collection User Information List (hidden "User Information List").

This is commonly used when troubleshooting identity resolution issues (e.g., PUID/claim mismatch symptoms)
where a user appears in the site but behaves inconsistently.

Internally it uses:
- Remove-PnPUser: Removes a user from the site collection User Information List. [1](https://pnp.github.io/powershell/cmdlets/Remove-PnPUser.html)
- New-PnPUser: Adds a user to the built-in Site User Info List. [2](https://pnp.github.io/powershell/cmdlets/New-PnPUser.html)
- Get-PnPTenantSite: Enumerates site collections; can include OneDrive sites via -IncludeOneDriveSites. [3](https://pnp.github.io/powershell/cmdlets/Get-PnPTenantSite.html)

.PARAMETER AdminUrl
Your SharePoint Admin Center URL, e.g. https://contoso-admin.sharepoint.com

.PARAMETER User
One or more user identifiers (UPN/email). This script matches by Email where possible and falls back to login name match.

.PARAMETER ClientId
ClientId for interactive auth

.PARAMETER IncludeOneDrive
If set, also processes OneDrive for Business (personal) site collections. [3](https://pnp.github.io/powershell/cmdlets/Get-PnPTenantSite.html)

.PARAMETER ExcludeRedirectSites
If set (default), excludes RedirectSite#0.

.PARAMETER ExcludePersonalSitesFromSPOList
If set (default), excludes *-my.sharepoint.com/personal* from the “SharePoint sites” loop.

.PARAMETER SiteFilter
Optional. A simple filter to reduce scope (e.g. "Url -like '/sites/HR'").
Note: This uses the -Filter parameter of Get-PnPTenantSite. [3](https://pnp.github.io/powershell/cmdlets/Get-PnPTenantSite.html)

.PARAMETER PassThru
If set, returns objects for each site processed (recommended). Otherwise writes host messages only.

.EXAMPLE
.\Reset-UserInfoEntryAcrossTenant.ps1 `
  -AdminUrl "https://contoso-admin.sharepoint.com" `
  -User "user@tenant.onmicrosoft.com" `
  -ClientId "00000000-0000-0000-0000-000000000000" `
  -IncludeOneDrive `
  -PassThru | Export-Csv .\reset-results.csv -NoTypeInformation

.EXAMPLE
# Dry run
.\Reset-UserInfoEntryAcrossTenant.ps1 `
  -AdminUrl "https://contoso-admin.sharepoint.com" `
  -User "user@tenant.onmicrosoft.com" `
  -WhatIf `
  -PassThru

.NOTES
Requires: PnP.PowerShell (PowerShell 7+ recommended)
PnP PowerShell is an open-source module with community support. [5](https://learn.microsoft.com/en-us/powershell/sharepoint/sharepoint-pnp/sharepoint-pnp-cmdlets)
Author: Dami Onabanjo
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string] $AdminUrl,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string[]] $User,

    [Parameter(Mandatory = $false)]
    [string] $ClientId,

    [Parameter(Mandatory = $false)]
    [switch] $IncludeOneDrive,

    [Parameter(Mandatory = $false)]
    [switch] $ExcludeRedirectSites = $true,

    [Parameter(Mandatory = $false)]
    [switch] $ExcludePersonalSitesFromSPOList = $true,

    [Parameter(Mandatory = $false)]
    [string] $SiteFilter,

    [Parameter(Mandatory = $false)]
    [switch] $PassThru
)

function Connect-PnPInteractiveSafe {
    param(
        [Parameter(Mandatory = $true)][string] $Url,
        [Parameter(Mandatory = $false)][string] $ClientId
    )

    if ($ClientId) {
        Connect-PnPOnline -Url $Url -Interactive -ClientId $ClientId
    } else {
        Connect-PnPOnline -Url $Url -Interactive
    }
}

function Reset-UserInfoEntry {
    param(
        [Parameter(Mandatory = $true)][string] $SiteUrl,
        [Parameter(Mandatory = $true)][string] $UserId
    )

    $result = [PSCustomObject]@{
        SiteUrl       = $SiteUrl
        UserInput     = $UserId
        Found         = $false
        Removed       = $false
        ReAdded       = $false
        Status        = "NotStarted"
        Error         = $null
        TimestampUtc  = (Get-Date).ToUniversalTime().ToString("s") + "Z"
    }

    try {
        Connect-PnPInteractiveSafe -Url $SiteUrl -ClientId $ClientId

        # Try to locate user by Email first, then fall back to loginname match
        $siteUsers = Get-PnPUser -ErrorAction Stop
        $match = $siteUsers | Where-Object {
            ($_.Email -and $_.Email -ieq $UserId) -or
            ($_.LoginName -and $_.LoginName -ilike "*$UserId*")
        } | Select-Object -First 1

        if (-not $match) {
            $result.Status = "UserNotFound"
            return $result
        }

        $result.Found = $true

        if ($PSCmdlet.ShouldProcess($SiteUrl, "Remove user '$($match.LoginName)' from User Information List")) {
            # Removes a user from the site collection User Information List [1](https://pnp.github.io/powershell/cmdlets/Remove-PnPUser.html)
            Remove-PnPUser -Identity $match.LoginName -Force -ErrorAction Stop
            $result.Removed = $true
        }

        if ($PSCmdlet.ShouldProcess($SiteUrl, "Re-add user '$UserId' to User Information List")) {
            # Adds a user to the built-in Site User Info List [2](https://pnp.github.io/powershell/cmdlets/New-PnPUser.html)
            New-PnPUser -LoginName $UserId -ErrorAction Stop | Out-Null
            $result.ReAdded = $true
        }

        $result.Status = "Success"
        return $result
    }
    catch {
        $result.Status = "Error"
        $result.Error = $_.Exception.Message
        return $result
    }
    finally {
        Disconnect-PnPOnline -ErrorAction SilentlyContinue | Out-Null
    }
}

# ----------------------------
# Main
# ----------------------------
Write-Verbose "Connecting to admin center: $AdminUrl"
Connect-PnPInteractiveSafe -Url $AdminUrl -ClientId $ClientId

# Get SPO sites (excluding OneDrive by default)
# Get-PnPTenantSite returns all sites (excluding OneDrive by default) and supports -IncludeOneDriveSites and -Filter [3](https://pnp.github.io/powershell/cmdlets/Get-PnPTenantSite.html)
$tenantSites = if ($SiteFilter) {
    Get-PnPTenantSite -Filter $SiteFilter -ErrorAction Stop
} else {
    Get-PnPTenantSite -ErrorAction Stop
}

if ($ExcludeRedirectSites) {
    $tenantSites = $tenantSites | Where-Object { $_.Template -ne "RedirectSite#0" }
}

if ($ExcludePersonalSitesFromSPOList) {
    $tenantSites = $tenantSites | Where-Object { $_.Url -notlike "*-my.sharepoint.com/personal*" }
}

# OneDrive sites (optional)
$oneDriveSites = @()
if ($IncludeOneDrive) {
    $oneDriveSites = if ($SiteFilter) {
        Get-PnPTenantSite -IncludeOneDriveSites -Filter $SiteFilter -ErrorAction Stop |
            Where-Object { $_.Url -like "*-my.sharepoint.com/personal*" }
    } else {
        Get-PnPTenantSite -IncludeOneDriveSites -ErrorAction Stop |
            Where-Object { $_.Url -like "*-my.sharepoint.com/personal*" }
    }
}

Disconnect-PnPOnline -ErrorAction SilentlyContinue | Out-Null

$allTargets = @(
    $tenantSites | ForEach-Object { [PSCustomObject]@{ Url = $_.Url; Kind = "SharePoint" } }
    $oneDriveSites | ForEach-Object { [PSCustomObject]@{ Url = $_.Url; Kind = "OneDrive" } }
)

$total = $allTargets.Count
$counter = 0

$results = New-Object System.Collections.Generic.List[object]

foreach ($target in $allTargets) {
    $counter++
    Write-Progress -Activity "Resetting User Info Entries" -Status "$counter / $total : $($target.Kind) : $($target.Url)" -PercentComplete (($counter / $total) * 100)

    foreach ($u in $User) {
        Write-Host "`n[$($target.Kind)] Processing site: $($target.Url) | User: $u" -ForegroundColor Cyan

        $r = Reset-UserInfoEntry -SiteUrl $target.Url -UserId $u

        # Add context fields
        $r | Add-Member -NotePropertyName SiteKind -NotePropertyValue $target.Kind -Force

        if ($r.Status -eq "Success") {
            Write-Host "Success: Removed/Re-added user entry." -ForegroundColor Green
        }
        elseif ($r.Status -eq "UserNotFound") {
            Write-Host "User not found in site." -ForegroundColor DarkGray
        }
        else {
            Write-Host "Error: $($r.Error)" -ForegroundColor Red
        }

        $results.Add($r) | Out-Null
    }
}

Write-Progress -Activity "Resetting User Info Entries" -Completed

if ($PassThru) {
    $results
}
