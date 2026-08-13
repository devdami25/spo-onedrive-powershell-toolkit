<#
.SYNOPSIS
Resets a user entry in the SharePoint User Information List across SharePoint Online sites (and optionally OneDrive).

.DESCRIPTION
This script loops through SharePoint Online site collections and removes/re-adds a specified user from the site
collection User Information List (hidden "User Information List").

This is commonly used when troubleshooting identity resolution issues (e.g., PUID/claim mismatch symptoms)
where a user appears in the site but behaves inconsistently.

.IMPORTANT
The account running this script ideally is a Site Collection Administrator on:
- all target site collections when running tenant-wide, or
- the specific target site when the scope is reduced (for example via -SiteFilter).

If it is NOT a Site Collection Administrator on a given site, the script will (unless
-DisableAutoElevate is set) automatically:
1. Detect the access-denied failure on that site.
2. Use the tenant admin connection to add the running account as a Site Collection
   Administrator on that specific site only (via Set-PnPTenantSite -Owners).
3. Retry and perform the User Information List reset.
4. Remove the running account as Site Collection Administrator from that site again
   (via Remove-PnPSiteCollectionAdmin), leaving all other admins untouched.

This self-elevate/de-elevate cycle only happens for sites where access was actually
missing - sites where the account is already a Site Collection Administrator are left
alone. It requires the account to be a SharePoint Administrator or Global Administrator
(i.e. able to call Set-PnPTenantSite from the tenant admin site).

Internally it uses:
- Remove-PnPUserInfo: Removes a user from the site collection User Information List.
- New-PnPUser: Adds a user to the built-in Site User Info List.
- Get-PnPTenantSite: Enumerates site collections; can include OneDrive sites via -IncludeOneDriveSites.
- Set-PnPTenantSite -Owners: Grants temporary Site Collection Administrator access (self-elevation).
- Remove-PnPSiteCollectionAdmin: Revokes the temporary access again (self-de-elevation).

.PARAMETER AdminUrl
Your SharePoint Admin Center URL, e.g. https://contoso-admin.sharepoint.com

.PARAMETER User
One or more user identifiers (UPN/email). This script matches by Email where possible and falls back to login name match.

.PARAMETER ClientId
ClientId required for Connect-PnPOnline -Interactive in this environment.

.PARAMETER IncludeOneDrive
If set, also processes OneDrive for Business (personal) site collections.

.PARAMETER SiteFilter
Optional. A simple filter to reduce scope (e.g. "Url -like '/sites/HR'").
Note: This uses the -Filter parameter of Get-PnPTenantSite.

.PARAMETER PassThru
If set, returns objects for each site processed (recommended). Otherwise writes host messages only.

.PARAMETER ReAdd
If set, re-adds the user to the User Information List after removal. Otherwise, only removes the user.

.PARAMETER AdminUpn
UPN of the account running the script (e.g. admin@contoso.onmicrosoft.com). Used only for the
self-elevate/de-elevate flow described above. If omitted, the script attempts to auto-detect it
from the current PnP access token; if that fails, auto-elevation is skipped and sites the account
cannot access will simply error out as before.

.PARAMETER DisableAutoElevate
If set, disables the self-elevate/de-elevate behavior entirely, restoring the original behavior
where a site the account cannot access simply fails with an error.

.EXAMPLE
.\Reset-UserInfoEntryAcrossTenant.ps1 `
  -AdminUrl "https://contoso-admin.sharepoint.com" `
  -User "user@tenant.onmicrosoft.com" `
  -ClientId "00000000-0000-0000-0000-000000000000" `
  -IncludeOneDrive `
  -ReAdd `
  -PassThru | Export-Csv .\reset-results.csv -NoTypeInformation

.EXAMPLE
# Dry run
.\Reset-UserInfoEntryAcrossTenant.ps1 `
  -AdminUrl "https://contoso-admin.sharepoint.com" `
  -User "user@tenant.onmicrosoft.com" `
  -ClientId "00000000-0000-0000-0000-000000000000" `
  -WhatIf `
  -PassThru

.EXAMPLE
# Explicit AdminUpn (skip auto-detection) with self-elevation for sites the admin lacks access to
.\Reset-UserInfoEntryAcrossTenant.ps1 `
  -AdminUrl "https://contoso-admin.sharepoint.com" `
  -User "user@tenant.onmicrosoft.com" `
  -ClientId "00000000-0000-0000-0000-000000000000" `
  -AdminUpn "admin@contoso.onmicrosoft.com" `
  -PassThru

.EXAMPLE
# Disable self-elevation, keep original strict behavior
.\Reset-UserInfoEntryAcrossTenant.ps1 `
  -AdminUrl "https://contoso-admin.sharepoint.com" `
  -User "user@tenant.onmicrosoft.com" `
  -ClientId "00000000-0000-0000-0000-000000000000" `
  -DisableAutoElevate `
  -PassThru

.NOTES
Requires: PnP.PowerShell (PowerShell 7+ recommended)
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

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string] $ClientId,

    [Parameter(Mandatory = $false)]
    [switch] $IncludeOneDrive,

    [Parameter(Mandatory = $false)]
    [string] $SiteFilter,

    [Parameter(Mandatory = $false)]
    [switch] $PassThru,

    [Parameter(Mandatory = $false)]
    [switch] $ReAdd,

    [Parameter(Mandatory = $false)]
    [string] $AdminUpn,

    [Parameter(Mandatory = $false)]
    [switch] $DisableAutoElevate
)

function Connect-PnPInteractiveSafe {
    param(
        [Parameter(Mandatory = $true)][string] $Url,
        [Parameter(Mandatory = $true)][string] $ClientId
    )

    Connect-PnPOnline -Url $Url -Interactive -ClientId $ClientId
}

function Test-PnPAccessDeniedError {
    param(
        [Parameter(Mandatory = $true)] $ErrorRecord
    )

    $message = $ErrorRecord.Exception.Message
    return ($message -match '(?i)access\s+is\s+denied|access\s+denied|unauthorized|forbidden|\b401\b|\b403\b|does not have permission|not have access')
}

function Get-PnPCurrentUserUpn {
    # Decodes the 'upn' (or fallback) claim out of the current PnP access token.
    # Best-effort only: returns $null if it cannot be determined, in which case
    # the caller should fall back to requiring -AdminUpn to be supplied explicitly.
    try {
        $token = Get-PnPAccessToken -ErrorAction Stop
        $payload = $token.Split('.')[1]
        $payload = $payload.Replace('-', '+').Replace('_', '/')
        switch ($payload.Length % 4) {
            2 { $payload += '==' }
            3 { $payload += '=' }
        }
        $claims = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($payload)) | ConvertFrom-Json

        if ($claims.upn) { return $claims.upn }
        if ($claims.preferred_username) { return $claims.preferred_username }
        if ($claims.unique_name) { return $claims.unique_name }
        return $null
    }
    catch {
        return $null
    }
}

function Grant-SiteCollectionAdminAccess {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true)][string] $SiteUrl,
        [Parameter(Mandatory = $true)][string] $AdminUpn,
        [Parameter(Mandatory = $true)][string] $AdminUrl,
        [Parameter(Mandatory = $true)][string] $ClientId
    )

    if ($PSCmdlet.ShouldProcess($SiteUrl, "Temporarily add '$AdminUpn' as Site Collection Administrator")) {
        Connect-PnPInteractiveSafe -Url $AdminUrl -ClientId $ClientId
        try {
            # Set-PnPTenantSite -Owners only appends; it does not remove existing admins.
            Set-PnPTenantSite -Url $SiteUrl -Owners $AdminUpn -ErrorAction Stop
        }
        finally {
            Disconnect-PnPOnline -ErrorAction SilentlyContinue | Out-Null
        }
    }
}

function Revoke-SiteCollectionAdminAccess {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true)][string] $SiteUrl,
        [Parameter(Mandatory = $true)][string] $AdminUpn,
        [Parameter(Mandatory = $true)][string] $ClientId
    )

    if ($PSCmdlet.ShouldProcess($SiteUrl, "Remove temporary Site Collection Administrator access for '$AdminUpn'")) {
        Connect-PnPInteractiveSafe -Url $SiteUrl -ClientId $ClientId
        try {
            # Removes only this account; all other site collection admins are left untouched.
            Remove-PnPSiteCollectionAdmin -Owners $AdminUpn -ErrorAction Stop
        }
        finally {
            Disconnect-PnPOnline -ErrorAction SilentlyContinue | Out-Null
        }
    }
}

function Reset-UserInfoEntry {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true)][string] $SiteUrl,
        [Parameter(Mandatory = $true)][string] $UserId,
        [Parameter(Mandatory = $false)][switch] $ReAdd,
        [Parameter(Mandatory = $true)][string] $AdminUrl,
        [Parameter(Mandatory = $false)][string] $AdminUpn,
        [Parameter(Mandatory = $false)][switch] $AutoElevate
    )

    $result = [PSCustomObject]@{
        SiteUrl      = $SiteUrl
        UserInput    = $UserId
        Found        = $false
        Removed      = $false
        ReAdded      = $false
        Elevated     = $false
        DeElevated   = $false
        Status       = "NotStarted"
        Error        = $null
        TimestampUtc = (Get-Date).ToUniversalTime().ToString("s") + "Z"
    }

    $elevated = $false

    try {
        Connect-PnPInteractiveSafe -Url $SiteUrl -ClientId $ClientId

        try {
            # Try to locate user by Email first, then fall back to loginname match
            $siteUsers = Get-PnPUser -ErrorAction Stop
        }
        catch {
            if ($AutoElevate -and $AdminUpn -and (Test-PnPAccessDeniedError -ErrorRecord $_)) {
                Write-Host "  Access denied on $SiteUrl - temporarily adding '$AdminUpn' as Site Collection Administrator." -ForegroundColor Yellow
                Disconnect-PnPOnline -ErrorAction SilentlyContinue | Out-Null

                Grant-SiteCollectionAdminAccess -SiteUrl $SiteUrl -AdminUpn $AdminUpn -AdminUrl $AdminUrl -ClientId $ClientId
                $elevated = $true
                $result.Elevated = $true

                Connect-PnPInteractiveSafe -Url $SiteUrl -ClientId $ClientId
                $siteUsers = Get-PnPUser -ErrorAction Stop
            }
            else {
                throw
            }
        }

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
            Remove-PnPUserInfo -Identity $match.LoginName -ErrorAction Stop
            $result.Removed = $true
        }

        if ($ReAdd -and $PSCmdlet.ShouldProcess($SiteUrl, "Re-add user '$UserId' to User Information List")) {
            New-PnPUser -LoginName $UserId -ErrorAction Stop | Out-Null
            $result.ReAdded = $true
        }

        $result.Status = "Success"
        return $result
    }
    catch {
        $result.Status = "Error"
        $result.Error  = $_.Exception.Message
        return $result
    }
    finally {
        Disconnect-PnPOnline -ErrorAction SilentlyContinue | Out-Null

        if ($elevated) {
            try {
                Revoke-SiteCollectionAdminAccess -SiteUrl $SiteUrl -AdminUpn $AdminUpn -ClientId $ClientId
                $result.DeElevated = $true
            }
            catch {
                $deElevationMessage = "Failed to remove temporary Site Collection Administrator access for '$AdminUpn' on $SiteUrl : $($_.Exception.Message)"
                Write-Warning $deElevationMessage
                $result.Error = if ($result.Error) { "$($result.Error) | $deElevationMessage" } else { $deElevationMessage }
            }
        }
    }
}

# ----------------------------
# Main
# ----------------------------
Write-Verbose "Connecting to admin center: $AdminUrl"
Connect-PnPInteractiveSafe -Url $AdminUrl -ClientId $ClientId

$AutoElevate = -not $DisableAutoElevate
if ($AutoElevate -and -not $AdminUpn) {
    $AdminUpn = Get-PnPCurrentUserUpn
    if ($AdminUpn) {
        Write-Verbose "Auto-detected signed-in admin UPN: $AdminUpn"
    }
    else {
        Write-Warning "Could not auto-detect the signed-in admin's UPN. Auto-elevation will be skipped for any site the account cannot access; pass -AdminUpn explicitly to enable it."
        $AutoElevate = $false
    }
}

# SharePoint sites (OneDrive excluded by default)
$tenantSites = if ($SiteFilter) {
    Get-PnPTenantSite -Filter $SiteFilter -ErrorAction Stop
}
else {
    Get-PnPTenantSite -ErrorAction Stop
}

# Hardcoded safety exclusions
$tenantSites = $tenantSites | Where-Object {
    $_.Template -ne "RedirectSite#0" -and
    $_.Url -notlike "*-my.sharepoint.com/personal*"
}

# OneDrive sites (optional)
$oneDriveSites = @()
if ($IncludeOneDrive) {
    $oneDriveSites = if ($SiteFilter) {
        Get-PnPTenantSite -IncludeOneDriveSites -Filter $SiteFilter -ErrorAction Stop |
            Where-Object { $_.Url -like "*-my.sharepoint.com/personal*" -and $_.Template -ne "RedirectSite#0" }
    }
    else {
        Get-PnPTenantSite -IncludeOneDriveSites -ErrorAction Stop |
            Where-Object { $_.Url -like "*-my.sharepoint.com/personal*" -and $_.Template -ne "RedirectSite#0" }
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

        $r = Reset-UserInfoEntry -SiteUrl $target.Url -UserId $u -ReAdd:$ReAdd `
            -AdminUrl $AdminUrl -AdminUpn $AdminUpn -AutoElevate:$AutoElevate

        # Add context fields
        $r | Add-Member -NotePropertyName SiteKind -NotePropertyValue $target.Kind -Force

        if ($r.Status -eq "Success") {
            $elevationNote = if ($r.Elevated) { if ($r.DeElevated) { " (temporary admin access granted and removed)" } else { " (temporary admin access granted - REMOVAL FAILED, see Error)" } } else { "" }
            Write-Host "Success: Removed/Re-added user entry.$elevationNote" -ForegroundColor Green
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