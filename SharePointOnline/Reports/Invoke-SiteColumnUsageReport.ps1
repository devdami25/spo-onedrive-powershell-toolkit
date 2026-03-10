<#
.SYNOPSIS
Reports (and optionally remediates) usage of a specified site column across content types and lists/libraries.

.DESCRIPTION
This script checks where a given site column (internal name) is used within:
- Content Types in the site
- Lists/Libraries in the site

By default it runs against ONE site (SiteUrl).
Optionally, it can run across ALL SharePoint sites in the tenant (excluding OneDrive and Redirect sites) and supports filtering scope.

IMPORTANT:
The account running this script must be a Site Collection Administrator on:
- the target site (single-site mode), OR
- all target sites (tenant-wide mode).

Internally it uses:
- Get-PnPContentType to enumerate content types [1](https://pnp.github.io/powershell/cmdlets/Get-PnPContentType.html)
- Get-PnPTenantSite for tenant-wide enumeration [2](https://deepwiki.com/pnp/powershell/6-microsoft-365-groups-management)
- Get-PnPList to enumerate lists/libraries [3](https://pnp.github.io/powershell/cmdlets/Get-PnPList.html)
- Remove-PnPFieldFromContentType to remove a site column from a content type [4](https://pnp.github.io/powershell/cmdlets/Remove-PnPFieldFromContentType.html)
- Remove-PnPField to remove a field from list or site [5](https://pnp.github.io/powershell/cmdlets/Remove-PnPField.html)

ClientId is required for PnP interactive connections in this environment.

.PARAMETER SiteUrl
Target site URL. Used in single-site mode (default).

.PARAMETER SiteColumnInternalName
Internal name of the site column to report on (e.g. "CandCTax1").

.PARAMETER ClientId
Required. ClientId used with Connect-PnPOnline -Interactive.

.PARAMETER TenantWide
If specified, runs across all SharePoint site collections (excludes OneDrive and Redirect sites by default).

.PARAMETER AdminUrl
Required when -TenantWide is used. Example: https://contoso-admin.sharepoint.com

.PARAMETER SiteFilter
Optional. Reduces scope in tenant-wide mode. Uses the -Filter parameter of Get-PnPTenantSite.
Example: "Url -like '/sites/HR'"

.PARAMETER Remediate
If specified, the script will remove the column from:
- any content types where it is found
- any lists/libraries where it is found
- and then attempt to remove the site column itself (best-effort)
A report is always produced regardless.

.PARAMETER ExportPath
CSV output path. If not provided, a timestamped CSV is created in the current folder.

.PARAMETER PassThru
Returns the report objects to the pipeline (useful for additional processing).

.EXAMPLE
# Report only (single site - default)
.\Invoke-SiteColumnUsageReport.ps1 `
  -SiteUrl "https://contoso.sharepoint.com/sites/ProjectX" `
  -SiteColumnInternalName "CandCTax1" `
  -ClientId "00000000-0000-0000-0000-000000000000"

.EXAMPLE
# Remediate + export (single site)
.\Invoke-SiteColumnUsageReport.ps1 `
  -SiteUrl "https://contoso.sharepoint.com/sites/ProjectX" `
  -SiteColumnInternalName "CandCTax1" `
  -ClientId "00000000-0000-0000-0000-000000000000" `
  -Remediate `
  -ExportPath ".\CandCTax1-usage.csv"

.EXAMPLE
# Tenant-wide report with scope filter
.\Invoke-SiteColumnUsageReport.ps1 `
  -TenantWide `
  -AdminUrl "https://contoso-admin.sharepoint.com" `
  -SiteFilter "Url -like '/sites/HR'" `
  -SiteColumnInternalName "CandCTax1" `
  -ClientId "00000000-0000-0000-0000-000000000000" `
  -ExportPath ".\CandCTax1-tenant-usage.csv"

.NOTES
Requires: PnP.PowerShell (PowerShell 7+ recommended)
Author: Dami Onabanjo
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string] $SiteUrl,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string] $SiteColumnInternalName,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string] $ClientId,

    [Parameter(Mandatory = $false)]
    [switch] $TenantWide,

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string] $AdminUrl,

    [Parameter(Mandatory = $false)]
    [string] $SiteFilter,

    [Parameter(Mandatory = $false)]
    [switch] $Remediate,

    [Parameter(Mandatory = $false)]
    [string] $ExportPath,

    [Parameter(Mandatory = $false)]
    [switch] $PassThru
)

function Connect-PnPInteractiveRequired {
    param(
        [Parameter(Mandatory = $true)][string] $Url
    )
    Connect-PnPOnline -Url $Url -Interactive -ClientId $ClientId -ReturnConnection
}

function Get-DefaultExportPath {
    $ts = (Get-Date).ToString("yyyyMMdd-HHmmss")
    return (Join-Path -Path (Get-Location) -ChildPath ("SiteColumnUsageReport-{0}.csv" -f $ts))
}

function Ensure-FolderExists {
    param([Parameter(Mandatory = $true)][string] $Path)
    $dir = Split-Path -Path $Path -Parent
    if ($dir -and -not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
}

function Process-Site {
    param(
        [Parameter(Mandatory = $true)][string] $TargetSiteUrl
    )

    $results = New-Object System.Collections.Generic.List[object]

    Write-Host "`nProcessing site: $TargetSiteUrl" -ForegroundColor Cyan
    $conn = Connect-PnPInteractiveRequired -Url $TargetSiteUrl

    # ---- Content Types (ALL) ----
    Write-Host -BackgroundColor Blue "Checking Content Types"

    # Get-PnPContentType can retrieve content types from the current web [1](https://pnp.github.io/powershell/cmdlets/Get-PnPContentType.html)
    $cts = Get-PnPContentType -Connection $conn

    foreach ($ct in $cts) {
        # Fields property is lazily loaded; request it explicitly
        $ctFields = Get-PnPProperty -ClientObject $ct -Property "Fields"
        $matchedField = $ctFields | Where-Object { $_.InternalName -eq $SiteColumnInternalName } | Select-Object -First 1

        if ($matchedField) {
            Write-Host -ForegroundColor Green "Found column '$SiteColumnInternalName' in Content Type: $($ct.Name)"

            $action = if ($Remediate) { "RemoveFromContentType" } else { "ReportOnly" }
            $actionResult = "N/A"

            if ($Remediate -and $PSCmdlet.ShouldProcess($TargetSiteUrl, "Remove '$SiteColumnInternalName' from content type '$($ct.Name)'")) {
                try {
                    # Removes a site column from a content type [4](https://pnp.github.io/powershell/cmdlets/Remove-PnPFieldFromContentType.html)
                    Remove-PnPFieldFromContentType -Field $matchedField -ContentType $ct -Connection $conn
                    $actionResult = "Removed"
                    Write-Host -ForegroundColor Yellow "Removed from Content Type: $($ct.Name)"
                }
                catch {
                    $actionResult = "Failed: $($_.Exception.Message)"
                    Write-Warning "Failed removing from Content Type '$($ct.Name)': $($_.Exception.Message)"
                }
            }

            $results.Add([PSCustomObject]@{
                SiteUrl            = $TargetSiteUrl
                LocationType       = "ContentType"
                LocationName       = $ct.Name
                FieldInternalName  = $SiteColumnInternalName
                Found              = $true
                Action             = $action
                ActionResult       = $actionResult
            }) | Out-Null
        }
    }

    # ---- Lists/Libraries ----
    Write-Host -BackgroundColor Blue "Checking Lists/Libraries"

    # Get-PnPList returns lists in the current web [3](https://pnp.github.io/powershell/cmdlets/Get-PnPList.html)
    $lists = Get-PnPList -Connection $conn | Where-Object { $_.Hidden -ne $true -and $_.IsSystemList -ne $true }

    foreach ($list in $lists) {
        $field = Get-PnPField -List $list -Connection $conn | Where-Object { $_.InternalName -eq $SiteColumnInternalName } | Select-Object -First 1

        if ($field) {
            Write-Host -ForegroundColor Green "Found column '$SiteColumnInternalName' in List/Library: $($list.Title)"

            $action = if ($Remediate) { "RemoveFromList" } else { "ReportOnly" }
            $actionResult = "N/A"

            if ($Remediate -and $PSCmdlet.ShouldProcess($TargetSiteUrl, "Remove '$SiteColumnInternalName' from list '$($list.Title)'")) {
                try {
                    # Removes a field from a list or a site [5](https://pnp.github.io/powershell/cmdlets/Remove-PnPField.html)
                    Remove-PnPField -Identity $field -List $list -Connection $conn -Force
                    $actionResult = "Removed"
                    Write-Host -ForegroundColor Yellow "Removed from List/Library: $($list.Title)"
                }
                catch {
                    $actionResult = "Failed: $($_.Exception.Message)"
                    Write-Warning "Failed removing from List '$($list.Title)': $($_.Exception.Message)"
                }
            }

            $results.Add([PSCustomObject]@{
                SiteUrl            = $TargetSiteUrl
                LocationType       = "ListOrLibrary"
                LocationName       = $list.Title
                FieldInternalName  = $SiteColumnInternalName
                Found              = $true
                Action             = $action
                ActionResult       = $actionResult
            }) | Out-Null
        }
    }

    # ---- Remove Site Column (best-effort) ----
    if ($Remediate -and $PSCmdlet.ShouldProcess($TargetSiteUrl, "Remove site column '$SiteColumnInternalName' from site")) {
        try {
            # Removes a field from a list or a site (site columns too) [5](https://pnp.github.io/powershell/cmdlets/Remove-PnPField.html)
            Remove-PnPField -Identity $SiteColumnInternalName -Connection $conn -Force
            $results.Add([PSCustomObject]@{
                SiteUrl            = $TargetSiteUrl
                LocationType       = "SiteColumn"
                LocationName       = "Site Columns"
                FieldInternalName  = $SiteColumnInternalName
                Found              = $true
                Action             = "RemoveSiteColumn"
                ActionResult       = "Removed (best-effort)"
            }) | Out-Null
            Write-Host -ForegroundColor Yellow "Attempted removal of Site Column: $SiteColumnInternalName"
        }
        catch {
            # Still record the outcome, because you want the report even when actions happen
            $results.Add([PSCustomObject]@{
                SiteUrl            = $TargetSiteUrl
                LocationType       = "SiteColumn"
                LocationName       = "Site Columns"
                FieldInternalName  = $SiteColumnInternalName
                Found              = $true
                Action             = "RemoveSiteColumn"
                ActionResult       = "Failed: $($_.Exception.Message)"
            }) | Out-Null
            Write-Warning "Failed removing site column '$SiteColumnInternalName' (may still be in use): $($_.Exception.Message)"
        }
    }

    Disconnect-PnPOnline -Connection $conn -ErrorAction SilentlyContinue | Out-Null
    return $results
}

# ----------------------------
# Main
# ----------------------------
if (-not $ExportPath) {
    $ExportPath = Get-DefaultExportPath
}
Ensure-FolderExists -Path $ExportPath

$final = New-Object System.Collections.Generic.List[object]

if ($TenantWide) {

    if (-not $AdminUrl) {
        throw "AdminUrl is required when using -TenantWide. Example: https://contoso-admin.sharepoint.com"
    }

    Write-Host "Tenant-wide mode enabled. Connecting to admin: $AdminUrl" -ForegroundColor Cyan
    $adminConn = Connect-PnPInteractiveRequired -Url $AdminUrl

    # Get-PnPTenantSite supports -Filter for reducing scope [2](https://deepwiki.com/pnp/powershell/6-microsoft-365-groups-management)
    $sites = if ($SiteFilter) {
        Get-PnPTenantSite -Connection $adminConn -Filter $SiteFilter
    } else {
        Get-PnPTenantSite -Connection $adminConn
    }

    # Hard-coded exclusions (consistent with your other scripts)
    $sites = $sites | Where-Object {
        $_.Template -ne "RedirectSite#0" -and
        $_.Url -notlike "*-my.sharepoint.com/personal*"
    }

    Disconnect-PnPOnline -Connection $adminConn -ErrorAction SilentlyContinue | Out-Null

    foreach ($s in $sites) {
        $siteResults = Process-Site -TargetSiteUrl $s.Url
        foreach ($r in $siteResults) { $final.Add($r) | Out-Null }
    }

} else {

    if (-not $SiteUrl) {
        throw "SiteUrl is required unless -TenantWide is specified."
    }

    $siteResults = Process-Site -TargetSiteUrl $SiteUrl
    foreach ($r in $siteResults) { $final.Add($r) | Out-Null }
}

# Export always (report is always generated)
$final | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
Write-Host "`nReport exported to: $ExportPath" -ForegroundColor Green

if ($PassThru) {
    $final
}