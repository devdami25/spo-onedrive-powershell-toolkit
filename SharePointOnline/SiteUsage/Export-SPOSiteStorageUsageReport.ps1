<#
.SYNOPSIS
Exports SharePoint Online site storage allocation and usage.

.DESCRIPTION
Connects to the SharePoint Online admin center using PnP.PowerShell and exports
storage quota, current usage, and warning level for site collections.

By default, the report excludes OneDrive and Redirect sites. Use
-IncludeOneDrive or -IncludeRedirectSites to include them.

.PARAMETER AdminUrl
SharePoint Online admin center URL. Example: https://contoso-admin.sharepoint.com

.PARAMETER ClientId
Client ID used with Connect-PnPOnline -Interactive.

.PARAMETER ExportPath
Optional CSV output path. If omitted, a timestamped CSV is created in the
current folder. If a directory is provided, a timestamped CSV is created in
that directory.

.PARAMETER IncludeOneDrive
Includes OneDrive for Business sites in the report.

.PARAMETER IncludeRedirectSites
Includes SharePoint redirect sites in the report.

.PARAMETER SiteFilter
Optional filter passed to Get-PnPTenantSite. Example: "Url -like '/sites/HR'"

.PARAMETER PassThru
Returns report objects to the pipeline after exporting them.

.EXAMPLE
.\Export-SPOSiteStorageUsageReport.ps1 `
  -AdminUrl "https://contoso-admin.sharepoint.com" `
  -ClientId "00000000-0000-0000-0000-000000000000"

.EXAMPLE
.\Export-SPOSiteStorageUsageReport.ps1 `
  -AdminUrl "https://contoso-admin.sharepoint.com" `
  -ClientId "00000000-0000-0000-0000-000000000000" `
  -IncludeOneDrive `
  -ExportPath "C:\Temp\SPOStorage.csv"

.NOTES
Requires: PnP.PowerShell (PowerShell 7+ recommended)
Author: Dami Onabanjo
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string] $AdminUrl,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string] $ClientId,

    [Parameter(Mandatory = $false)]
    [string] $ExportPath,

    [Parameter(Mandatory = $false)]
    [switch] $IncludeOneDrive,

    [Parameter(Mandatory = $false)]
    [switch] $IncludeRedirectSites,

    [Parameter(Mandatory = $false)]
    [string] $SiteFilter,

    [Parameter(Mandatory = $false)]
    [switch] $PassThru
)

function Get-DefaultExportPath {
    $timestamp = (Get-Date).ToString("yyyyMMdd-HHmmss")
    Join-Path -Path (Get-Location) -ChildPath ("SPOSiteStorageUsage-{0}.csv" -f $timestamp)
}

function Ensure-FolderExists {
    param([Parameter(Mandatory = $true)][string] $Path)

    $directory = Split-Path -Path $Path -Parent
    if ($directory -and -not (Test-Path -Path $directory)) {
        New-Item -ItemType Directory -Path $directory -Force | Out-Null
    }
}

if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    throw "PnP.PowerShell is required. Install it with: Install-Module PnP.PowerShell -Scope CurrentUser"
}

if ([string]::IsNullOrWhiteSpace($ExportPath)) {
    $ExportPath = Get-DefaultExportPath
}
elseif (Test-Path -Path $ExportPath -PathType Container) {
    $ExportPath = Join-Path -Path $ExportPath -ChildPath (Split-Path -Path (Get-DefaultExportPath) -Leaf)
}

Ensure-FolderExists -Path $ExportPath

Write-Host "Connecting to SharePoint Online admin center: $AdminUrl" -ForegroundColor Cyan
$connection = Connect-PnPOnline -Url $AdminUrl -Interactive -ClientId $ClientId -ReturnConnection

$tenantSiteParameters = @{
    Connection = $connection
    Detailed   = $true
}

if ($SiteFilter) {
    $tenantSiteParameters.Filter = $SiteFilter
}

$siteCollections = Get-PnPTenantSite @tenantSiteParameters | Where-Object {
    ($IncludeOneDrive -or $_.Template -notlike "SPSPERS*") -and
    ($IncludeRedirectSites -or $_.Template -ne "REDIRECTSITE#0")
}

Write-Host "Total number of site collections found: $($siteCollections.Count)" -ForegroundColor Yellow

$report = foreach ($site in $siteCollections) {
    Write-Host "Processing site collection: $($site.Url)" -ForegroundColor Yellow

    [PSCustomObject]@{
        SiteUrl                 = $site.Url
        Title                   = $site.Title
        Template                = $site.Template
        StorageQuotaMB          = $site.StorageQuota
        StorageUsedMB           = $site.StorageUsageCurrent
        StorageWarningLevelMB   = $site.StorageQuotaWarningLevel
        StorageUsagePercent     = if ($site.StorageQuota -gt 0) {
            [math]::Round(($site.StorageUsageCurrent / $site.StorageQuota) * 100, 2)
        }
        else {
            $null
        }
        LastContentModifiedDate = $site.LastContentModifiedDate
    }
}

$report | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding utf8 -ErrorAction Stop
Write-Host "Site storage usage report generated successfully: $ExportPath" -ForegroundColor Green

if ($PassThru) {
    $report
}