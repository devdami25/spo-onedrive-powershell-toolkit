<#
.SYNOPSIS
Reports and optionally removes SharePoint Online file/folder sharing links to reduce unique
role assignments (security scopes) on a site collection, helping avoid the 50,000 unique
permissions-per-list/library limit.

.DESCRIPTION
Scans document libraries in a site collection (or tenant-wide) for files/folders with unique
permissions (HasUniqueRoleAssignments), then enumerates their sharing links via
Get-PnPFileSharingLink / Get-PnPFolderSharingLink.

Sharing links can be filtered by:
- Link type: Anyone (Anonymous), Organization (Company), Specific People
- Link status: Active, Expired, Never Expires, Has Expiration, Soon-to-expire (within N days)

When -RemoveSharingLinks is specified (and confirmed via ShouldProcess/-Confirm), matching links
are removed via Remove-PnPFileSharingLink / Remove-PnPFolderSharingLink. Without this switch, the
script only reports on links found (dry run / preview mode).

PERFORMANCE / TIMEOUT PROTECTIONS
- Streams output to CSV (flushes in batches) to avoid large in-memory collections
- Uses paging for list item retrieval
- Includes retry/backoff on throttling/timeouts

IMPORTANT / PREREQUISITES
- The account running this script must be Site Collection Admin on the target site(s).
- ClientId is required for PnP interactive authentication in this environment.
- Removal is destructive; run without -RemoveSharingLinks first to preview results.

.PARAMETER SiteUrl
Target SharePoint Online site collection URL (single-site mode).

.PARAMETER ClientId
Required. ClientId for Connect-PnPOnline -Interactive.

.PARAMETER TenantWide
Optional. If set, runs across all SharePoint site collections in the tenant (excludes OneDrive
and system template sites). Requires -AdminUrl.

.PARAMETER AdminUrl
Required when -TenantWide is used. Example: https://contoso-admin.sharepoint.com

.PARAMETER SiteFilter
Optional. Reduces tenant-wide scope. Passed to Get-PnPTenantSite -Filter.

.PARAMETER LinkType
Optional. Filter sharing links by type: Anyone, Organization, SpecificPeople, or All. Default All.

.PARAMETER ActiveLinks
Optional. If set, only includes active (non-expired) links.

.PARAMETER ExpiredLinks
Optional. If set, only includes expired links.

.PARAMETER NeverExpiresLinks
Optional. If set, only includes links with no expiration set.

.PARAMETER LinksWithExpiration
Optional. If set, only includes links that have an expiration date set.

.PARAMETER SoonToExpireInDays
Optional. If set, only includes active links expiring within this many days.

.PARAMETER RemoveSharingLinks
Optional. If set, removes matching sharing links (subject to -Confirm/-WhatIf via ShouldProcess).
Without this switch, the script only reports.

.PARAMETER ExportPath
Optional. CSV output path. If not specified, a timestamped file is created in the current folder.

.PARAMETER ListItemPageSize
Page size for list item retrieval. Default 200.

.PARAMETER FlushEvery
How many rows to buffer before writing to CSV. Default 500.

.PARAMETER MaxRetries
Max retry attempts for throttling/transient failures. Default 6.

.PARAMETER BaseRetryDelaySeconds
Base delay for exponential backoff. Default 2.

.PARAMETER PassThru
Returns collected output objects to pipeline (note: can be large; CSV streaming is primary).

.EXAMPLE
# Preview all sharing links on a single site (no removal)
.\Remove-SPOSharingLinks.ps1 `
  -SiteUrl "https://contoso.sharepoint.com/sites/HR" `
  -ClientId "00000000-0000-0000-0000-000000000000"

.EXAMPLE
# Remove expired Anyone links tenant-wide
.\Remove-SPOSharingLinks.ps1 `
  -TenantWide `
  -AdminUrl "https://contoso-admin.sharepoint.com" `
  -ClientId "00000000-0000-0000-0000-000000000000" `
  -LinkType Anyone `
  -ExpiredLinks `
  -RemoveSharingLinks `
  -Confirm

.EXAMPLE
# Remove links expiring within 7 days on a single site
.\Remove-SPOSharingLinks.ps1 `
  -SiteUrl "https://contoso.sharepoint.com/sites/Finance" `
  -ClientId "00000000-0000-0000-0000-000000000000" `
  -SoonToExpireInDays 7 `
  -RemoveSharingLinks `
  -Confirm

.NOTES
References:
- Get-PnPFileSharingLink / Get-PnPFolderSharingLink
- Remove-PnPFileSharingLink / Remove-PnPFolderSharingLink
- Get-PnPTenantSite

Author: Dami Onabanjo
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $false)]
    [string] $SiteUrl,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string] $ClientId,

    [Parameter(Mandatory = $false)]
    [switch] $TenantWide,

    [Parameter(Mandatory = $false)]
    [string] $AdminUrl,

    [Parameter(Mandatory = $false)]
    [string] $SiteFilter,

    [Parameter(Mandatory = $false)]
    [ValidateSet("All", "Anyone", "Organization", "SpecificPeople")]
    [string] $LinkType = "All",

    [Parameter(Mandatory = $false)]
    [switch] $ActiveLinks,

    [Parameter(Mandatory = $false)]
    [switch] $ExpiredLinks,

    [Parameter(Mandatory = $false)]
    [switch] $NeverExpiresLinks,

    [Parameter(Mandatory = $false)]
    [switch] $LinksWithExpiration,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 3650)]
    [int] $SoonToExpireInDays,

    [Parameter(Mandatory = $false)]
    [switch] $RemoveSharingLinks,

    [Parameter(Mandatory = $false)]
    [string] $ExportPath,

    [Parameter(Mandatory = $false)]
    [ValidateRange(50, 2000)]
    [int] $ListItemPageSize = 200,

    [Parameter(Mandatory = $false)]
    [ValidateRange(50, 5000)]
    [int] $FlushEvery = 500,

    [Parameter(Mandatory = $false)]
    [ValidateRange(0, 20)]
    [int] $MaxRetries = 6,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 60)]
    [int] $BaseRetryDelaySeconds = 2,

    [Parameter(Mandatory = $false)]
    [switch] $PassThru
)

# Excluded system libraries that never carry meaningful sharing links
$script:ExcludedLists = @(
    "Form Templates", "Style Library", "Site Assets", "Site Pages",
    "Preservation Hold Library", "Pages", "Images",
    "Site Collection Documents", "Site Collection Images"
)

# ----------------------------
# Utilities
# ----------------------------
function Get-DefaultExportPath {
    $ts = (Get-Date).ToString("yyyyMMdd-HHmmss")
    Join-Path -Path (Get-Location) -ChildPath ("SPO-SharingLinks-{0}.csv" -f $ts)
}

function Ensure-FolderExists {
    param([Parameter(Mandatory = $true)][string] $Path)
    $dir = Split-Path -Path $Path -Parent
    if ($dir -and -not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
}

function Invoke-WithRetry {
    param(
        [Parameter(Mandatory = $true)][scriptblock] $ScriptBlock,
        [Parameter(Mandatory = $true)][string] $OperationName
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

            if (-not $isThrottle -or $attempt -gt $MaxRetries) {
                throw
            }

            $delay = [Math]::Min(300, ($BaseRetryDelaySeconds * [Math]::Pow(2, ($attempt - 1))))
            Write-Warning "[$OperationName] transient failure (attempt $attempt/$MaxRetries). Waiting $delay seconds then retrying... $msg"
            Start-Sleep -Seconds $delay
        }
    }
}

function Connect-PnPWithClientId {
    param([Parameter(Mandatory = $true)][string] $Url)

    Invoke-WithRetry -OperationName "Connect-PnPOnline" -ScriptBlock {
        Connect-PnPOnline -Url $Url -Interactive -ClientId $ClientId -ReturnConnection
    }
}

function Report-Progress {
    param(
        [Parameter(Mandatory = $true)][string] $Activity,
        [Parameter(Mandatory = $true)][string] $Status,
        [Parameter(Mandatory = $true)][int] $PercentComplete
    )

    $p = [math]::Max(0, [math]::Min(100, $PercentComplete))
    Write-Progress -Activity $Activity -Status $Status -PercentComplete $p
}

# ----------------------------
# CSV streaming
# ----------------------------
if (-not $ExportPath) { $ExportPath = Get-DefaultExportPath }
Ensure-FolderExists -Path $ExportPath

$buffer = New-Object System.Collections.Generic.List[object]
$headerWritten = $false
$allForPassThru = if ($PassThru) { New-Object System.Collections.Generic.List[object] } else { $null }
$script:ItemCount = 0

function Flush-BufferToCsv {
    if ($buffer.Count -eq 0) { return }

    if (-not $headerWritten) {
        $buffer | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
        $script:headerWritten = $true
    }
    else {
        $buffer | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8 -Append
    }

    $buffer.Clear()
}

function Add-Row {
    param([Parameter(Mandatory = $true)] $RowObject)

    $buffer.Add($RowObject) | Out-Null
    if ($PassThru) { $allForPassThru.Add($RowObject) | Out-Null }
    $script:ItemCount++

    if ($buffer.Count -ge $FlushEvery) {
        Flush-BufferToCsv
    }
}

# ----------------------------
# Link filtering
# ----------------------------
function Test-LinkTypeMatches {
    param([Parameter(Mandatory = $true)][string] $Scope)

    switch ($LinkType) {
        "Anyone" { return $Scope -eq "Anonymous" }
        "Organization" { return $Scope -eq "Organization" }
        "SpecificPeople" { return $Scope -eq "Users" }
        default { return $true }
    }
}

function Get-LinkStatusInfo {
    param([Parameter(Mandatory = $false)] $ExpirationDate)

    $currentDateTime = (Get-Date).Date

    if ($null -eq $ExpirationDate) {
        return [PSCustomObject]@{
            Status            = "Active"
            ExpiryDays        = $null
            FriendlyExpiry    = "Never Expires"
        }
    }

    $expiryDate = ([DateTime]$ExpirationDate).ToLocalTime()
    $expiryDays = (New-TimeSpan -Start $currentDateTime -End $expiryDate).Days

    if ($expiryDate -lt $currentDateTime) {
        return [PSCustomObject]@{
            Status         = "Expired"
            ExpiryDays     = $expiryDays
            FriendlyExpiry = "Expired $([Math]::Abs($expiryDays)) days ago"
        }
    }

    return [PSCustomObject]@{
        Status         = "Active"
        ExpiryDays     = $expiryDays
        FriendlyExpiry = "Expires in $expiryDays days"
    }
}

function Test-LinkStatusMatches {
    param(
        [Parameter(Mandatory = $true)] $StatusInfo,
        [Parameter(Mandatory = $false)] $ExpirationDate
    )

    if ($ActiveLinks -and $StatusInfo.Status -ne "Active") { return $false }
    if ($ExpiredLinks -and $StatusInfo.Status -ne "Expired") { return $false }
    if ($LinksWithExpiration -and $null -eq $ExpirationDate) { return $false }
    if ($NeverExpiresLinks -and $StatusInfo.FriendlyExpiry -ne "Never Expires") { return $false }

    if ($SoonToExpireInDays -gt 0) {
        if ($null -eq $ExpirationDate -or $StatusInfo.ExpiryDays -lt 0 -or $StatusInfo.ExpiryDays -gt $SoonToExpireInDays) {
            return $false
        }
    }

    return $true
}

# ----------------------------
# Sharing link scan/removal
# ----------------------------
function Remove-SharingLinkForObject {
    param(
        [Parameter(Mandatory = $true)][string] $ObjectType,
        [Parameter(Mandatory = $true)][string] $FileUrl,
        [Parameter(Mandatory = $true)][string] $LinkId
    )

    try {
        if ($ObjectType -eq "File") {
            Invoke-WithRetry -OperationName "Remove-PnPFileSharingLink" -ScriptBlock {
                Remove-PnPFileSharingLink -FileUrl $FileUrl -Identity $LinkId -Force
            }
        }
        else {
            Invoke-WithRetry -OperationName "Remove-PnPFolderSharingLink" -ScriptBlock {
                Remove-PnPFolderSharingLink -Folder $FileUrl -Identity $LinkId -Force
            }
        }
        return "Success"
    }
    catch {
        Write-Warning "Failed to remove sharing link on '$FileUrl': $($_.Exception.Message)"
        return "Error occurred"
    }
}

function Get-SiteSharingLinks {
    param(
        [Parameter(Mandatory = $true)][string] $SiteName,
        [Parameter(Mandatory = $true)][string] $SiteCollectionUrl
    )

    $documentLibraries = Invoke-WithRetry -OperationName "Get-PnPList" -ScriptBlock {
        Get-PnPList | Where-Object { $_.Hidden -eq $false -and $_.Title -notin $script:ExcludedLists -and $_.BaseType -eq "DocumentLibrary" }
    }

    foreach ($list in $documentLibraries) {

        $listItems = Invoke-WithRetry -OperationName "Get-PnPListItem($($list.Title))" -ScriptBlock {
            Get-PnPListItem -List $list -PageSize $ListItemPageSize
        }

        foreach ($item in $listItems) {
            $fileName = $item.FieldValues.FileLeafRef
            $objectType = $item.FileSystemObjectType.ToString()

            Report-Progress -Activity "Site: $SiteName" -Status "Library: $($list.Title) | Item: $fileName" -PercentComplete 0

            if ($objectType -notin @("File", "Folder")) { continue }

            $hasUniquePermissions = Invoke-WithRetry -OperationName "Get-PnPProperty(HasUniqueRoleAssignments)" -ScriptBlock {
                Get-PnPProperty -ClientObject $item -Property HasUniqueRoleAssignments
            }

            if (-not $hasUniquePermissions) { continue }

            $fileUrl = $item.FieldValues.FileRef

            $sharingLinks = Invoke-WithRetry -OperationName "Get-PnP$($objectType)SharingLink" -ScriptBlock {
                if ($objectType -eq "File") {
                    Get-PnPFileSharingLink -Identity $fileUrl
                }
                else {
                    Get-PnPFolderSharingLink -Folder $fileUrl
                }
            }

            foreach ($sharingLink in $sharingLinks) {
                $link = $sharingLink.Link
                $scope = $link.Scope

                if (-not (Test-LinkTypeMatches -Scope $scope)) { continue }

                $expirationDate = $sharingLink.ExpirationDateTime
                $statusInfo = Get-LinkStatusInfo -ExpirationDate $expirationDate

                if (-not (Test-LinkStatusMatches -StatusInfo $statusInfo -ExpirationDate $expirationDate)) { continue }

                $linkRemovalStatus = "No action performed"
                if ($RemoveSharingLinks) {
                    if ($PSCmdlet.ShouldProcess($fileUrl, "Remove $scope sharing link")) {
                        $linkRemovalStatus = Remove-SharingLinkForObject -ObjectType $objectType -FileUrl $fileUrl -LinkId $sharingLink.Id
                    }
                    else {
                        $linkRemovalStatus = "Skipped (WhatIf)"
                    }
                }

                $directUsers = ($sharingLink.GrantedToIdentitiesV2.User.Email) -join ","

                Add-Row ([PSCustomObject]@{
                    SiteName            = $SiteName
                    SiteCollectionUrl   = $SiteCollectionUrl
                    Library             = $list.Title
                    ObjectType          = $objectType
                    FileName            = $fileName
                    FileUrl             = $fileUrl
                    LinkType            = $scope
                    AccessType          = $link.Type
                    Roles               = ($sharingLink.Roles -join ",")
                    Users               = $directUsers
                    FileType            = $item.FieldValues.File_x0020_Type
                    LinkStatus          = $statusInfo.Status
                    LinkExpiryDate      = $expirationDate
                    FriendlyExpiryTime  = $statusInfo.FriendlyExpiry
                    PasswordProtected   = $sharingLink.HasPassword
                    BlockDownload       = $link.PreventsDownload
                    SharedLink          = $link.WebUrl
                    LinkRemovalStatus   = $linkRemovalStatus
                })
            }
        }
    }
}

# ----------------------------
# Main
# ----------------------------
if ($TenantWide -and -not $AdminUrl) {
    throw "-AdminUrl is required when -TenantWide is specified."
}
if (-not $TenantWide -and -not $SiteUrl) {
    throw "-SiteUrl is required unless -TenantWide is specified."
}

$targetSites = @()

if ($TenantWide) {
    Connect-PnPWithClientId -Url $AdminUrl | Out-Null

    $tenantSites = Invoke-WithRetry -OperationName "Get-PnPTenantSite" -ScriptBlock {
        if ($SiteFilter) {
            Get-PnPTenantSite -Filter $SiteFilter
        }
        else {
            Get-PnPTenantSite
        }
    }

    $targetSites = $tenantSites | Where-Object {
        $_.Template -notin @("SRCHCEN#0", "REDIRECTSITE#0", "SPSMSITEHOST#0", "APPCATALOG#0", "POINTPUBLISHINGHUB#0", "EDISC#0", "STS#-1") -and
        $_.Url -notlike "*-my.sharepoint.com/personal/*"
    } | ForEach-Object { $_.Url }

    Disconnect-PnPOnline -ErrorAction SilentlyContinue | Out-Null
}
else {
    $targetSites = @($SiteUrl)
}

$total = $targetSites.Count
$counter = 0

foreach ($url in $targetSites) {
    $counter++
    Report-Progress -Activity "Scanning site collections" -Status "$counter / $total : $url" -PercentComplete (($counter / [Math]::Max($total, 1)) * 100)

    try {
        Connect-PnPWithClientId -Url $url | Out-Null
        $siteName = (Invoke-WithRetry -OperationName "Get-PnPWeb" -ScriptBlock { Get-PnPWeb | Select-Object -ExpandProperty Title })

        Get-SiteSharingLinks -SiteName $siteName -SiteCollectionUrl $url
    }
    catch {
        Write-Warning "Failed to process site '$url': $($_.Exception.Message)"
    }
    finally {
        Disconnect-PnPOnline -ErrorAction SilentlyContinue | Out-Null
    }
}

Report-Progress -Activity "Scanning site collections" -Status "Completed" -PercentComplete 100
Flush-BufferToCsv

Write-Host "`nDone. Found $script:ItemCount sharing link(s)." -ForegroundColor Green
if (Test-Path -Path $ExportPath) {
    Write-Host "Report available at: $ExportPath`n" -ForegroundColor Yellow
}

if ($PassThru) { $allForPassThru }
