<#
.SYNOPSIS
Exports a detailed report of SharePoint Online objects with unique permissions (broken inheritance),
including expansion of SharePoint Groups and Entra ID Groups (Security/Microsoft 365 groups).

.DESCRIPTION
Default mode scans ONE site collection (SiteUrl).
Optional tenant-wide mode scans ALL SharePoint site collections (TenantWide), excluding OneDrive and Redirect sites,
and supports scoping via SiteFilter.

The report includes permissions for securable objects with unique role assignments:
- Web (and optionally subsites)
- Lists/Libraries
- Items/Folders/Files (unique permissions only)

For each role assignment:
- Outputs the principal and permission levels
- If principal is a SharePoint Group, expands members via Get-PnPGroupMember. [2](https://github.com/pnp/powershell/blob/dev/documentation/Add-PnPMicrosoft365GroupMember.md)
- If principal is an Entra ID Group (Security/M365/Distribution), expands members via Get-PnPEntraIDGroupMember. [3](https://alyaconsulting.ch/Solutions/PnP.PowerShell/Connect-PnPOnline)
  - Direct members only by default
  - Transitive expansion is optional via -IncludeTransitiveMembers [3](https://alyaconsulting.ch/Solutions/PnP.PowerShell/Connect-PnPOnline)

PERFORMANCE / TIMEOUT PROTECTIONS
- Streams output to CSV (flushes in batches) to avoid large in-memory collections
- Caches expanded group membership to avoid repeated Graph/CSOM calls
- Uses paging for list items
- Supports -SkipItemLevel for faster runs
- Supports -MaxItemsPerList to prevent very large libraries from running forever
- Includes retry/backoff on throttling/timeouts

IMPORTANT / PREREQUISITES
- The account running this script must be Site Collection Admin on:
  - the target site (single-site mode), OR
  - all scanned sites (tenant-wide mode).
- ClientId is required for PnP interactive authentication in this environment.
- Entra group expansion requires Microsoft Graph permissions as per cmdlet documentation
  (e.g., Group.Read.All, Directory.Read.All). [3](https://alyaconsulting.ch/Solutions/PnP.PowerShell/Connect-PnPOnline)
- PnP uses CSOM and often requires explicitly loading properties via Get-PnPProperty. [5](https://deepwiki.com/pnp/powershell/4.4-user-profile-management)[6](https://www.sharepointdiary.com/2018/04/find-all-onedrive-site-collections-in-sharepoint-online-using-powershell.html)

.PARAMETER SiteUrl
Target SharePoint Online site collection URL (default mode).

.PARAMETER ClientId
Required. ClientId for Connect-PnPOnline -Interactive (PnP requires your own Entra app). [4](https://pnp.github.io/powershell/cmdlets/Remove-PnPUser.html)

.PARAMETER ExportPath
Optional. CSV output path. If not specified, a timestamped file is created in the current folder.

.PARAMETER TenantWide
Optional. If set, runs across all SharePoint site collections in the tenant. Uses Get-PnPTenantSite. [1](https://www.sharepointdiary.com/2019/02/get-all-document-libraries-sharepoint-online-pnp-powershell.html)

.PARAMETER AdminUrl
Required when TenantWide is used. Example: https://contoso-admin.sharepoint.com

.PARAMETER SiteFilter
Optional. Reduces tenant-wide scope. Passed to Get-PnPTenantSite -Filter. [1](https://www.sharepointdiary.com/2019/02/get-all-document-libraries-sharepoint-online-pnp-powershell.html)

.PARAMETER IncludeSubsites
Optional. If set, scans subsites as well (where they exist).

.PARAMETER SkipItemLevel
Optional. If set, skips item/folder/file level scanning (faster).

.PARAMETER IncludeTransitiveMembers
Optional. If set, expands Entra group membership transitively (nested groups). [3](https://alyaconsulting.ch/Solutions/PnP.PowerShell/Connect-PnPOnline)
Default is direct members only.

.PARAMETER ListItemPageSize
Page size for list item retrieval. Default 200.

.PARAMETER MaxItemsPerList
Maximum number of items to scan per list/library (safety cap). Default 5000.

.PARAMETER FlushEvery
How many rows to buffer before writing to CSV. Default 500.

.PARAMETER MaxRetries
Max retry attempts for throttling/transient failures. Default 6.

.PARAMETER BaseRetryDelaySeconds
Base delay for exponential backoff. Default 2.

.PARAMETER PassThru
Returns collected output objects to pipeline (note: can be large; CSV streaming is primary).

.EXAMPLE
# Single site collection (default)
.\Export-SPOBrokenInheritanceDetailed.ps1 `
  -SiteUrl "https://contoso.sharepoint.com/sites/HR" `
  -ClientId "00000000-0000-0000-0000-000000000000"

.EXAMPLE
# Tenant-wide with filter (SharePoint sites only)
.\Export-SPOBrokenInheritanceDetailed.ps1 `
  -TenantWide `
  -AdminUrl "https://contoso-admin.sharepoint.com" `
  -SiteFilter "Url -like '/sites/HR'" `
  -ClientId "00000000-0000-0000-0000-000000000000" `
  -ExportPath ".\UniquePerms-HR.csv"

.EXAMPLE
# Include transitive expansion for nested Entra groups (heavier)
.\Export-SPOBrokenInheritanceDetailed.ps1 `
  -SiteUrl "https://contoso.sharepoint.com/sites/Finance" `
  -ClientId "00000000-0000-0000-0000-000000000000" `
  -IncludeTransitiveMembers

.NOTES
References:
- Get-PnPTenantSite [1](https://www.sharepointdiary.com/2019/02/get-all-document-libraries-sharepoint-online-pnp-powershell.html)
- Get-PnPProperty [5](https://deepwiki.com/pnp/powershell/4.4-user-profile-management)[6](https://www.sharepointdiary.com/2018/04/find-all-onedrive-site-collections-in-sharepoint-online-using-powershell.html)
- Get-PnPGroupMember [2](https://github.com/pnp/powershell/blob/dev/documentation/Add-PnPMicrosoft365GroupMember.md)
- Get-PnPEntraIDGroupMember [3](https://alyaconsulting.ch/Solutions/PnP.PowerShell/Connect-PnPOnline)
- PnP auth app requirement context [4](https://pnp.github.io/powershell/cmdlets/Remove-PnPUser.html)

Author: Dami Onabanjo
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string] $SiteUrl,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string] $ClientId,

    [Parameter(Mandatory = $false)]
    [string] $ExportPath,

    [Parameter(Mandatory = $false)]
    [switch] $TenantWide,

    [Parameter(Mandatory = $false)]
    [string] $AdminUrl,

    [Parameter(Mandatory = $false)]
    [string] $SiteFilter,

    [Parameter(Mandatory = $false)]
    [switch] $IncludeSubsites,

    [Parameter(Mandatory = $false)]
    [switch] $SkipItemLevel,

    [Parameter(Mandatory = $false)]
    [switch] $IncludeTransitiveMembers,

    [Parameter(Mandatory = $false)]
    [ValidateRange(50, 2000)]
    [int] $ListItemPageSize = 200,

    [Parameter(Mandatory = $false)]
    [ValidateRange(100, 500000)]
    [int] $MaxItemsPerList = 5000,

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

# ----------------------------
# Utilities
# ----------------------------
function Get-DefaultExportPath {
    $ts = (Get-Date).ToString("yyyyMMdd-HHmmss")
    Join-Path -Path (Get-Location) -ChildPath ("BrokenInheritanceDetailed-{0}.csv" -f $ts)
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
        [Parameter(Mandatory=$true)][scriptblock] $ScriptBlock,
        [Parameter(Mandatory=$true)][string] $OperationName
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
        [Parameter(Mandatory=$true)][string] $Activity,
        [Parameter(Mandatory=$true)][string] $Status,
        [Parameter(Mandatory=$true)][int] $PercentComplete
    )

    $p = [math]::Max(0, [math]::Min(100, $PercentComplete))
    Write-Progress -Activity $Activity -Status $Status -PercentComplete $p
}

# ----------------------------
# Group expansion helpers + caching
# ----------------------------
$script:GroupMemberCache = @{}

function Try-GetGroupIdFromLoginName {
    param([string]$LoginName)

    if ([string]::IsNullOrWhiteSpace($LoginName)) { return $null }

    # Extract GUID at end of claims login name (optionally with _o suffix)
    $m = [regex]::Match($LoginName, '(?i)\|([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})(?:_o)?$')
    if ($m.Success) { return $m.Groups[1].Value }

    return $null
}

function Should-TryEntraExpansion {
    param([string]$LoginName)

    if ([string]::IsNullOrWhiteSpace($LoginName)) { return $false }

    # Common Entra/M365 group claim providers in SPO
    return ($LoginName -like "c:0t.c|tenant|*" -or $LoginName -like "c:0o.c|federateddirectoryclaimprovider|*")
}

function Get-ExpandedMembersForPrincipal {
    param(
        [Parameter(Mandatory=$true)] $Connection,
        [Parameter(Mandatory=$true)] $RoleAssignmentMember
    )

    $principalType  = $RoleAssignmentMember.PrincipalType.ToString()
    $principalTitle = $RoleAssignmentMember.Title
    $principalLogin = $RoleAssignmentMember.LoginName

    # 1) SharePoint Group expansion (in-site)
    if ($principalType -eq "SharePointGroup") {

        $cacheKey = "SPG:$principalTitle"
        if ($script:GroupMemberCache.ContainsKey($cacheKey)) { return $script:GroupMemberCache[$cacheKey] }

        $members = Invoke-WithRetry -OperationName "Get-PnPGroupMember($principalTitle)" -ScriptBlock {
            Get-PnPGroupMember -Connection $Connection -Group $principalTitle  # [2](https://github.com/pnp/powershell/blob/dev/documentation/Add-PnPMicrosoft365GroupMember.md)
        }

        $script:GroupMemberCache[$cacheKey] = $members
        return $members
    }

    # 2) Entra ID group expansion (Security/M365/Distribution)
    # PrincipalType often shows SecurityGroup for Entra groups; we also check claims prefix before attempting.
    $groupId = Try-GetGroupIdFromLoginName -LoginName $principalLogin
    if ($groupId -and (Should-TryEntraExpansion -LoginName $principalLogin) -and ($principalType -in @("SecurityGroup", "DistributionList", "None", "Unknown"))) {

        $cacheKey = if ($IncludeTransitiveMembers) { "EIDT:$groupId" } else { "EID:$groupId" }
        if ($script:GroupMemberCache.ContainsKey($cacheKey)) { return $script:GroupMemberCache[$cacheKey] }

        $members = Invoke-WithRetry -OperationName "Get-PnPEntraIDGroupMember($groupId)" -ScriptBlock {
            if ($IncludeTransitiveMembers) {
                Get-PnPEntraIDGroupMember -Connection $Connection -Identity $groupId -Transitive  # [3](https://alyaconsulting.ch/Solutions/PnP.PowerShell/Connect-PnPOnline)
            } else {
                Get-PnPEntraIDGroupMember -Connection $Connection -Identity $groupId  # [3](https://alyaconsulting.ch/Solutions/PnP.PowerShell/Connect-PnPOnline)
            }
        }

        $script:GroupMemberCache[$cacheKey] = $members
        return $members
    }

    return $null
}

# ----------------------------
# CSV streaming
# ----------------------------
if (-not $ExportPath) { $ExportPath = Get-DefaultExportPath }
Ensure-FolderExists -Path $ExportPath

$buffer = New-Object System.Collections.Generic.List[object]
$headerWritten = $false
$allForPassThru = if ($PassThru) { New-Object System.Collections.Generic.List[object] } else { $null }

function Flush-BufferToCsv {
    if ($buffer.Count -eq 0) { return }

    if (-not $headerWritten) {
        $buffer | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
        $script:headerWritten = $true
    } else {
        $buffer | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8 -Append
    }

    $buffer.Clear()
}

function Add-Row {
    param([Parameter(Mandatory=$true)] $RowObject)

    $buffer.Add($RowObject) | Out-Null
    if ($PassThru) { $allForPassThru.Add($RowObject) | Out-Null }

    if ($buffer.Count -ge $FlushEvery) {
        Flush-BufferToCsv
    }
}

# ----------------------------
# Permission extraction
# ----------------------------
function Add-RoleAssignmentRows {
    param(
        [Parameter(Mandatory = $true)] $Connection,
        [Parameter(Mandatory = $true)] $SecurableObject,
        [Parameter(Mandatory = $true)] [string] $SiteCollectionUrl,
        [Parameter(Mandatory = $true)] [string] $ObjectType,
        [Parameter(Mandatory = $true)] [string] $ObjectTitle,
        [Parameter(Mandatory = $true)] [string] $ObjectUrl,
        [Parameter(Mandatory = $false)] [string] $ListTitle,
        [Parameter(Mandatory = $false)] [string] $ItemId
    )

    # Load HasUniqueRoleAssignments and RoleAssignments for the object
    Invoke-WithRetry -OperationName "Load(RoleAssignments)" -ScriptBlock {
        $null = Get-PnPProperty -Connection $Connection -ClientObject $SecurableObject -Property @("HasUniqueRoleAssignments","RoleAssignments")  # [5](https://deepwiki.com/pnp/powershell/4.4-user-profile-management)[6](https://www.sharepointdiary.com/2018/04/find-all-onedrive-site-collections-in-sharepoint-online-using-powershell.html)
    }

    if (-not $SecurableObject.HasUniqueRoleAssignments) { return }

    foreach ($ra in $SecurableObject.RoleAssignments) {

        Invoke-WithRetry -OperationName "Load(RoleDefinitionBindings/Member)" -ScriptBlock {
            $null = Get-PnPProperty -Connection $Connection -ClientObject $ra -Property @("Member","RoleDefinitionBindings")  # [5](https://deepwiki.com/pnp/powershell/4.4-user-profile-management)[6](https://www.sharepointdiary.com/2018/04/find-all-onedrive-site-collections-in-sharepoint-online-using-powershell.html)
        }

        $principalTitle = $ra.Member.Title
        $principalLogin = $ra.Member.LoginName
        $principalType  = $ra.Member.PrincipalType.ToString()

        $permLevels = ($ra.RoleDefinitionBindings | Select-Object -ExpandProperty Name) |
            Where-Object { $_ -ne "Limited Access" } |
            ForEach-Object { $_.Trim() }

        $permLevelsStr = ($permLevels -join ";")
        if ([string]::IsNullOrWhiteSpace($permLevelsStr)) { $permLevelsStr = "None/Unknown" }

        # Base row
        Add-Row ([PSCustomObject]@{
            SiteCollectionUrl   = $SiteCollectionUrl
            ObjectType          = $ObjectType
            ObjectTitle         = $ObjectTitle
            ObjectUrl           = $ObjectUrl
            ListTitle           = $ListTitle
            ItemId              = $ItemId
            PrincipalType       = $principalType
            PrincipalTitle      = $principalTitle
            PrincipalLoginName  = $principalLogin
            PermissionLevels    = $permLevelsStr
            GrantedThrough      = "Direct"
            ExpandedMemberTitle = $null
            ExpandedMemberEmail = $null
            ExpandedMemberLogin = $null
            ExpandedMemberType  = $null
            ExpansionStatus     = $null
        })

        # Expand group members (SharePoint + Entra)
        try {
            $expandedMembers = Get-ExpandedMembersForPrincipal -Connection $Connection -RoleAssignmentMember $ra.Member

            if ($expandedMembers) {
                foreach ($m in $expandedMembers) {

                    # Normalize likely properties across different member types
                    $mTitle = $m.Title
                    $mEmail = $m.Email
                    $mLogin = $m.LoginName
                    $mType  = $null

                    try { $mType = $m.PrincipalType.ToString() } catch { $mType = $null }

                    if (-not $mTitle -and $m.DisplayName) { $mTitle = $m.DisplayName }
                    if (-not $mEmail -and $m.UserPrincipalName) { $mEmail = $m.UserPrincipalName }
                    if (-not $mLogin -and $m.Id) { $mLogin = $m.Id }
                    if (-not $mType) { $mType = "Unknown" }

                    Add-Row ([PSCustomObject]@{
                        SiteCollectionUrl   = $SiteCollectionUrl
                        ObjectType          = $ObjectType
                        ObjectTitle         = $ObjectTitle
                        ObjectUrl           = $ObjectUrl
                        ListTitle           = $ListTitle
                        ItemId              = $ItemId
                        PrincipalType       = $principalType
                        PrincipalTitle      = $principalTitle
                        PrincipalLoginName  = $principalLogin
                        PermissionLevels    = $permLevelsStr
                        GrantedThrough      = $principalTitle
                        ExpandedMemberTitle = $mTitle
                        ExpandedMemberEmail = $mEmail
                        ExpandedMemberLogin = $mLogin
                        ExpandedMemberType  = $mType
                        ExpansionStatus     = "Expanded"
                    })
                }
            }
        }
        catch {
            Add-Row ([PSCustomObject]@{
                SiteCollectionUrl   = $SiteCollectionUrl
                ObjectType          = $ObjectType
                ObjectTitle         = $ObjectTitle
                ObjectUrl           = $ObjectUrl
                ListTitle           = $ListTitle
                ItemId              = $ItemId
                PrincipalType       = $principalType
                PrincipalTitle      = $principalTitle
                PrincipalLoginName  = $principalLogin
                PermissionLevels    = $permLevelsStr
                GrantedThrough      = $principalTitle
                ExpandedMemberTitle = "[FailedToExpandGroup]"
                ExpandedMemberEmail = $null
                ExpandedMemberLogin = $null
                ExpandedMemberType  = $null
                ExpansionStatus     = $_.Exception.Message
            })
        }
    }
}

function Process-Web {
    param(
        [Parameter(Mandatory = $true)] $Connection,
        [Parameter(Mandatory = $true)] [string] $SiteCollectionUrl,
        [Parameter(Mandatory = $true)] $Web,
        [Parameter(Mandatory = $true)] [string] $WebUrl
    )

    Report-Progress -Activity "Scanning site" -Status "Web: $WebUrl" -PercentComplete 0

    # Web permissions
    Add-RoleAssignmentRows -Connection $Connection -SecurableObject $Web -SiteCollectionUrl $SiteCollectionUrl `
        -ObjectType "Web" -ObjectTitle $Web.Title -ObjectUrl $WebUrl

    # Lists/Libraries
    $skipListTitles = @(
        "Site Assets",
        "Site Pages",
        "Style Library",
        "TaxonomyHiddenList",
        "Form Templates"
    )

    $lists = Get-PnPList -Connection $Connection | Where-Object {
        ($_.Hidden -ne $true) -and
        ($_.IsSystemList -ne $true) -and
        (-not ($skipListTitles -contains $_.Title))
    }  # [7](https://alyaconsulting.ch/Solutions/PnP.PowerShell/Get-PnPGroup)
    $totalLists = $lists.Count
    $listIndex = 0

    foreach ($list in $lists) {
        $listIndex++
        $listPercent = if ($totalLists -gt 0) { [math]::Floor((($listIndex - 1) / $totalLists) * 100) } else { 0 }
        Report-Progress -Activity "Scanning site" -Status "Site: $SiteCollectionUrl; List ($listIndex/$totalLists): $($list.Title)" -PercentComplete $listPercent

        # List URL
        $listUrl = $null
        try {
            $null = Get-PnPProperty -Connection $Connection -ClientObject $list -Property @("RootFolder")  # [5](https://deepwiki.com/pnp/powershell/4.4-user-profile-management)
            $listUrl = $list.RootFolder.ServerRelativeUrl
        } catch { $listUrl = $null }

        # List unique perms
        Add-RoleAssignmentRows -Connection $Connection -SecurableObject $list -SiteCollectionUrl $SiteCollectionUrl `
            -ObjectType "ListOrLibrary" -ObjectTitle $list.Title -ObjectUrl $listUrl -ListTitle $list.Title

        if ($SkipItemLevel) { continue }

        # Safety cap: stop after MaxItemsPerList
        $processed = 0

        # Item-level scanning: try to only fetch items with unique permissions via CAML
        $listItemQuery = "<View><Query><Where><Eq><FieldRef Name='HasUniqueRoleAssignments'/><Value Type='Integer'>1</Value></Eq></Where></Query><RowLimit Paged='TRUE'>$ListItemPageSize</RowLimit></View>"
        $items = $null
        $fetchedOnlyUnique = $false

        try {
            $items = Invoke-WithRetry -OperationName "Get-PnPListItem($($list.Title))" -ScriptBlock {
                Get-PnPListItem -Connection $Connection -List $list -PageSize $ListItemPageSize -Query $listItemQuery -Fields "ID","FileRef","FileLeafRef","Title","HasUniqueRoleAssignments"
            }
            $fetchedOnlyUnique = $true
        }
        catch {
            Write-Verbose "CAML-based unique-perms filter failed for '$($list.Title)': $_. Falling back to full item scan."
            $items = Invoke-WithRetry -OperationName "Get-PnPListItem($($list.Title))" -ScriptBlock {
                Get-PnPListItem -Connection $Connection -List $list -PageSize $ListItemPageSize
            }
        }

        $totalItems = if ($items -ne $null) { $items.Count } else { 0 }

        foreach ($item in $items) {
            $processed++
            $itemPercent = if ($totalItems -gt 0) { [math]::Floor(($processed / $totalItems) * 100) } else { 0 }
            Report-Progress -Activity "Scanning site" -Status "Site: $SiteCollectionUrl; List: $($list.Title); Item $processed/$totalItems" -PercentComplete $itemPercent

            if ($processed -gt $MaxItemsPerList) {
                Write-Warning "List '$($list.Title)' exceeded MaxItemsPerList ($MaxItemsPerList). Stopping item-level scan for this list."
                break
            }

            if (-not $fetchedOnlyUnique) {
                # Load uniqueness + role assignments in fallback mode
                Invoke-WithRetry -OperationName "Load(ItemRoleAssignments)" -ScriptBlock {
                    $null = Get-PnPProperty -Connection $Connection -ClientObject $item -Property @("HasUniqueRoleAssignments","RoleAssignments")
                }

                if (-not $item.HasUniqueRoleAssignments) { continue }
            }

            # Best-effort item URL + title
            $itemUrl = $null
            $itemTitle = $null
            try { $itemUrl = $item.FieldValues["FileRef"] } catch {}
            try {
                $itemTitle = $item.FieldValues["FileLeafRef"]
                if (-not $itemTitle) { $itemTitle = $item.FieldValues["Title"] }
            } catch {}

            Add-RoleAssignmentRows -Connection $Connection -SecurableObject $item -SiteCollectionUrl $SiteCollectionUrl `
                -ObjectType "Item" -ObjectTitle $itemTitle -ObjectUrl $itemUrl -ListTitle $list.Title -ItemId $item.Id
        }
    }
}

function Process-SiteCollection {
    param([Parameter(Mandatory = $true)][string] $TargetSiteUrl)

    Write-Host "`nProcessing site collection: $TargetSiteUrl" -ForegroundColor Cyan
    Report-Progress -Activity "Processing site collection" -Status "Connecting to $TargetSiteUrl" -PercentComplete 0

    $conn = Connect-PnPWithClientId -Url $TargetSiteUrl

    try {
        $web = Get-PnPWeb -Connection $conn
        Process-Web -Connection $conn -SiteCollectionUrl $TargetSiteUrl -Web $web -WebUrl $TargetSiteUrl

        if ($IncludeSubsites) {
            $subwebs = Get-PnPSubWeb -Connection $conn -Recurse -ErrorAction SilentlyContinue
            foreach ($sw in $subwebs) {
                Process-Web -Connection $conn -SiteCollectionUrl $TargetSiteUrl -Web $sw -WebUrl $sw.Url
            }
        }

        Report-Progress -Activity "Processing site collection" -Status "Completed $TargetSiteUrl" -PercentComplete 100
        Write-Progress -Activity "Processing site collection" -Completed
    }
    finally {
        Disconnect-PnPOnline -Connection $conn -ErrorAction SilentlyContinue | Out-Null
    }
}

# ----------------------------
# MAIN
# ----------------------------
try {
    if ($TenantWide) {
        if (-not $AdminUrl) { throw "AdminUrl is required when using -TenantWide." }

        Write-Host "Tenant-wide mode enabled. Connecting to: $AdminUrl" -ForegroundColor Cyan
        $adminConn = Connect-PnPWithClientId -Url $AdminUrl

        try {
            $sites = if ($SiteFilter) {
                Get-PnPTenantSite -Connection $adminConn -Filter $SiteFilter  # [1](https://www.sharepointdiary.com/2019/02/get-all-document-libraries-sharepoint-online-pnp-powershell.html)
            } else {
                Get-PnPTenantSite -Connection $adminConn  # [1](https://www.sharepointdiary.com/2019/02/get-all-document-libraries-sharepoint-online-pnp-powershell.html)
            }

            # Hard-coded exclusions: OneDrive + Redirect sites
            $sites = $sites | Where-Object {
                $_.Template -ne "RedirectSite#0" -and $_.Url -notlike "*-my.sharepoint.com/personal*"
            }
        }
        finally {
            Disconnect-PnPOnline -ErrorAction SilentlyContinue | Out-Null
        }

        $i = 0
        foreach ($s in $sites) {
            $i++
            Write-Progress -Activity "Scanning sites" -Status "$i / $($sites.Count) : $($s.Url)" -PercentComplete (($i / [Math]::Max(1,$sites.Count)) * 100)

            Invoke-WithRetry -OperationName "Process-SiteCollection" -ScriptBlock {
                Process-SiteCollection -TargetSiteUrl $s.Url
            }
            Flush-BufferToCsv
        }

        Write-Progress -Activity "Scanning sites" -Completed
    }
    else {
        if (-not $SiteUrl) { throw "SiteUrl is required unless -TenantWide is specified." }

        Invoke-WithRetry -OperationName "Process-SiteCollection" -ScriptBlock {
            Process-SiteCollection -TargetSiteUrl $SiteUrl
        }
        Flush-BufferToCsv
    }
}
finally {
    Flush-BufferToCsv
}

Write-Host "`nReport exported to: $ExportPath" -ForegroundColor Green

if ($PassThru) {
    $allForPassThru
}