<#
.SYNOPSIS
Remove a specified user as site collection admin from all OneDrive for Business site collections.

.DESCRIPTION
Enumerates all OneDrive sites in the tenant and removes a given account from site collection administrators.
If the account running the script does not have rights on a specific site, the script temporarily adds the runner as SCAdmin, performs the removal, then removes the runner.

.PREREQUISITES
- PnP.PowerShell installed, logged permissions to tenant admin and OneDrive sites.
- Running account must have tenant-level rights to enumerate OneDrive sites.
- ClientId is required for Connect-PnPOnline -Interactive in this environment.

.PARAMETER AdminUrl
Tenant admin URL (eg https://contoso-admin.sharepoint.com).

.PARAMETER ClientId
ClientId for interactive PnP connection.

.PARAMETER UserToRemove
User principal name/login to remove from OneDrive site collection admins.

.PARAMETER PageSize
Page size for Get-PnPTenantSite (default 200).

.PARAMETER MaxRetries
Retry count for transient failures (default 3).

.PARAMETER BaseRetryDelaySeconds
Base retry delay for transient failures (default 2).

.PARAMETER PassThru
Return results as objects to pipeline.

.EXAMPLE
.\OneDrive\Permissions\Onedrive.ps1 -AdminUrl 'https://contoso-admin.sharepoint.com' -ClientId '00000000-0000-0000-0000-000000000000' -UserToRemove 'user@contoso.com' -Confirm

.NOTES
Constructed for repo style and requirements from user request.
Author: Dami Onabanjo
#>

[CmdletBinding(SupportsShouldProcess=$true)]
param(
    [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][string]$AdminUrl,
    [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][string]$ClientId,
    [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][string]$UserToRemove,
    [Parameter(Mandatory=$false)][ValidateRange(10,500)][int]$PageSize = 200,
    [Parameter(Mandatory=$false)][ValidateRange(0,10)][int]$MaxRetries = 3,
    [Parameter(Mandatory=$false)][ValidateRange(1,120)][int]$BaseRetryDelaySeconds = 2,
    [Parameter(Mandatory=$false)][switch]$PassThru
)

function Invoke-WithRetry {
    param(
        [Parameter(Mandatory=$true)][scriptblock]$ScriptBlock,
        [Parameter(Mandatory=$true)][string]$OperationName,
        [int]$maxRetries = $MaxRetries,
        [int]$baseDelaySeconds = $BaseRetryDelaySeconds
    )

    $attempt = 0
    while ($true) {
        try {
            return & $ScriptBlock
        }
        catch {
            $attempt++
            $raw = $_.Exception.Message
            $isThrottle = ($raw -match '429' -or $raw -match 'throttl' -or $raw -match 'Too Many Requests' -or $raw -match '503' -or $raw -match 'temporarily unavailable' -or $raw -match 'timeout')

            if (-not $isThrottle -or $attempt -gt $maxRetries) {
                throw "[$OperationName] failed after $attempt attempts: $raw"
            }

            $delay = [Math]::Min(300, $baseDelaySeconds * [Math]::Pow(2, $attempt - 1))
            Write-Warning "[$OperationName] transient issue (attempt $attempt/$maxRetries): $raw. Retrying in $delay seconds."
            Start-Sleep -Seconds $delay
        }
    }
}

function Remove-UserFromSiteCollectionAdmin {
    param(
        [Parameter(Mandatory=$true)][string]$SiteUrl,
        [Parameter(Mandatory=$true)][object]$Connection,
        [Parameter(Mandatory=$true)][string]$UserToRemove,
        [Parameter(Mandatory=$true)][string]$RunnerLogin
    )

    $result = [PSCustomObject]@{
        SiteUrl = $SiteUrl
        Removed = $false
        Notes = ''
        Status = 'Skipped'
        Timestamp = (Get-Date)
    }

    try {
        $siteAdmins = Get-PnPSiteCollectionAdmin -Connection $Connection -ErrorAction Stop
        $matching = $siteAdmins | Where-Object { $_.LoginName -ieq $UserToRemove -or ($_.UserPrincipalName -and $_.UserPrincipalName -ieq $UserToRemove) }

        if (-not $matching) {
            $result.Notes = "User not a site collection admin."
            return $result
        }

        if ($PSCmdlet.ShouldProcess($SiteUrl, "Remove $UserToRemove from site collection admins")) {
            try {
                Remove-PnPSiteCollectionAdmin -Connection $Connection -Owners $UserToRemove -ErrorAction Stop
                $result.Removed = $true
                $result.Status = 'Removed'
                $result.Notes = 'Removed directly.'
                return $result
            }
            catch {
                $result.Notes = "Direct remove failed: $($_.Exception.Message)"
                Write-Verbose "Direct removal failed on $SiteUrl: $($_.Exception.Message)"
            }

            if ($RunnerLogin -and ($RunnerLogin -ne $UserToRemove)) {
                try {
                    Write-Host "Temporarily adding runner as SCAdmin on $SiteUrl" -ForegroundColor Yellow
                    Add-PnPSiteCollectionAdmin -Connection $Connection -Owners $RunnerLogin -ErrorAction Stop | Out-Null

                    Write-Host "Retry removing $UserToRemove on $SiteUrl" -ForegroundColor Yellow
                    Remove-PnPSiteCollectionAdmin -Connection $Connection -Owners $UserToRemove -ErrorAction Stop
                    $result.Removed = $true
                    $result.Status = 'Removed'
                    $result.Notes = 'Temporarily elevated runner, removed target.'

                    Write-Host "Removing runner from SCAdmin on $SiteUrl" -ForegroundColor Yellow
                    Remove-PnPSiteCollectionAdmin -Connection $Connection -Owners $RunnerLogin -ErrorAction Stop
                    $result.Notes += ' Runner removed after action.'
                }
                catch {
                    $result.Status = 'PartialFailure'
                    $result.Notes = "Failed in elevation/removal workflow: $($_.Exception.Message)"
                }
            }
            else {
                $result.Status = 'Failed'
                $result.Notes += ' Cannot elevate runner because runner and target are the same or runner missing.'
            }
        }

        return $result
    }
    catch {
        $result.Status = 'Failed'
        $result.Notes = "Unexpected error: $($_.Exception.Message)"
        return $result
    }
}

$results = @()
$adminConn = $null

try {
    $adminConn = Invoke-WithRetry -OperationName 'Connect-PnPOnline(Admin)' -ScriptBlock {
        Connect-PnPOnline -Url $AdminUrl -Interactive -ClientId $ClientId -ReturnConnection -ErrorAction Stop
    }

    $runner = Invoke-WithRetry -OperationName 'Get-PnPCurrentUser' -ScriptBlock {
        Get-PnPCurrentUser -Connection $adminConn -ErrorAction Stop
    }
    $runnerLoginName = $runner.LoginName

    $tenantSites = Invoke-WithRetry -OperationName 'Get-PnPTenantSite(OneDrive)' -ScriptBlock {
        Get-PnPTenantSite -Connection $adminConn -IncludeOneDriveSites -Filter "Url -like '-my.sharepoint.com/personal/'" -PageSize $PageSize -ErrorAction Stop
    }

    foreach ($site in $tenantSites) {
        $siteUrl = $site.Url
        Write-Host "Processing site: $siteUrl" -ForegroundColor Cyan

        try {
            $siteConn = Invoke-WithRetry -OperationName "Connect-PnPOnline($siteUrl)" -ScriptBlock {
                Connect-PnPOnline -Url $siteUrl -Interactive -ClientId $ClientId -ReturnConnection -ErrorAction Stop
            }

            $siteResult = Remove-UserFromSiteCollectionAdmin -SiteUrl $siteUrl -Connection $siteConn -UserToRemove $UserToRemove -RunnerLogin $runnerLoginName
            $results += $siteResult

            Disconnect-PnPOnline -Connection $siteConn -ErrorAction SilentlyContinue | Out-Null
        }
        catch {
            $results += [PSCustomObject]@{
                SiteUrl = $siteUrl
                Removed = $false
                Notes = "Could not connect or process site: $($_.Exception.Message)"
                Status = 'Failed'
                Timestamp = (Get-Date)
            }
        }
    }

    if ($PassThru) { $results }

    $summary = $results | Group-Object -Property Status | ForEach-Object { "$_`.Name: $_.Count" }
    Write-Host "Done. Site result summary:`n$($summary -join '\n')" -ForegroundColor Green
}
catch {
    Write-Error "Unhandled error: $($_.Exception.Message)"
    throw
}
finally {
    if ($adminConn) { Disconnect-PnPOnline -Connection $adminConn -ErrorAction SilentlyContinue | Out-Null }
}
