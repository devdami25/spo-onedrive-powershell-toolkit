#This PowerShell retrieves SharePoint Online site collection quota in MBs.

#Config Parameters
$AdminSiteURL="https://domain-admin.sharepoint.com"
$ReportOutput="C:\Temp\SPOStorage.csv"
 

#Connect to SharePoint Online Admin Center
Connect-SPOService -Url $AdminSiteURL
 
#Get all Site collections
$SiteCollections = Get-SPOSite -Limit All
Write-Host "Total Number of Site collections Found:"$SiteCollections.count -f Yellow
 
#Array to store Result
$ResultSet = @()
 
Foreach($Site in $SiteCollections)
{
    Write-Host "Processing Site Collection :"$Site.URL -f Yellow
    #Send the Result to CSV
    $Result = new-object PSObject
    $Result| add-member -membertype NoteProperty -name "SiteURL" -Value $Site.URL
    $Result | add-member -membertype NoteProperty -name "Allocated" -Value $Site.StorageQuota
    $Result | add-member -membertype NoteProperty -name "Used" -Value $Site.StorageUsageCurrent
    $Result | add-member -membertype NoteProperty -name "Warning Level" -Value  $site.StorageQuotaWarningLevel
    $ResultSet += $Result
}
 
#Export Result to csv file
$ResultSet |  Export-Csv $ReportOutput -notypeinformation
 
Write-Host "Site Quota Report Generated Successfully!" -f Green