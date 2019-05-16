param($Credential=(Get-Credential))
Connect-AzureAD -Credential $Credential

# Connect Exchange Online 
$EXOSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Credential -Authentication Basic -AllowRedirection
$EXOComments = Import-PSSession $EXOSession -DisableNameChecking
 
# Get SharePoint AdminCenter URL
$SharePointURL = (Get-OrganizationConfig).SharePointURL
$SharePointAdminURL = $SharePointURL -Replace(".sharepoint.com","-admin.sharepoint.com")

Write-Host $SharePointURL
Write-Host $SharePointAdminURL

# Connect SharePoint Online
$SPOSession = Connect-SPOService -Credential $Credential -Url $SharePointAdminURL

# Connect Teams
$TeamsSession = Connect-MicrosoftTeams -Credential $Credential 

# Set Csv export path and filename
$CSVFile="C:\TeamsSize.CSV"

# Get All Teams
$Teams = Get-Team
$TeamsSizeReport = @()

foreach ($Team in $Teams)
{   
    #Get Teamのoffice365 group Information
    $UnifiedGroup = Get-UnifiedGroup -Identity $Team.GroupId
    #Get Team site url 
    $SPOSite = Get-SPOSite -Identity $UnifiedGroup.SharePointSiteUrl
    #Get Teamchat folder form Office 365 mailbox
    $FolderStatistics = Get-MailboxFolderStatistics -Identity $UnifiedGroup.Identity | Where {$_.FolderPath -eq "/Conversation History/Team Chat"}
 
    $OutputItem = New-Object Object
    $OutputItem | Add-Member TeamDisplayName $Team.DisplayName
    $OutputItem | Add-Member TeamAddress $UnifiedGroup.PrimarySmtpAddress
    $OutputItem | Add-Member StorageUsedMB $SPOSite.StorageUsageCurrent
    $OutputItem | Add-Member TeamChatsinMBX $FolderStatistics.ItemsInFolder
 
    $TeamsSizeReport+=$OutputItem
}
 
Remove-PSSession -Session $EXOSession
Disconnect-SPOService
Disconnect-MicrosoftTeams

$TeamsSizeReport
# Export information to csv 
$TeamsSizeReport | Select * | Export-Csv -Path $CSVFile -NoTypeInformation -Encoding UTF8
