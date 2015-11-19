# Outlook Connector module file
# Current version 0.92, November 2015

# File List
$FileList = @(
    'Connect-Outlook.ps1',
    'Export-OutlookFolder.ps1',
    'Export-OutlookMessage.ps1',
    'Export-OutlookMessageBody.ps1',
    'Get-OutlookMessage.ps1',
    'Get-OutlookFolder.ps1',
    'Get-OutlookInbox.ps1',
    'Start-Outlook.ps1',
    'OutlookConnector.HelperFunctions.ps1')

# Import all files in folder
foreach ($File in $FileList) {
    . (Join-Path -Path $PSScriptRoot -ChildPath $File) # -Verbose:$False
}
