function Get-OutlookFolder {

<#
.SYNOPSIS
OutlookConnector function: Returns array of Outlook folders from current Outlook session.

.DESCRIPTION
Function returns array of Outlook first level folders, or all folders from current Outlook session.
Function is returning all properties from a folder for further processing via filtering. Run Get-OutlookFolder | Get-Member for a list of properties.
Output can be piped to Export-OutlookFolder to save all messages to disk in MSG format, which is functionality not provided in Outlook application.
Type Get-Help Get-OutlookFolder -Examples for more usage examples.

.EXAMPLE
Get-OutlookFolder | Select Name
List names of all first level folders in current Outlook session.

.EXAMPLE
(Get-OutlookFolder -Recurse).FullFolderPath | Sort
Lists FullFolderPath for all folders in current Outlook session.

.EXAMPLE
Get-OutlookFolder -Recurse | Select Name,FullFolderPath,@{Name = "Count"; Expression = {$_.Items.Count}} | ? Count -gt 0 | Sort Count -Descending | Out-GridView
Lists all Outlook folders with more than zero messages and displays it in Grid View.

.PARAMETER Outlook
Optional parameter that specifies Outlook session from whch it will obtain needed data.
If omitted, function will connect automatically using Connect-Outlook function.

.PARAMETER Recurse
Optional parameter that specifies to list all folders, and not just first level. In most of the cases, you would need to specify this parameter.

.PARAMETER Progress
If current Outlook session is connected online to remote Exchange server, querying all folders might take a minute. You may display standard progress bar while obtaining that list.

.PARAMETER MainOnly
Specifies not to list any additional folders, except main mailbox. Main mailbox is determined by default Inbox folder. Can be combined with Recurse switch.

.PARAMETER Filter
Specifies string which will be used to filter returned folders. Only folders with names matching filter string will be returned.
It is not needed to use * to specify part of skipped name, i.e. filter string will be searched anywhere within folder names.

.OUTPUTS
Function returns array of Outlook folder objects. Output can be filtered and piped to Export-OutlookFolder.
Outlook can contain also other containers, like Calendar, connected and Archive mailboxes, etc.
Best practice is to run first command without -Recurse to see structure of returned data.

.LINK
about_OutlookConnector

.NOTES
NAME:       Get-OutlookFolder
AUTHOR:     Igor Iric, iricigor@gmail.com
CREATEDATE: September 29, 2015
#>

# ---------------------- [Parameters definitions] ------------------------

[CmdletBinding()]

Param(
    [parameter(Mandatory=$false,ValueFromPipeline=$false)]$Outlook = (Connect-Outlook),
    [switch]$Recurse,
    [switch]$Progress,
    [switch]$MainOnly,
    [parameter(Mandatory=$false,ValueFromPipeline=$false)]$Filter

) #end param

# ------------------------- [Function start] -----------------------------

Write-Verbose -Message 'Get-OutlookFolder starting'

try {
    $Folders = @($Outlook.Folders | ForEach-Object {$_}) # Converting Collection to Array
    Write-Verbose -Message ('    Added '+(@($Folders).Count)+' base folders')
} catch {
    throw 'Folders list not obtained'
}

if ($MainOnly) {
    Write-verbose -Message '    Filtering out folders not in main mailbox.'
    try {
        $olFolders = "Microsoft.Office.Interop.Outlook.OlDefaultFolders" -as [type]
        $InboxDef = $Outlook.GetDefaultFolder($olFolders::olFolderInBox)
        $InboxPath = (($InboxDef.FullFolderPath) -split '\\' | Where-Object {$_})[0]
    } catch {
        throw 'Main mailbox could not be identified.'
    }
    $Folders = @($Folders | Where-Object FullFolderPath -Match $InboxPath)
    Write-Verbose -Message ('    Remaining '+(@($Folders).Count)+' base folders.')
}

if (!$Folders) {throw 'Folders list not obtained'}

# recursivly add subfolders
if ($Recurse) {
    $i = 0
    while ($i -lt (@($Folders).Count)) {
        if ($Progress) {
            $ActivityText = 'Folder: '+($Folders[$i]).FullFolderPath + " ($i/"+ (@($Folders).Count) + ')'
            Write-Progress -Activity $ActivityText -PercentComplete ($i/($Folders.count)*100)
        }
        $Subfolders = $Folders[$i].Folders
        if ($Subfolders -and ((@($Subfolders).Count) -gt 0)) {
            $Folders += ($Subfolders | ForEach-Object {$_})
            Write-Verbose -Message ('    Added '+(@($Subfolders).Count)+' folders from '+($Folders[$i]).Name)
        }
        $i++
    }
} # end of recursive search
if ($Progress) {Write-Progress -Activity ' ' -Completed}

# filtering names
if ($Filter) {
    $Folders = $Folders | Where-Object Name -Match $Filter
}

# return value
$Folders
Write-Verbose -Message 'Get-OutlookFolder finished'

# ------------------------- [End of function] ----------------------------

}
