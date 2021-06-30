﻿function Export-OutlookFolder {

<#
.SYNOPSIS
OutlookConnector function: Saves all messages from passed Outlook folder to a disk.

.DESCRIPTION
Saves all messages from passed Outlook folder(s) to a disk. It will also process all contained subfolders.
Function is internally calling Export-OutlookMessage function.
Folder(s) can be obtained via Get-OutlookFolder function and piped or used as an argument.

.EXAMPLE
Get-OutlookFolder -Recurse -MainOnly | ? Name -eq 'Done' | Export-OutlookFolder -OutputFolder 'C:\tmp\Done'
Saves all messages from folder named 'Done' to a disk using default file naming.

.EXAMPLE
Get-OutlookFolder -recurse  | WHERE Name -in 'Home','Finance'   | Export-OutlookFolder -OutputFolder 'C:\temp\Export-Mail-Temp' -ExportFormat HTML
Saves all messages from folders and sub-folders of 'Home' and 'Finance' (if not empty) to folder 'C:\temp\Export-Mail-Temp' in the HTML format.
Using this feature, one can avoid passing objects from Get-OutlookFolder via Export-OutlookMessage over to Export-OutlookMessageBody, 
in order to export folder messages to HTML format.

.PARAMETER InputFolder
Mandatory parameter that specifies which Outlook folder needs to be exported. Easies is to obtain it via 

.PARAMETER OutputFolder
Mandatory parameter which specifies to which folder messages will be saved. It can be both local disk, as well as network location.
If folder is not existing, it will be created.
Entire Outlok folder structure will be generated within specified folder.

.PARAMETER FileNameFormat
Optional parameter that specifies how individual files will be named based. If omitted, files will be saved in format 'FROM= %SenderName% SUBJECT= %Subject%'.
File name can contain any of message parameters surrounded with %. For list of parameters, type Get-OutlookInbox | Get-Member.
Custom format can be specified after a | character within the %, e.g. %ReceivedTime|yyyyMMddhhmmss%.
Parameter is passed to Export-OutlookMessage function.

.PARAMETER ExportFormat
Optional parameter which specifies to which format message body will be exported to. Allowed values are MSG (default), HTML, TXT (text) and RTF (rich-text).
In case of MSG, the Export-OutlookMessage will be called to do the job. For all others, the Export-OutlookMessageBody will be called.
This makes it possible to for example export folders from Outlook to HTML-format

.PARAMETER Filter
Optional parameter that can contain a filter string expression to be applied to restrict items to be exported.
For syntax see https://msdn.microsoft.com/en-us/vba/outlook-vba/articles/items-restrict-method-outlook.

.PARAMETER IncludeTypes
Optional parameter to specify specific types of items to be exported, such as olMail for e-mail items.
To list all possible values: [enum]::GetNames([Microsoft.Office.Interop.Outlook.OlObjectClass])

.PARAMETER ExcludeTypes
Optional parameter to specify specific types of items to not export, such as olContact for contact items.
To list all possible values: [enum]::GetNames([Microsoft.Office.Interop.Outlook.OlObjectClass])

.PARAMETER Progress
If current Outlook session is connected online to remote Exchange server, saving all folders might take a minute. You may display standard progress bar while obtaining that list.

.OUTPUTS
Function returns array of Outlook folder objects. Output can be filtered and piped to Export-OutlookFolder.
Outlook can contain also other containers, like Calendar, connected and Archive mailboxes, etc.
Best practice is to run first command without -Recurse to see structure of returned data.

.LINK
about_OutlookConnector

.NOTES
NAME:       Export-OutlookFolder
AUTHOR:     Igor Iric, iricigor@gmail.com
CREATEDATE: September 29, 2015
UPDATE:     May 18, 2020 by Peter Rosenberg aka PetRose
#>

# ---------------------- [Parameters definitions] ------------------------

[CmdletBinding()]

Param(
    [parameter(Mandatory=$true,ValueFromPipeline=$true,Position=0)][psobject[]]$InputFolder,
    [parameter(Mandatory=$true,ValueFromPipeline=$false)][string]$OutputFolder,
    [parameter(Mandatory=$false,ValueFromPipeline=$false)][string]$FileNameFormat='FROM= %SenderName% SUBJECT= %Subject%',
    [parameter(Mandatory=$false,ValueFromPipeline=$false)][string][ValidateSet('MSG','HTML','TXT','RTF')] [string]$ExportFormat='MSG',
    [parameter(Mandatory=$false,ValueFromPipeline=$false)][string]$Filter,
    [parameter(Mandatory=$false,ValueFromPipeline=$false)][Microsoft.Office.Interop.Outlook.OlObjectClass[]]$IncludeTypes,
    [parameter(Mandatory=$false,ValueFromPipeline=$false)][Microsoft.Office.Interop.Outlook.OlObjectClass[]]$ExcludeTypes,
    [switch]$Progress

) #end param

# ------------------------- [Function start] -----------------------------

BEGIN {

    Write-Verbose -Message 'Export-OutlookFolder starting...'
    $ReqProps = @('Items','FullFolderPath','Folders')
    $OutputFolderPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputFolder)

} # End of BEGIN block

PROCESS {

    :folderloop foreach ($F in $InputFolder) {

        # check input object
        $NotFoundProps = Validate-Properties -InputObject $F -RequiredProperties $ReqProps # Validate-Properties is internal function
        if ($NotFoundProps) {
            Write-Error -Message ('Folder ' + $F.ToString() + ' is not proper object. Missing: ' + ($NotFoundProps -join ','))
            Continue # next foreach
        }
        $FolderPath = $F.FolderPath

        Write-Verbose -Message ('    Checking: '+($FolderPath))
        # check number of items
        if ($Filter) {
            $Items = $F.Items.Restrict($Filter)
        } else {
            $Items = $F.Items
        }
        $ItemsCount = $Items.Count
        $SubCount = $F.Folders.Count

        if ($ItemsCount -gt 0) {

            Write-Verbose -Message ('    Exporting ' + $FolderPath + ' , ' + $ItemsCount + ' message(s).') # may be fewer actual exported items to due to include/exclude type filter not considered yet, and also any errors such as missing properties
            # TODO Try foreach
            $msg = $Items.GetFirst()
            $itemCounter = 0
            $exportCounter = 0
            :itemloop do {
                if ($Progress) {Write-Progress -Activity ($FolderPath) -Status (' '+$msg.subject+' ') -PercentComplete (($itemCounter++)*100/$ItemsCount)}
                # TODO Add numbering of folders in Progress, like (1/5)
                if ((-not $IncludeTypes -or $msg.Class -in $IncludeTypes) -and (-not $ExcludeTypes -or $msg.Class -notin $ExcludeTypes)) {
                    if ($exportCounter -eq 0) {
                        # before first export, create folder container if needed
                        $TargetFolder = ($OutputFolderPath+$FolderPath.Replace('\\', '\')).Replace('\\', '\')
                        try {
                            New-Folder -TargetFolder $TargetFolder # internal commands
                        } catch {
                            Write-Error -Message $_
                            Continue folderloop # next folder
                        }
                    }
                     if ($Exportformat -eq 'MSG') {
                         Export-OutlookMessage -Messages $msg -OutputFolder $TargetFolder -FileNameFormat $FileNameFormat
                     } else { Export-OutlookMessageBody -Messages $msg -OutputFolder $TargetFolder -FileNameFormat $FileNameFormat -ExportFormat $ExportFormat}
                    ++$exportCounter
                } else {
                    Write-Verbose -Message ('Excluding message of type ' + [enum]::GetName([Microsoft.Office.Interop.Outlook.OlObjectClass], $msg.Class))
                }
                $msg = $Items.GetNext()
            } while ($msg)
            if ($Progress) {Write-Progress -Completed -Activity $F}
        }

        if ($SubCount -gt 0) {
            # export subfolders
            foreach ($subfolder in ($F.Folders)) {
                Export-OutlookFolder -InputFolder $subfolder -OutputFolder $OutputFolderPath -FileNameFormat $FileNameFormat -ExportFormat $ExportFormat -Filter $Filter -IncludeTypes $IncludeTypes -ExcludeTypes $ExcludeTypes -Progress:$Progress
            }
        }
    } # End of foreach

} # End of PROCESS block

END {

    Write-Verbose -Message 'Export-OutlookFolder completed.'

} # End of END block

# ------------------------- [End of function] ----------------------------

}
