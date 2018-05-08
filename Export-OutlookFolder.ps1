function Export-OutlookFolder {

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
Get-OutlookFolder -MainOnly | Export-OutlookFolder -OutputFolder 'C:\tmp\all' -Progress -WarningAction SilentlyContinue
Saves all messages from main mailbox 'C:\tmp\all' to a disk using default file naming. Warnings for messages without subject used in file naming, are ignored.

.PARAMETER InputFolder
Mandatory parameter that specifies which Outlook folder needs to be exported. Easies is to obtain it via 

.PARAMETER OutputFolder
Mandatory parameter which specifies to which folder messages will be saved. It can be both local disk, as well as network location.
If folder is not existing, it will be created.
Entire Outlok folder structure will be generated within specified folder.

.PARAMETER FileNameFormat
Optional parameter that specifies how individual files will be named based. If omitted, files will be saved in format 'FROM= %SenderName% SUBJECT= %Subject%'.
File name can contain any of message parameters surrounded with %. For list of parameters, type Get-OutlookInbox | Get-Member.
Parameter is passed to Export-OutlookMessage function.

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
#>

    # ---------------------- [Parameters definitions] ------------------------

    [CmdletBinding()]

    Param(
        [parameter(Mandatory=$true,ValueFromPipeline=$true,Position=0)][psobject[]]$InputFolder,
        [parameter(Mandatory=$true,ValueFromPipeline=$false)][string]$OutputFolder,
        [parameter(Mandatory=$false,ValueFromPipeline=$false)][string]$FileNameFormat='FROM= %SenderName% SUBJECT= %Subject%',
        [switch]$Progress

    ) #end param

    # ------------------------- [Function start] -----------------------------

    BEGIN {
        Write-Verbose -Message 'Export-OutlookFolder starting...'
        $ReqProps = @('Items','FullFolderPath','Folders')
        $OutputFolderPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputFolder)
    }

    PROCESS {

        foreach ($F in $InputFolder) {

            # check input object
            $NotFoundProps = Validate-Properties -InputObject $F -RequiredProperties $ReqProps # Validate-Properties is internal function
            if ($NotFoundProps) {
                Write-Error -Message ('Folder ' + $F.ToString() + ' is not proper object. Missing: ' + ($NotFoundProps -join ','))
                Continue # next foreach
                }

            Write-Verbose -Message ('    Checking: '+($F.FolderPath))
            # check number of items
            $MsgCount = $F.Items.Count
            $SubCount = $F.Folders.Count

            if ($MsgCount -gt 0) {

                # if needed, create folder container
                $TargetFolder = $OutputFolderPath+((($F.FolderPath) -replace '\\\\','\')) -replace '\\\\','\'
                New-Folder -TargetFolder $TargetFolder # internal commands
                Write-Verbose -Message ('    Exporting'+$F.FolderPath+', '+$MsgCount+' message(s).')
                $messages = $F.Items
                # TODO Try foreach
                $msg = $messages.GetFirst()
                $i = 0
                do {
                    if ($Progress) {Write-Progress -Activity ($F.FolderPath) -Status (' '+$msg.subject+' ') -PercentComplete (($i++)*100/$MsgCount)}
                    # TODO Add numbering of folders in Progress, like (1/5)
                    Export-OutlookMessage -Messages $msg -OutputFolder $TargetFolder -FileNameFormat $FileNameFormat
                    $msg = $messages.GetNext()
                } while ($msg)
                if ($Progress) {Write-Progress -Completed -Activity $F}
            }

            if ($SubCount -gt 0) {
                # export subfolders
                foreach ($subfolder in ($F.Folders)) {
                    Export-OutlookFolder -InputFolder $subfolder -OutputFolder $OutputFolderPath -FileNameFormat $FileNameFormat -Progress:$Progress
                }
            }
        } # end foreach
    } # end PROCESS

    END {
        Write-Verbose -Message 'Export-OutlookFolder finished.'
    }


    # ------------------------- [End of function] ----------------------------
}