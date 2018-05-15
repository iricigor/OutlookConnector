function Export-OutlookMessage {

<#
.SYNOPSIS
OutlookConnector function: Saves Outlook message to a file on disk.

.DESCRIPTION
Saves one or messages to a file on disk at specified path. All messages are saved in same folder, and file names are built based on customizable parameter FileNameFormat.
Messages can be obtained by one of Get-Outlook commands. Messages are saved in MSG format. If message with same name exists, numbering will be added at the end of file name.

.EXAMPLE
(Get-OutlookInbox | sort receivedtime -Descending)[0] | Export-OutlookMessage -OutputFolder 'C:\tmp'
Save last message from Inbox to C:\tmp folder

.EXAMPLE
Get-OutlookMessage Inbox | ? Sendername -Match 'boss' | Export-OutlookMessage -OutputFolder 'C:\boss'
Get-OutlookMessage SentMail | ? To -Match 'boss' | Export-OutlookMessage -OutputFolder 'C:\boss'
Save your entire communication with boss to a folder on a disk.

.PARAMETER Messages
Mandatory parameter which is content of one or more messages that will be saved to disk.
Messages can be obtained with one of Get-Outlook commands. Type Get-Command Get-Outlook* for list of commands.
Messages can be provided either as parameter, or via pipeline.

.PARAMETER OutputFolder
Mandatory parameter which specifies to which folder messages will be saved. It can be both local disk, as well as network location. It must exist.

.PARAMETER FileNameFormat
Optional parameter that specifies how individual files will be named based. If omitted, files will be saved in format 'FROM= %SenderName% SUBJECT= %Subject%'.
File name can contain any of message parameters surrounded with %. For list of parameters, type Get-OutlookInbox | Get-Member.
Custom format can be specified after a | character within the %, e.g. %ReceivedTime|yyyyMMddhhmmss%.

.PARAMETER SkippedMessages
Optional parameter that specifies varaible to which will be stored messages that can not be processed.
Messages can be skipped for different reasons (wrong object, missing property specified in FileNameFormat parameter, etc.
Variable must be referenced, i.e. sent in format [ref]$Variable, and it must be declared in advance. Current value of variable will be deleted.

.OUTPUTS
Function is not returning any value.

.LINK
about_OutlookConnector

.NOTES
NAME:       Export-OutlookMessage
AUTHOR:     Igor Iric, iricigor@gmail.com
CREATEDATE: September 29, 2015
#>

# ---------------------- [Parameters definitions] ------------------------

[CmdletBinding()]

Param(
    [parameter(Mandatory=$true,ValueFromPipeline=$true)] [ValidateNotNullOrEmpty()] [psobject[]]$Messages,
    [parameter(Mandatory=$true,ValueFromPipeline=$false)] [string]$OutputFolder,
    [parameter(Mandatory=$false,ValueFromPipeline=$false)] [string]$FileNameFormat='FROM= %SenderName% SUBJECT= %Subject%',
    [parameter(Mandatory=$false,ValueFromPipeline=$false)] [ref]$SkippedMessages

) #end param

# ------------------------- [Function start] -----------------------------

BEGIN {

    Write-Verbose -Message 'Export-OutlookMessage starting...'
    $olSaveAsTypes = "Microsoft.Office.Interop.Outlook.OlSaveAsType" -as [type]

    # convert format message to real file name, replace %...% with message attribute
    $ReqProps = @('Subject','SaveAs')
    $ReqProps += Get-Properties($FileNameFormat)

    # resolve relative path since MailItem.SaveAs does not support them
    $OutputFolderPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputFolder)

    # initialize queue for skipped messages, if it is passed
    if ($SkippedMessages) {
        $SkippedMessages.Value = @()
    }

} # End of BEGIN block

PROCESS {

    foreach ($Message in $Messages) {

        # check input object
        try {
            Validate-Properties -InputObject $Message -RequiredProperties $ReqProps
        } catch {
            if ($SkippedMessages) {
                $SkippedMessages.Value += $Message # adding skipped messages to referenced variable if passed
            }
            Write-Error $_
            Continue # next foreach
        }

        Write-Verbose -Message ('Processing '+($Message.Subject))

        # create base file name
        $FileName = Create-FileName -InputObject $Message -FileNameFormat $FileNameFormat   # Create-FileName is internal function

        # fix file name
        $FileName = Get-ValidFileName -FileName $FileName
        try {
            $FullFilePath = Get-UniqueFilePath -FolderPath $OutputFolderPath -FileName $FileName -Extension 'msg'
        } catch {
            Write-Error $_
            Continue # next foreach
        }

        # save message to disk
        Write-Verbose -Message "Saving message to $FullFilePath"
        try {
            $Message.SaveAs($FullFilePath,$olSaveAsTypes::olMSGUnicode)
        } catch {
            if ($SkippedMessages) {
                $SkippedMessages.Value += $Message # adding skipped messages to referenced variable if passed
            }
            Write-Error -Message ('Message save exception.'+$Error[0].Exception)
        }

    } # End of foreach

} # End of PROCESS block

END {

    Write-Verbose -Message 'Export-OutlookMessage completed.'

} # End of END block

# ------------------------- [End of function] ----------------------------

}
