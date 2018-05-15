function Export-OutlookMessageBody {

<#
.SYNOPSIS
OutlookConnector function: Saves message body of Outlook message to a file on disk.

.DESCRIPTION
Saves body text of one or messages to a file on disk at specified path. All messages are saved in same folder, and file names are built based on customizable parameter FileNameFormat.
Messages can be obtained by one of Get-Outlook commands. Messages are saved in HTML, TXT or RTF format. If message with same name exists, numbering will be added at the end of file name.
Message body does not contain header information, i.e. info who sent message, when, etc. Also, it has no attachments.

.Example
(Get-OutlookInbox)[0] | Export-OutlookMessageBody -OutputFolder 'C:\tmp' -ExportFormat HTML
Saves body of first message in Inbox as HTML file

.EXAMPLE
Get-OutlookMessage -DefaultFolder DeletedItems | Export-OutlookMessageBody -OutputFolder 'C:\tmp\deleted' -ExportFormat TXT
Saves all messaged from Deleted items Outlook folder to a txt files on a disk

.PARAMETER Messages
Mandatory parameter which is content of one or more messages that will be saved to disk.
Messages can be obtained with one of Get-Outlook commands. Type Get-Command Get-Outlook* for list of commands.
Messages can be provided either as parameter, or via pipeline.

.PARAMETER OutputFolder
Mandatory parameter which specifies to which folder messages will be saved. It can be both local disk, as well as network location.

.PARAMETER FileNameFormat
Optional parameter that specifies how individual files will be named based. If omitted, files will be saved in format 'FROM= %SenderName% SUBJECT= %Subject%'.
File name can contain any of message parameters surrounded with %. For list of parameters, type Get-OutlookInbox | Get-Member.
Custom format can be specified after a | character within the %, e.g. %ReceivedTime|yyyyMMddhhmmss%.

.PARAMETER ExportFormat
Mandatory parameter which specifies to which format message body will be exported to. Allowed values are HTML, TXT (text) and RTF (rich-text).

.PARAMETER SkippedMessages
Optional parameter that specifies varaible to which will be stored messages that can not be processed.
Messages can be skipped for different reasons (wrong object, missing property specified in FileNameFormat parameter, etc.
Variable must be referenced, i.e. sent in format [ref]$Variable, and it must be declared in advance. Current value of variable will be deleted.

.OUTPUTS
Function is not returning any value.

.LINK
about_OutlookConnector

.NOTES
NAME:       Export-OutlookMessageBody
AUTHOR:     Igor Iric, IricIgor@GMail.com
CREATEDATE: November, 2015

#>

# ---------------------- [Parameters definitions] ------------------------

[cmdletbinding()]

Param(
    [parameter(Mandatory=$true,ValueFromPipeline=$true)] [ValidateNotNullOrEmpty()] [psobject[]]$Messages,
    [parameter(Mandatory=$true,ValueFromPipeline=$false)] [string]$OutputFolder,
    [parameter(Mandatory=$false,ValueFromPipeline=$false)] [string]$FileNameFormat='FROM= %SenderName% SUBJECT= %Subject%',
    [parameter(Mandatory=$true,ValueFromPipeline=$false)] [ValidateSet('HTML','TXT','RTF')] [string]$ExportFormat,
    [parameter(Mandatory=$false,ValueFromPipeline=$false)] [ref]$SkippedMessages

) #end param

# ------------------------- [Function start] -----------------------------

BEGIN {

    Write-Verbose -Message 'Export-OutlookMessageBody starting...'

    # convert format message to real file name, replace %...% with message attribute
    $ReqProps = @('Subject','HTMLBody','RTFBody','Body')
    $ReqProps += Get-Properties($FileNameFormat)

    # resolve relative path since MailItem.SaveAs does not support them
    $OutputFolderPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputFolder)

    # initialize queue for skipped messages, if it is passed
    if ($SkippedMessages) {
        $SkippedMessages.Value = @()
    }

} # End of BEGIN block

PROCESS {

    foreach ($Message in @($Messages)) {
    
        # check input object
        $NotFoundProps = Validate-Properties -InputObject $Message -RequiredProperties $ReqProps
        if ($NotFoundProps) {
            Report-MissingProperties -InputObject $Message -MissingProperties $NotFoundProps
            if ($SkippedMessages) {
                $SkippedMessages.Value += $Message # adding skipped messages to referenced variable if passed
            }
            Continue # next foreach
        }

        Write-Verbose -Message ('Processing '+($Message.Subject))

        # main code
        $FileName = Create-FileName -InputObject $Message -FileNameFormat $FileNameFormat   # Create-FileName is internal function

        # fix file name
        $FileName = Get-ValidFileName -FileName $FileName
        $FullFilePath = Add-Numbering -FileName (Join-Path -Path $OutputFolderPath -ChildPath $FileName) -FileExtension $ExportFormat
        Write-Verbose -Message "Saving message body to $FullFilePath"

        try{
            switch ($ExportFormat) {
                'HTML' {Set-Content -Value ($Message.HTMLBody) -LiteralPath $FullFilePath}
                'RTF'  {Set-Content -Value ($Message.RTFBody) -LiteralPath $FullFilePath -Encoding Byte}
                'TXT'  {Set-Content -Value ($Message.Body) -LiteralPath $FullFilePath}
            }
        } catch {
            if ($SkippedMessages) {
                $SkippedMessages.Value += $Message # adding skipped messages to referenced variable if passed
            }
            Write-Error -Message ('Message save exception. '+$Error[0].Exception)
        }

    } # End of foreach

} # End of PROCESS block

END {

    Write-Verbose -Message 'Export-OutlookMessageBody completed.'

} # End of END block

# ------------------------- [End of function] ----------------------------

}
