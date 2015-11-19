function Get-OutlookMessage {

<#
.SYNOPSIS
OutlookConnector function: Returns array of messages from one of default folders in Outlook.

.DESCRIPTION
Returns array of messages from one of default folders, i.e. Inbox, Drafts, SentMail, etc.

.EXAMPLE
Get-OutlookMessage SentMail | Group To | Sort Count -Descending | Select -First 2
Find out to whom you are sending the most messages

.EXAMPLE
Get-OutlookMessage -ListAvailableFolders
Lists all accepted names for folders. Names 

.PARAMETER DefaultFolder
Mandatory parameter which specifies names of default folders from which messages can be obtained. Names are builtin Outlook names, i.e. they do not have to be the same as displayed names in outlook.
For example, default name Inbox will always correspond to folder which is receiving new messages, regardles if it is renamed on the system.
Names can be passed via pipeline.

.PARAMETER Outlook
Optional parameter that specifies Outlook session from whch it will obtain needed data.
If omitted, function will connect automatically using Connect-Outlook function.

.PARAMETER ListAvailableFolders
If switch ListAvailableFolders is used, function will list all default folder names that can be used as parameter DefaultFolder.

.OUTPUTS
Function returns array of messages.

.LINK
about_OutlookConnector

.NOTES
NAME:       Get-OutlookMessage
AUTHOR:     Igor Iric, iricigor@gmail.com
CREATEDATE: September 29, 2015
#>


    # ---------------------- [Parameters definitions] ------------------------

    [CmdletBinding()]

    Param(
        [parameter(ParameterSetName='Messages',Mandatory=$true,ValueFromPipeline=$true,Position=0)][string[]]$DefaultFolder,
        [parameter(ParameterSetName='Messages',ValueFromPipeline=$false)][psobject]$Outlook = (Connect-Outlook),
        [parameter(ParameterSetName='FolderNames',Mandatory=$true)][switch]$ListAvailableFolders

    ) #end param

    # ------------------------- [Function start] -----------------------------

    BEGIN {
        Write-Verbose -Message 'Get-OutlookMessage obtaining type information'
        $olFolders = "Microsoft.Office.Interop.Outlook.olDefaultFolders" -as [type]
        $KeyWord = 'olFolder'
        try {
            $AllFolders = ($olFolders.GetEnumNames() | Where-Object {$_ -match "^$KeyWord"}) -replace $KeyWord,''
            if (@($AllFolders).Count -lt 2) {
                throw 'Error obtaining default folder names.'
            }
        } catch {
            throw 'Error obtaining default folder names.'
        }
                

        if ($ListAvailableFolders) {
            Write-Verbose -Message 'Listing all input values'
            $AllFolders
            $DefaultFolder = @() # avoid processing below
        }
        
    }

    PROCESS {
        
        foreach ($F in $DefaultFolder) {

            Write-Verbose -Message "Processing $F"
            if ($F -in $AllFolders) {
                $FullName = $KeyWord + ($AllFolders | Where-Object {$_ -eq $F}) # getting proper capitalization and full name
                $FolderDef = $Outlook.GetDefaultFolder([Microsoft.Office.Interop.Outlook.olDefaultFolders]$FullName)
                # return value
                $FolderDef.Items
                Write-Verbose -Message ('Returned '+(@($FolderDef.Items).Count)+' messages.')
            } else {
                Write-Error -Message "Folder with name $F is not found. Use ListAvailableFolders parameter to get all the options."
            }
        }
    }


    END {

    # function closing phase
    } 


    # ------------------------- [End of function] ----------------------------
}
