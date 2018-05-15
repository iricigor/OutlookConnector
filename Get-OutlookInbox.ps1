function Get-OutlookInbox {

<#
.SYNOPSIS
OutlookConnector function: Returns array of messages from default Inbox folder.

.DESCRIPTION
Returns array of messages from default Inbox folder. If Outlook session is not passed as parameter Outlook, function will connect automatically.
Messages are not filtered in any way, i.e. all messages and all properties are returned. Filtering can be applied via piping to Where-Object.
For list of properties of each message, pipe result to Get-Member.
This is special case of function Get-OutlookMessage Inbox

.EXAMPLE
Get-OutlookInbox | ? SenderName -match 'Microsoft' | Select Subject,ReceivedTime,UnRead
Lists all Inbox messages sent by Microsoft

.EXAMPLE
Write-Host 'You have:' ((Get-OutlookInbox | ? Unread).Count) 'unread messages in inbox.'
Writes information about number of unread messages in Inbox.

.EXAMPLE
Get-OutlookInbox | Select SenderName,Subject,SentOn,Size | Sort SentOn -Descending | Out-GridView
Display grid view window with columns similar to default Outlook columns

.EXAMPLE
Get-OutlookInbox | Group SenderName | Select Count,Name | ? Count -gt 1
Displays from whom you have more than 1 email in Inbox

.EXAMPLE
Get-OutlookInbox | % {($_.ReceivedTime).DayOfWeek} | Group
Find out from which weekday you are having the most messages

.PARAMETER Outlook
Optional parameter that specifies Outlook session from whch it will obtain needed data.
If omitted, function will connect automatically using Connect-Outlook function.

.OUTPUTS
Function returns array of messages.

.LINK
about_OutlookConnector

.NOTES
NAME:       Get-OutlookInbox
AUTHOR:     Igor Iric, iricigor@gmail.com
CREATEDATE: September 29, 2015
#>

# based on function by Microsoft Scripting Guy, Ed Wilson 
# http://blogs.technet.com/b/heyscriptingguy/archive/2011/05/26/use-powershell-to-data-mine-your-outlook-inbox.aspx

# ---------------------- [Parameters definitions] ------------------------

[CmdletBinding()]  

Param(
    [parameter(Mandatory=$false,ValueFromPipeline=$false)]$Outlook = (Connect-Outlook)
) #end param

# ------------------------- [Function start] -----------------------------

try {
    Write-Verbose -Message 'Obtaining messages from outlook Inbox'
    $olFolders = "Microsoft.Office.Interop.Outlook.OlDefaultFolders" -as [type]
    $InboxDef = $Outlook.GetDefaultFolder($olFolders::olFolderInbox)
    if (!$InboxDef) {
        throw ('Obtaining inbox definition failed. '+($Error[0].Exception))
    } else {
        # return values
        $Inboxdef.Items
        Write-Verbose -Message ('Successfully obtained '+ ($Inboxdef.Items.Count) +' messages from outlook Inbox')
    }
} catch {
    throw ('Obtaining inbox messages failed. '+($Error[0].Exception))
}

# ------------------------- [End of function] ----------------------------

}
