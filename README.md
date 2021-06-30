# Outlook Connector
Microsoft Outlook connector module, version 0.91

author: Igor Iric, iricigor@gmail.com
- - - - - - - - - - - - - - - - - - - - - - - - - - - 

The module **Outlook Connector** will enable you to connect to MS Outlook session on your computer via few simple to use command line functions. You can pipe array of messages to PowerShell commands and do all the scripting with provided data (grouping, examining, logging, etc.). Or, you can export all or some messages to file system, which is functionality not provided in Outlook itself!

## List of commands
```powershell
Connect-Outlook
Start-Outlook
Get-OutlookInbox 
Get-OutlookMessage
Get-OutlookFolder 
Export-OutlookMessage
Export-OutlookMessageBody
Export-OutlookFolder
```

All commands have documented help system. Type *Get-Help about_OutlookConnector*, or *Get-Command -Module OutlookConnector* for more info.



## Commands overview
- - - - - - - - -


### Connecting to Outlook
- - - - - - - - - - -
- **Connect-Outlook**      - creates Outlook instance in memory, returns MAPI instance
- **Start-Outlook**        - starts MS Outlook application GUI, usefull for troubleshooting


### Getting messages to memory
- - - - - - - - - - - - - -
- **Get-OutlookInbox**    [-Outlook]                 - returns array of messages in default Inbox folder from Outlook instance, based on the Scripting Guy
- **Get-OutlookFolder**   [-Outlook] [–Recurse]      - lists all folder (and optionally subfolders) inside Outlook instance, returns array of Outlook folders
- **Get-OutlookMessage**  –DefaultFolder [-Outlook]  - returns messages in one of default folders based on the name (i.e SentItems, Drafts, etc.)
[-Outlook] - if Outlook session is not specified, commands will automatically connect

Get functions are returning all properties of a message, so it's good practice to select only properties you need before output to screen. To list all properties of a message type for example Get-OutlookInbox | Get-Member.


### Saving messages to disk
- - - - - - - - - - - - 
- **Export-OutlookFolder**      –InputFolder -OutputFolder -FilenameFormat [-ExportFormat]  - saves all messages to folder on a disk, in one of 4 formats MSG, HTML, RTF or TXT
- **Export-OutlookMessage**     –Message     –OutputFolder -FileNameFormat   - saves individual message to folder on a disk
Input parameter (folder or message) can be piped. Export functions are saving messages in individual MSG files.


## Examples
- - - -
```powershell
Get-OutlookFolder | Export-OutlookFolder 'C:\Email'   # saves all emails to disk
Get-OutlookFolder -recurse  | WHERE Name -in 'Home','Finance'   | Export-OutlookFolder -OutputFolder 'C:\Export-Mail' -ExportFormat HTML # saves messages from two folders to disk in HTML format
Get-OutlookInbox -Verbose | Group Sendername          # group inbox messages by sendername
Get-Help Export-OutlookMessage -Examples              # all messages have standard synopsis
```

## Version history
- - - - - - - - 
  - **0.90** - Sep '15 - initial version, 7 functions, read only access to data
  - **0.91** - Oct '15 - split to multiple files, separate module and manifest file, more help and verbose
  - **0.92** - Nov '15 - 2nd public release, one more command added (Export Body), corrected email address
  - **0.93** - May '20 - Added feature to allow Export-OutlookFolder to accept ExportFormat parameter in order to handle all formats


## Requests for next versions
- - - - - - - - - - - - - 
- Export-OutlookFolder - add failed messages to error variable; this functionality is implemented in other Export-Outlook* functions

***Any further suggestion is welcome!***

**PetRose**: Found that it was difficult perhaps even impossible, to pass Outvariable from one of the functions to the other using pipes.
I needed this 'logically equivalent' capability of this: 
      Get-OutlookFolder -recurse -Progress | ? Name -eq 'Home'  | Export-OutlookMessageBody -OutputFolder 'C:\temp\Export-Mail' -ExportFormat HTML
but this throws error: 
  Export-OutlookMessageBody : Message of type Folder is not proper object. Missing: Subject, HTMLBody, RTFBody, Body, SenderName, Subject.

So, I took the opportunity to enhance the Export-OutlookFolder function to accept the -ExportFormat parameter and then simply pass its value to either:
  Export-OutlookMessage (in the case of MSG format)
or
  Export-OutlookMessageBody (in all other cases)

Its tested, and seems to work as version 0.93