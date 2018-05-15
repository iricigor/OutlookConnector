function Connect-Outlook {

<#
.SYNOPSIS
OutlookConnector function: Returns Outlook object that can be re-used for multiple consecutive calls in same session.

.DESCRIPTION
Returns Outlook COM namespace session object that can be re-used for multiple consecutive calls with other commands.
If you are using only one command, COM object will be created by calling function itself.
But, if you will be calling multiple functions, it is better to assign COM object to a variable, and re-use in consecutive calls.
In case of issues with connecting (usually due to slow or interactive Outlook start), use command Start-Outlook to open application GUI, and then re-try the command.

.INPUTS
This function is not accepting any parameters.

.OUTPUTS
Function returns COM object linked to Outlook application. Precisly, it returns Outlook MAPI namespace.
This COM object can be used in other module commands.

.EXAMPLE
$Outlook = Connect-Outlook
Returns COM object that is saved in variable $Outlook which can be used with other commands.

.LINK
about_OutlookConnector

.NOTES
NAME:       Connect-Outlook
AUTHOR:     Igor Iric, iricigor@gmail.com
CREATEDATE: September 29, 2015
#>
        
# based on function by Microsoft Scripting Guy, Ed Wilson 
# http://blogs.technet.com/b/heyscriptingguy/archive/2011/05/26/use-powershell-to-data-mine-your-outlook-inbox.aspx


# ---------------------- [Parameters definitions] ------------------------

[CmdletBinding()] 

Param ()

# ------------------------- [Function start] -----------------------------

# add types
# $olFolders = "Microsoft.Office.Interop.Outlook.OlDefaultFolders" -as [type]
# $olSaveAsTypes = "Microsoft.Office.Interop.Outlook.OlSaveAsType" -as [type]

# create new Outlook object and return it's MAPI namespace
try {
    Write-Verbose -Message 'Connecting to Outlook session'
    $outlook = new-object -comobject outlook.application
    $outlook.GetNameSpace("MAPI") # this is return object
    # MAPI Namespace https://msdn.microsoft.com/en-us/library/office/ff865800.aspx
    # Session https://msdn.microsoft.com/en-us/library/office/ff866436.aspx
    if ($outlook) {Write-Verbose -Message 'Connected successfully.'}
} catch {
    throw ('Can not obtain Outlook COM object. Try running Start-Outlook and then repeat command. '+($Error[0].Exception))
}

# ------------------------- [End of function] ----------------------------
        
}
