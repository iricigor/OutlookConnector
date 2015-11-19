function Start-Outlook {

<#
.SYNOPSIS
OutlookConnector function: Starts MS Outlook GUI. Used in case of issues with Connect-Outlook function.

.DESCRIPTION
Starts MS Outlook GUI. Used in case of issues with Connect-Outlook function.
Issues usually occure due to slow or interactive Outlook start. If GUI starts successfully, Connect-Outlook should run without issues.
Outlook executable is found via Registry search.

.INPUTS
This function is not accepting any parameters.

.OUTPUTS
Function returns nothing.

.EXAMPLE
Start-Outlook
Opens MS Outlook GUI.

.LINK
about_OutlookConnector

.NOTES
NAME:       Start-Outlook
AUTHOR:     Igor Iric, iricigor@gmail.com
CREATEDATE: September 29, 2015
#>

# ---------------------- [Parameters definitions] ------------------------

[CmdletBinding()] 

Param ()

# ------------------------- [Function start] -----------------------------

if (Get-Process | Where-Object name -eq outlook) {
    Write-Verbose -Message 'Outlook already running. No action needed.'
    }
else {
    $key = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE\'

    if (!(Test-Path -Path $key)) {
        throw 'Path to Outlook executable not found.'

    } else {
        $exe = (Get-ItemProperty -Path $key).'(default)'
        if (Test-Path -Path $exe) {
            Write-Verbose -Message 'Starting Outlook application...'
            Invoke-Item -Path  $exe
        } else {
            throw 'Outlook executable not found.'
        }
    }
}

# ------------------------- [End of function] ----------------------------

}
