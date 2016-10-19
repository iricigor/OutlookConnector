#----------------------------------------------------[Function synopsis]-----------------------------------------------------

Function Convert-Credential {

<#
   .SYNOPSIS
    Converts saved credentials from old to new format. Part of CredentialsManager module.

   .DESCRIPTION
    Coinverts credential file(s) (User Name and Password) from old format (version 0.9) to new format (v.1.0).
    If you have not used old module, then there is no need for using this command.
    Old files are not deleted from disk.

    .EXAMPLE
    Convert-Credential -Environment Dev
    Converts Dev credentials file for usage with new module.

    .EXAMPLE
    Convert-Credential -Path P:\cred
    Converts credentials for all environments at specified path to new format.

    .PARAMETER Environment
    Optional. Name of environment for which credentials should be converted. It must be already saved with old Write-Credential function.
    It is also used as part of file names in Repository.
    If not specified, all credentials will be converted!

    .PARAMETER Path
    Optional. Defines folder location where credentials are saved on the disk.
    If not specified, %APPDATA%\CredentialsManager is used.

    .PARAMETER ListAvailable
    If specified, function will list all credentials at default or provided path. No conversion will occur!

    .OUTPUTS
    Function returns list of strings if ListAvailable is specified. Otherways (during conversion), it returns nothing.

    .LINK
    https://www.powershellgallery.com/packages/CredentialsManager

    .NOTES
    NAME:       Convert-Credential
    AUTHOR:     Igor Iric, IricIgor@Gmail.com
    CREATEDATE: October, 2016

 #>


#-------------------------------------------------[Parameters definitions]--------------------------------------------------

[cmdletbinding()]

Param(
  [parameter(Mandatory=$false,ValueFromPipeline=$true,ParameterSetName='Environments',Position=1)][string[]]$Environment,
  [parameter(Mandatory=$false,ValueFromPipeline=$false)][string]$Path=$Script:CredentialsPath,
  [parameter(Mandatory=$true,ValueFromPipeline=$false,ParameterSetName='Available')][switch]$ListAvailable
) #end param


#-------------------------------------------------[Function initialization]--------------------------------------------------
BEGIN {
    # function begin phase

    # process Path parameter, add default value, check if it exists
    if (!$Path) {
        # Path not provided, neither global is defined
        $Script:CredentialsPath = Join-Path -Path $env:APPDATA -ChildPath 'CredentialsManager'
        Write-Verbose -Message "Setting default path to: $Script:CredentialsPath"
        $Path = $Script:CredentialsPath
        } 

    if (!(Test-Path -Path $Path)) {
        Throw "Folder $Path is not existing. Please provide another path."
    } elseif ($Path -ne $Script:CredentialsPath) {
        # Path correct and specified directly, it will be saved for next usage
        Write-Verbose -Message "Updating default path to $Path."
        $Script:CredentialsPath = $Path
    } 

    # process ListAvailable
    if ($ListAvailable -or (!$Environment)) {
        Write-Verbose -Message 'Obtaining list of available environments'

        $UserNames = (Get-ChildItem -Path $Path -Filter '*_UserName.cred') -replace '_UserName.cred',''
        $PassNames = (Get-ChildItem -Path $Path -Filter '*_Password.cred') -replace '_Password.cred',''
        if ($UserNames -and $PassNames) {
            $Environments = (Compare-Object -ReferenceObject $UserNames -DifferenceObject $PassNames -ExcludeDifferent -IncludeEqual).InputObject
        }

        if (!$Environments) {
            Write-Warning -Message 'No stored credentials found. Please provide another path.'
        } else {
            Write-Verbose -Message ('Obtained '+($Environments.Count)+' environments.')
            if ($ListAvailable) {$Environments} # output object
            else {$Environment = $Environments} # passing for PROCESS section
        } 
    }
}

#---------------------------------------------------[Function processing]----------------------------------------------------
PROCESS {

    foreach ($E in $Environment) {
        # function process phase, executed once for each element in main Prameter
        Write-Verbose -Message '----------------------'
        Write-Verbose -Message "Processing environment $E..."

        $FileName = Join-Path -Path $Path -ChildPath ($E + '.creds')
        $FileNameUser = Join-Path -Path $Path -ChildPath ($E + '_UserName.cred')
        $FileNamePass = Join-Path -Path $Path -ChildPath ($E + '_Password.cred')
        $VerboseMessage = "Environment $E processed with issues." # it will be updated at the end, if successfull

        if ((!(Test-Path -Path $FileNameUser)) -or  (!(Test-Path -Path $FileNamePass))) {
            Write-Error -Message "Credentials files not found."
        } else {
            Write-Verbose -Message 'Credentials files found.'

            $EncUser = Get-Content -Path $FileNameUser
            $EncPass = Get-Content -Path $FileNamePass

            # if file is existing, delete it
            if (Test-Path -Path $FileName) {
                try {
                    Write-Verbose -Message "Trying to delete existing file(s) for $E..."
                    Remove-Item -Path $FileName -ErrorAction Stop
                    Write-Verbose -Message 'File deleted.'
                } catch {
                    Write-Error -Message ('Credential file can not be updated. '+($Error[0].Exception))
                }
            }

            # if file is not existing, export values
            if (!(Test-Path -Path $FileName)) {
                try {
                    Write-Verbose -Message "Attempting to save credentials to $FileName..."
                    $EncUser | Out-File $FileName
                    $EncPass | Out-File $FileName -Append
                    $VerboseMessage = "Environment $E processed successfully."
                    Write-Verbose -Message 'Credentials saved.'
                } catch {
                    Write-Error -Message ('Credential files can not be updated. '+($Error[0].Exception))
                }
            }


        Write-Verbose -Message $VerboseMessage
        }
    } # end of foreach $Environment

} # end of PROCESS block

#-----------------------------------------------------[Function closing]-----------------------------------------------------

END {
    # function closing phase
    Write-Verbose -Message 'Read-Credential finishing.'

}

} # end of function code
#----------------------------------------------------[End of function]------------------------------------------------------

#---------------------------------------------------[Comments section]------------------------------------------------------
