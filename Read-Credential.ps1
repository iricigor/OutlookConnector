#----------------------------------------------------[Function synopsis]-----------------------------------------------------

Function Read-Credential {

<#
   .SYNOPSIS
    Reads credentials from disk, saved with Write-Credential. Part of CredentialsManager module.

   .DESCRIPTION
    Reads credentials (User Name and Password) from disk, saved with Write-Credential.
    Credentials are saved as encrypted data. Encryption is done with current user identifier, so only current user can decrypt saved credentials.
    Other users cannot decrypt content, even if they get access to credential files (i.e. saved on network share).

    .EXAMPLE
    $Cred = Read-Credential -Environment Dev
    Gets username and password from files saved in repository and saves it in variable.

    .EXAMPLE
    Get-WMIObject Win32_Service -Computer '58.138.249.212' -Credential (Read-Credential -Environment Dev)
    Gets credentials from repository and directly uses it in network operation.

    .EXAMPLE
    Get-Item '\\dfs\restricted\*' -Credential (Read-Credential -Environment BAcc)
    Gets B account credentials from repository and directly uses it to access restricted folder content.

    .PARAMETER Environment
    Name of environment for which credentials should be obtained. It must be already saved with Write-Credential. 
    It is also used as part of file names in Repository.

    .PARAMETER Path
    Optional parameter which defines where credentials are saved on the disk.
    If not specified, %APPDATA%\CredentialsManager is used.
    If specified, default value will be updated and re-used on next calls within the same PowerShell session.

    .PARAMETER ListAvailable
    If specified, function will list all credentials at default or provided path. Returned object(s) will have added property Environment which is named according to Write-Credential function.

    .OUTPUTS
    Function returns Credentials object the same as Get-Credential, System.Management.Automation.PSCredential.

    .LINK
    https://www.powershellgallery.com/packages/CredentialsManager

    .NOTES
    NAME:       Read-Credential
    AUTHOR:     Igor Iric, IricIgor@Gmail.com
    CREATEDATE: October, 2015

 #>


#-------------------------------------------------[Parameters definitions]--------------------------------------------------

[cmdletbinding()]

Param(
  [parameter(Mandatory=$true,ValueFromPipeline=$true,ParameterSetName='Environments',Position=1)][string[]]$Environment,
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
        Throw "Folder $Path is not existing. Please provide another path, or run Write-Credential first."
    } elseif ($Path -ne $Script:CredentialsPath) {
        # Path correct and specified directly, it will be saved for next usage
        Write-Verbose -Message "Updating default path to $Path."
        $Script:CredentialsPath = $Path
    } 

    # process ListAvailable
    if ($ListAvailable) {
        Write-Verbose -Message 'Obtaining list of available environments'
        <#
        $UserNames = (Get-ChildItem -Path $Path -Filter '*_UserName.cred') -replace '_UserName.cred',''
        $PassNames = (Get-ChildItem -Path $Path -Filter '*_Password.cred') -replace '_Password.cred',''
        if ($UserNames -and $PassNames) {
            $Environment = (Compare-Object -ReferenceObject $UserNames -DifferenceObject $PassNames -ExcludeDifferent -IncludeEqual).InputObject
        }
        #>
        $Environment = (Get-ChildItem -Path $Path -Filter '*.creds') -replace '.creds',''

        if (!$Environment) {
            Write-Warning -Message 'No stored credentials found. Try using Write-Credential, or Read-Credential with Path and ListAvailable parameters.'
        } else {
            Write-Verbose -Message ('Obtained '+($Environment.Count)+' environments.')
        } 
    }
}

#---------------------------------------------------[Function processing]----------------------------------------------------
PROCESS {

    foreach ($E in @($Environment)) {
        # function process phase, executed once for each element in main Prameter
        Write-Verbose -Message '----------------------'
        Write-Verbose -Message "Processing environment $E..."
        $VerboseMessage = "Environment $E processed with issues." # it will be updated at the end, if successfull
        $Cred = $null

        # main code
        #$FileNameUser = Join-Path -Path $Path -ChildPath ($E + '_UserName.cred')
        #$FileNamePass = Join-Path -Path $Path -ChildPath ($E + '_Password.cred')
        $FileName = Join-Path -Path $Path -ChildPath ($E + '.creds')

        #if ((!(Test-Path -Path $FileNameUser)) -or  (!(Test-Path -Path $FileNamePass))) {
        if (!(Test-Path -Path $FileName)) {
            Write-Error -Message "Credentials file not found."
        } else {
            Write-Verbose -Message 'Credentials file found.'

            $Content = Get-Content -Path $FileName
            if (($Content.Count) -ne 2) {
                Write-Error -Message "Credentials file not in proper format."
            } else {
            
                try {           
                    Write-Verbose -Message 'Attempting to decrypt data.'
                    #$UserSec = Get-Content -Path $FileNameUser | ConvertTo-SecureString
                    #$PassSec = Get-Content -Path $FileNamePass | ConvertTo-SecureString
                    $UserSec = $Content[0] | ConvertTo-SecureString
                    $PassSec = $Content[1] | ConvertTo-SecureString
                    $UserPlain = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($UserSec))

                    $Cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $UserPlain,$PassSec
                    Write-Verbose -Message 'Data encryption completed.'
                    $VerboseMessage = "Environment $E processed successfully."
                } catch {
                    if (($Error[0].Exception.Message) -eq 'Key not valid for use in specified state.') {
                        Write-Error 'Data encryption failed. Only user account that saved data can decrypt it.'
                    } else {
                        Write-Error -Message ('Data encryption failed. '+$Error[0].Exception.Message)
                    }
                }
            }
        }

        # return values
        if ($Cred) {
            if ($ListAvailable) {
                Write-Verbose -Message 'Adding Environment property.'
                $Cred | Add-Member -MemberType NoteProperty -Name Environment -Value $E
            }
            Write-Verbose -Message "Returning $E credentials value to output."
            $Cred
        }

        Write-Verbose -Message $VerboseMessage

    } # end of foreach

} # end of function

#-----------------------------------------------------[Function closing]-----------------------------------------------------

END {
    # function closing phase
    Write-Verbose -Message 'Read-Credential finishing.'

}

} # end of function code
#----------------------------------------------------[End of function]------------------------------------------------------

#---------------------------------------------------[Comments section]------------------------------------------------------
