
#----------------------------------------------------[Function synopsis]-----------------------------------------------------

Function Write-Credential {

<#
    .SYNOPSIS
    Saves credentials to disk system as secure string for re-use. Part of CredentialsManager module.

    .DESCRIPTION
    Saves credentials to disk system as secure string for re-use.
    User will be prompted to enter user name and password which will be saved. Username will be re-used exactly as specified, i.e. depending on target use you may want to specify domain\user or user@domain.com
    Only currently logged in user will be able to decrypt and use saved credentials. Other users even if they will have access to saved files, will not be able to decrypt it.

    .EXAMPLE
    Write-Credential Dev
    User is prompted for credentials (user name and password) which are saved to default location.

    .EXAMPLE
    Write-Credential BAcc -Path P:\cred
    User is prompted for credentials for B account, which are afterwards saved to P: disk.

    .EXAMPLE
    Write-Credential Azure -Credential $AzureAdmin
    Saving credentials from variable $AzureAdmin to file Azure.cred on the disk.

    .PARAMETER Environment
    Name of environment for which credentials are saved. It will be used in Get commands, and also as part of file names.
    If you want to save credentials for multiple users in same environment, use different Environment names, like EnvAdmin and EnvUser. 
    You may specify multiple environments.

    .PARAMETER Path
    Optional parameter which defines where credentials will be saved. If not specified, %APPDATA%\CredentialsManager is used.
    If specified, default value will be updated and re-used on next calls within the same PowerShell session.

    .PARAMETER PassThru
    Switch parameter which defines if credentials created by this cmdlet will be passed through the pipeline. Default is false. 

    .PARAMETER Credential
    By default user will be prompted to enter credentials interactivly. If you already have credentials in variable, then specify this parameter.

    .OUTPUTS
    Cmdlet is not returning any value by default.
    If switch PassThru is specified, it returns Credentials object as Get-Credential, System.Management.Automation.PSCredential.

    .LINK
    https://www.powershellgallery.com/packages/CredentialsManager

    .NOTES
    NAME:       Write-Credential
    AUTHOR:     Igor Iric, IricIgor@Gmail.com
    CREATEDATE: October, 2015

 #>


#-------------------------------------------------[Parameters definitions]--------------------------------------------------

[cmdletbinding()]

Param(
  [parameter(Mandatory=$true,ValueFromPipeline=$true)][string[]]$Environment,
  [parameter(Mandatory=$false,ValueFromPipeline=$false)][string]$Path=$Script:CredentialsPath,
  [parameter(Mandatory=$false,ValueFromPipeline=$false)][System.Management.Automation.PSCredential]$Credential,
  [switch]$PassThru
) #end param

#-------------------------------------------------[Constant declarations]---------------------------------------------------




#-------------------------------------------------[Function initialization]--------------------------------------------------
BEGIN {
    # function begin phase
    # handling Path parameter, if folder not exisiting create it and set hidden

    
    if (!$Path) {
        # if path is not specified, it means also global needs to be initialized
        $Script:CredentialsPath = Join-Path -Path $env:APPDATA  -ChildPath 'CredentialsManager'
        Write-Verbose -Message "Setting default path to: $Script:CredentialsPath"
        $Path = $Script:CredentialsPath
        Write-Verbose -Message "Using default path $Path"
    }

    if (!(Test-Path -Path $Path)) {
        # folder not existing -> create it
        try {
            Write-Verbose -Message "Trying to create exports folder $Path"
            $Folder = New-Item -Path $Path -ErrorAction Stop -ItemType Directory
            $PathFolderCreated = $true
        } catch {
            # folder creation error
            $PathFolderCreated = $false
            throw $Error[0]
        }

        if ($PathFolderCreated) {
            Write-Verbose -Message 'Exports folder created.'
            # set folder to be hidden, if creating it
            $Folder.Attributes = $Folder.Attributes -bor [io.fileattributes]::Hidden
            if ($Path -ne $Script:CredentialsPath) {
                Write-Verbose -Message "Updating default path to $Path."
                $Script:CredentialsPath = $Path
            }               
        }
    } # end of folder creation part
}

#---------------------------------------------------[Function processing]----------------------------------------------------
PROCESS {

    foreach ($E in @($Environment)) {
    # function process phase, executed once for each element in main Parameter
    
        # main code
        Write-Verbose -Message '----------------------'
        Write-Verbose -Message "Processing environment $E..."
        $VerboseMessage = "Environment $E processed with issues." # it will be updated at the end, if successfull

        # prompt user for credentials
        if (!($Credential)) {
            $Cred = Get-Credential -Message "Enter your credentials for $E..."
        } else {
            $Cred = $Credential
        }

        if (!$Cred) {
            Write-Error -Message 'Credentials not obtained. Ensure you are not clicking on Cancel.'
        } elseif ($Cred.Password.Length -eq 0) {
            Write-Error -Message 'Blank passwords not supported.'
        } else {
            # if credentials are obtained, proceed
            Write-Verbose -Message "Credentials for $E obtained."
            #$FileNameUser = Join-Path -Path $Path -ChildPath ($E + '_UserName.cred')
            #$FileNamePass = Join-Path -Path $Path -ChildPath ($E + '_Password.cred')
            $FileName = Join-Path -Path $Path -ChildPath ($E + '.creds')

            # if file is existing, delete it
            try {
                Write-Verbose -Message "Trying to delete existing file(s) for $E..."
                #if (Test-Path -Path $FileNameUser) {Remove-Item -Path $FileNameUser -ErrorAction Stop}
                #if (Test-Path -Path $FileNamePass) {Remove-Item -Path $FileNamePass -ErrorAction Stop}
                if (Test-Path -Path $FileName) {Remove-Item -Path $FileName -ErrorAction Stop}

                Write-Verbose -Message 'File deleted.'
            } catch {
                Write-Error -Message ('Credential file can not be updated. '+($Error[0].Exception))
            }

            # if file is not existing, export values
            #if ((!(Test-Path -Path $FileNameUser)) -and (!(Test-Path -Path $FileNamePass))) {
            if (!(Test-Path -Path $FileName)) {
                # create encrypted values
                $EncUser = ConvertTo-SecureString -String ($Cred.UserName) -AsPlainText -Force | ConvertFrom-SecureString 
                $EncPass = ConvertFrom-SecureString -SecureString ($Cred.Password)
                # save them
                try {
                    #Write-Verbose -Message "Attempting to save credentials to $Path..."
                    Write-Verbose -Message "Attempting to save credentials to $FileName..."
                    #$EncUser | Out-File $FileNameUser
                    #$EncPass | Out-File $FileNamePass
                    $EncUser | Out-File $FileName
                    $EncPass | Out-File $FileName -Append
                    $VerboseMessage = "Environment $E processed successfully."
                    Write-Verbose -Message 'Credentials saved.'
                } catch {
                    Write-Error -Message ('Credential files can not be updated. '+($Error[0].Exception))
                }
            }
        }
        
        # return values
        if ($PassThru) {
            Write-Verbose -Message "Returning credentials for $E to output."
            Return $Cred
        }

        Write-Verbose -Message $VerboseMessage
    } # end of foreach

} # end of function

#-----------------------------------------------------[Function closing]-----------------------------------------------------

END {
    # function closing phase
    Write-Verbose -Message 'Write-Credential finishing.'

}

} # end of function code
#----------------------------------------------------[End of function]------------------------------------------------------

#---------------------------------------------------[Comments section]------------------------------------------------------
