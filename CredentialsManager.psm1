# Credentials Manager module file
# Current version 1.00, December 2015

# File List
$FileList = @(
    'Read-Credential.ps1',
    'Write-Credential.ps1',
    'Convert-Credential.ps1')

# Import all files from file list
foreach ($File in $FileList) {
    . (Join-Path -Path $PSScriptRoot -ChildPath $File) # -Verbose:$False
}

if (!($Script:CredentialsPath)) {
    $Script:CredentialsPath = Join-Path -Path $env:APPDATA -ChildPath 'CredentialsManager'
    }
