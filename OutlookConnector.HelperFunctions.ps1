# Helper functions used within Outlook Connector module
# Functions are not exported out of module

function Trim-Length {
    param(
        [parameter(ValueFromPipeline=$True)][string] $Str,
        [parameter(Mandatory=$true,Position=1)][ValidateRange(1,[int]::MaxValue)][int] $Length
    )
    ($Str.TrimStart()[0..($Length-1)] -join "").TrimEnd()
}

function New-Folder {
    # creates new folder if not existing
    param([Parameter(Mandatory=$true)][String]$TargetFolder)

    if (!(Test-Path -Path $TargetFolder)) {
        try {
            New-Item -ItemType Directory -Path $TargetFolder -ErrorAction Stop | Out-Null
        } catch {
            throw "Target folder $TargetFolder can't be created."
        }
    }
}

function Get-Properties {
    # get list of properties from provided pattern
    param(
        [Parameter(Mandatory=$true)][String]$FileNameFormat
    )
    $RegEx = '(?:\%)(.+?)(?:(?:\|)(.*?))?(?:\%)'
    [regex]::Matches($FileNameFormat,$RegEx) | ForEach-Object { $_.Groups[1].Value }
}

function Validate-Properties {
    # verifies if sent object has all needed properties
    # it returns $null if everything is fine, or list of missing properties
    # it should be used as if (Validate-Properties) {there are errors} else {no errors}
    param(
        [Parameter(Mandatory=$true)][psobject]$InputObject,
        [Parameter(Mandatory=$true)][String[]]$RequiredProperties
    )
    $ObjectProperties = ($InputObject | Get-Member).Name
    $NotFoundProperties = @()

    foreach ($Prop in $RequiredProperties) {
        if ($Prop -notin $ObjectProperties) {
            $NotFoundProperties += $Prop
        }
    }

    if ($NotFoundProperties.Length -gt 0) {
        try {
            $ClassName = [enum]::GetName([Microsoft.Office.Interop.Outlook.OlObjectClass], $Message.Class) -replace '^ol'
        } catch {
            $ClassName = ""
        }
        if ($Message.Subject) { # TODO Simplify this section
            $ErrorMessage = 'Message "' + $($Message.Parent.FolderPath) + '\' + $Message.Subject + '" of type ' + $ClassName + ' is not proper object.'
        } elseif ($ClassName) {
            $ErrorMessage = 'Message of type ' + $ClassName + ' is not proper object.'
        } else {
            $ErrorMessage = 'Message is not proper object.'
        }
        $ErrorMessage += ' Missing: ' + ($NotFoundProperties -join ', ') + '.'
        throw $ErrorMessage
    }
}

function Expand-Properties {
    # generates string based on provided pattern and object
    # replaces each property in pattern specified with %PropertyName|format% with value of Property from sent object
    # calling function should verify that all properties exist
    param(
        [Parameter(Mandatory=$true)][psobject]$InputObject,
        [Parameter(Mandatory=$true)][String]$Pattern
    )
    $RegEx = '(?:\%)(.+?)(?:(?:\|)(.*?))?(?:\%)'
    $EnumProperties = @('Class','Sensitivity','RemoteStatus','Importance','FlagStatus','FlagIcon','BodyFormat')

    $ExpandedString = $Pattern
    while ($ExpandedString -match $RegEx) {
        $match = $Matches[0]
        $property = $Matches[1]
        if ($Matches.Count -ge 3) {
            $format = $Matches[2]
        } else {
            $format = ""
        }
        $propertyValue = $InputObject.($property)
        if ($property -in $EnumProperties) {
            # the predefined set of recognized enum properties are expanded to the name of the enumeration entry instead of just the integer value
            if ($property -eq 'Class') { # non standard, property Class returns enum value OlObjectClass and not OlClass
                $propertyEnumName = "OlObjectClass"
            } else {
                $propertyEnumName = "Ol${property}"
            }
            $propertyEnum = "Microsoft.Office.Interop.Outlook.${propertyEnumName}" -as [type]
            $propertyValue = [enum]::GetName($propertyEnum, $propertyValue) -replace '^ol'
        }
        if ($format -match '^[\d]+$') { # if format is just an integer value then treat it as max length
            $ExpandedString = $ExpandedString.Replace($match, ($propertyValue | Trim-Length $format))
        } else {
            $ExpandedString = $ExpandedString.Replace($match, "{0:$format}" -f $propertyValue)
        }
    }

    # return value
    $ExpandedString
}

function Get-ValidFileName {
    # reference
    # https://msdn.microsoft.com/en-us/library/windows/desktop/aa365247%28v=vs.85%29.aspx?f=255&MSPPError=-2147217396
    # https://gallery.technet.microsoft.com/scriptcenter/Save-Email-from-Outlook-to-3abf1ff3#content
    
    param([Parameter(Mandatory=$true)][String]$FileName)

    # remove illegal characters
    foreach ($char in ([System.IO.Path]::GetInvalidFileNameChars())) {
        $FileName = $FileName.Replace($char, '_')
    }

    # trim whitespace from both ends
    $FileName = $FileName.Trim()

    # return value
    $FileName
}

function Get-UniqueFilePath {
    # Generates a unique full file path based on supplied folder path, file name and extension.
    # If file with that name exists, it will add numbering like (1), (2), etc. at the end of name
    # Specified folder must already exist.
    param(
        [Parameter(Mandatory=$true)][String]$FolderPath,
        [Parameter(Mandatory=$true)][String]$FileName,
        [Parameter(Mandatory=$true)][String]$Extension
    )

    # Handling of path limitations:
    # We are limited by the regular Windows API limit defined as MAX_PATH with
    # constant value 260: The maximum number of characters, including a null terminator,
    # for the fully-pathed filename.
    # Remember that, although total file path is normally limited to 259 characters,
    # the folder path is limited to 247 characters (we do not check that limit here
    # as we assume the specified folder already exists), so we should always have some
    # available space for filename and extension.
    $MaxPath = 260-1
    if ((Join-Path $FolderPath ".${Extension}").Length -ge $MaxPath) {
        throw "Path is too long" # No room for any filename other than the extension!
    }
    $MaxBaseFilePath = $MaxPath - $Extension.Length - 1 
    $BaseFilePath = Join-Path -Path $FolderPath -ChildPath $FileName
    if ($BaseFilePath.Length -gt $MaxBaseFilePath) {
        $FullFilePath = ($BaseFilePath | Trim-Length $MaxBaseFilePath) + ".${Extension}"
        $Truncated = $true
    } else {
        $FullFilePath = "${BaseFilePath}.${Extension}"
        $Truncated = $false
    }

    # Check if file exists, and if yes, update name with numbering
    $i = 0
    while (Test-Path -LiteralPath $FullFilePath) {
        $Numbering = '~' + (++$i)
        if ((Join-Path $FolderPath "${Numbering}.${Extension}").Length -gt $MaxPath) {
            throw "Path is not unique and it is too long to add numbering"
        }
        if ($BaseFilePath.Length + $Numbering.Length -gt $MaxBaseFilePath) {
            $FullFilePath = ($BaseFilePath | Trim-Length ($MaxBaseFilePath - $Numbering.Length)) + "${Numbering}.${Extension}"
            $Truncated = $true
        } else {
            $FullFilePath = "${BaseFilePath}${Numbering}.${Extension}"
        }
    }
    if ($i -gt 0) {
        Write-Verbose "Filename suffix added to get unique file path: $FullFilePath"
    }
    if ($Truncated) {
        Write-Warning "Filename truncated to get valid path: $FullFilePath"
    }
    $FullFilePath
}