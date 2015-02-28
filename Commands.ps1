function Set-PolicyFileEntry
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        [string] $Path,

        [Parameter(Mandatory = $true, Position = 1)]
        [string] $Key,

        [Parameter(Mandatory = $true, Position = 1)]
        [string] $ValueName,

        [Parameter(Mandatory = $true, Position = 2)]
        [AllowEmptyString()]
        [AllowEmptyCollection()]
        [object] $Data,

        [ValidateScript({
            if ($_ -eq [Microsoft.Win32.RegistryValueKind]::Unknown)
            {
                throw 'Unknown is not a valid value for the Type parameter'
            }

            if ($_ -eq [Microsoft.Win32.RegistryValueKind]::None)
            {
                throw 'None is not a valid value for the Type parameter'
            }

            return $true
        })]
        [Microsoft.Win32.RegistryValueKind] $Type = [Microsoft.Win32.RegistryValueKind]::String,

        [switch] $NoGptIniUpdate
    )

    $policyFile = OpenPolicyFile -Path $Path -ErrorAction Stop

    try
    {
        switch ($Type)
        {
            ([Microsoft.Win32.RegistryValueKind]::Binary)
            {
                $bytes = $Data -as [byte[]]
                if ($null -eq $bytes)
                {
                    throw 'When -Type is set to Binary, -Data must be passed a Byte[] array.'
                }
                else
                {
                    $policyFile.SetBinaryValue($Key, $ValueName, $bytes)
                }

                break
            }

            ([Microsoft.Win32.RegistryValueKind]::String)
            {
                $string = $Data.ToString()
                $policyFile.SetStringValue($Key, $ValueName, $string)
                break
            }

            ([Microsoft.Win32.RegistryValueKind]::ExpandString)
            {
                $string = $Data.ToString()
                $policyFile.SetStringValue($Key, $ValueName, $string, $true)
                break
            }

            ([Microsoft.Win32.RegistryValueKind]::DWord)
            {
                $dword = $Data -as [UInt32]
                if ($null -eq $dword)
                {
                    throw 'When -Type is set to DWord, -Data must be passed a valid UInt32 value.'
                }
                else
                {
                    $policyFile.SetDWORDValue($key, $ValueName, $dword)
                }

                break
            }

            ([Microsoft.Win32.RegistryValueKind]::QWord)
            {
                $qword = $Data -as [UInt64]
                if ($null -eq $qword)
                {
                    throw 'When -Type is set to QWord, -Data must be passed a valid UInt64 value.'
                }
                else
                {
                    $policyFile.SetQWORDValue($key, $ValueName, $qword)
                }

                break
            }

            ([Microsoft.Win32.RegistryValueKind]::MultiString)
            {
                $strings = [string[]] @(
                    foreach ($item in $data)
                    {
                        $item.ToString()
                    }
                )

                $policyFile.SetMultiStringValue($Key, $ValueName, $strings)

                break
            }

        } # switch ($Type)

        $doUpdateGptIni = -not $NoGptIniUpdate
        SavePolicyFile -PolicyFile $policyFile -UpdateGptIni:$doUpdateGptIni -ErrorAction Stop
    }
    catch
    {
        throw
    }
}

function Get-PolicyFileEntry
{
    [CmdletBinding(DefaultParameterSetName = 'ByKeyAndValue')]
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        [string] $Path,

        [Parameter(Mandatory = $true, Position = 1, ParameterSetName = 'ByKeyAndValue')]
        [string] $Key,

        [Parameter(Mandatory = $true, Position = 2, ParameterSetName = 'ByKeyAndValue')]
        [string] $ValueName,

        [Parameter(Mandatory = $true, ParameterSetName = 'All')]
        [switch] $All
    )

    $policyFile = OpenPolicyFile -Path $Path -ErrorAction Stop

    if ($PSCmdlet.ParameterSetName -eq 'ByKeyAndValue')
    {
        $entry = $policyFile.GetValue($Key, $ValueName)

        if ($null -ne $entry)
        {
            PolEntryToPsObject -PolEntry $entry
        }
    }
    else
    {
        foreach ($entry in $policyFile.Entries)
        {
            PolEntryToPsObject -PolEntry $entry
        }
    }
}

function Remove-PolicyFileEntry
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        [string] $Path,

        [Parameter(Mandatory = $true, Position = 1)]
        [string] $Key,

        [Parameter(Mandatory = $true, Position = 2)]
        [string] $ValueName,

        [switch] $NoGptIniUpdate
    )

    $policyFile = OpenPolicyFile -Path $Path -ErrorAction Stop
    $entry = $policyFile.GetValue($Key, $ValueName)

    if ($null -ne $entry)
    {
        $policyFile.DeleteValue($Key, $ValueName)
        $doUpdateGptIni = -not $NoGptIniUpdate
        SavePolicyFile -PolicyFile $policyFile -UpdateGptIni:$doUpdateGptIni -ErrorAction Stop
    }
}

function Update-GptIniVersion
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [ValidateScript({
            if (Test-Path -LiteralPath $_ -PathType Leaf)
            {
                return $true
            }

            throw "Path '$_' does not exist."
        })]
        [string] $Path,

        [Parameter(Mandatory = $true)]
        [ValidateSet('Machine', 'User')]
        [string[]] $PolicyType
    )

    IncrementGptIniVersion @PSBoundParameters
}

function OpenPolicyFile
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string] $Path
    )

    $policyFile = New-Object TJX.PolFileEditor.PolFile
    $policyFile.FileName = $PSCmdlet.GetUnresolvedProviderPathFromPSPath($Path)

    if (Test-Path -LiteralPath $policyFile.FileName)
    {
        try
        {
            $policyFile.LoadFile()
        }
        catch
        {
            $errorRecord = $_
            $message = "Error loading policy file at path '$Path': $($errorRecord.Exception.Message)"
            $exception = New-Object System.Exception($message, $errorRecord.Exception)
            throw $exception
        }
    }

    return $policyFile
}

function PolEntryToPsObject
{
    param (
        [TJX.PolFileEditor.PolEntry] $PolEntry
    )

    $type = PolEntryTypeToRegistryValueKind $PolEntry.Type
    $data = GetEntryData -Entry $PolEntry -Type $type

    return New-Object psobject -Property @{
        Key       = $PolEntry.KeyName
        ValueName = $PolEntry.ValueName
        Type      = $type
        Data      = $data
    }
}

function GetEntryData
{
    param (
        [TJX.PolFileEditor.PolEntry] $Entry,
        [Microsoft.Win32.RegistryValueKind] $Type
    )

    switch ($type)
    {
        ([Microsoft.Win32.RegistryValueKind]::Binary)
        {
            return $Entry.BinaryValue
        }

        ([Microsoft.Win32.RegistryValueKind]::DWord)
        {
            return $Entry.DWORDValue
        }

        ([Microsoft.Win32.RegistryValueKind]::ExpandString)
        {
            return $Entry.StringValue
        }

        ([Microsoft.Win32.RegistryValueKind]::MultiString)
        {
            return $Entry.MultiStringValue
        }

        ([Microsoft.Win32.RegistryValueKind]::QWord)
        {
            return $Entry.QWORDValue
        }

        ([Microsoft.Win32.RegistryValueKind]::String)
        {
            return $Entry.StringValue
        }
    }

}

function SavePolicyFile
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [TJX.PolFileEditor.PolFile] $PolicyFile,

        [switch] $UpdateGptIni
    )

    $parentPath = Split-Path $PolicyFile.FileName -Parent
    if (-not (Test-Path -LiteralPath $parentPath -PathType Container))
    {
        try
        {
            $null = New-Item -Path $parentPath -ItemType Directory -ErrorAction Stop
        }
        catch
        {
            $errorRecord = $_
            $message = "Error creating parent folder of path '$Path': $($errorRecord.Exception.Message)"
            $exception = New-Object System.Exception($message, $errorRecord.Exception)
            throw $exception
        }
    }

    try
    {
        $PolicyFile.SaveFile()
    }
    catch
    {
        $errorRecord = $_
        $message = "Error saving policy file to path '$($PolicyFile.FileName)': $($errorRecord.Exception.Message)"
        $exception = New-Object System.Exception($message, $errorRecord.Exception)
        throw $exception
    }

    if ($UpdateGptIni)
    {
        if ($policyFile.FileName -match '^(.*)\\+([^\\]+)\\+[^\\]+$' -and
            $Matches[2] -eq 'User' -or $Matches[2] -eq 'Machine')
        {
            $iniPath = Join-Path $Matches[1] GPT.ini

            if (Test-Path -LiteralPath $iniPath -PathType Leaf)
            {
                IncrementGptIniVersion -Path $iniPath -PolicyType $Matches[2]
            }
            else
            {
                Write-Warning "File '$iniPath' does not exist, and the -NoGptIniUpdate switch was not specified."
            }
        }
    }
}

function IncrementGptIniVersion
{
    param (
        [string] $Path,
        [string[]] $PolicyType
    )

    $foundVersionLine = $false
    $section = ''

    $newContents = @(
        foreach ($line in Get-Content $Path)
        {
            # This might not be the most unreadable regex ever, but it's trying hard to be!
            # It's looking for section lines:  [SectionName]
            if ($line -match '^\s*\[([^\]]+)\]\s*$')
            {
                $section = $matches[1]
            }
            elseif ($section -eq 'General' -and
                    $line -match '^\s*Version\s*=\s*(\d+)\s*$' -and
                    $null -ne ($version = $matches[1] -as [uint32]))
            {
                $foundVersionLine = $true

                # User version is the high 16 bits, Machine version is the low 16 bits.
                # Reference:  http://blogs.technet.com/b/grouppolicy/archive/2007/12/14/understanding-the-gpo-version-number.aspx

                $pair = UInt32ToUInt16Pair -UInt32 $version

                if ($PolicyType -contains 'User')
                {
                    $pair.HighPart++
                }

                if ($PolicyType -contains 'Machine')
                {
                    $pair.LowPart++
                }

                $newVersion = UInt16PairToUInt32 -UInt16Pair $pair

                $line = "Version=$newVersion"
            }

            $line
        }
    )

    if (-not $foundVersionLine)
    {
        throw "GPT ini file '$Path' did not contain a valid 'Version=<number>' line in the General section."
    }

    Set-Content -Path $Path -Value $newContents -Encoding Ascii
}

function UInt32ToUInt16Pair
{
    param ([UInt32] $UInt32)

    # Deliberately avoiding bitwise shift operators here, for PowerShell v2 compatibility.

    $lowPart  = $UInt32 -band 0xFFFF
    $highPart = ($UInt32 - $lowPart) / 0x10000

    return New-Object psobject -Property @{
        LowPart  = [UInt16] $lowPart
        HighPart = [UInt16] $highPart
    }
}

function UInt16PairToUInt32
{
    param ([object] $UInt16Pair)

    # Deliberately avoiding bitwise shift operators here, for PowerShell v2 compatibility.

    return ([UInt32] $UInt16Pair.HighPart) * 0x10000 + $UInt16Pair.LowPart
}

function PolEntryTypeToRegistryValueKind
{
    param ([TJX.PolFileEditor.PolEntryType] $PolEntryType)

    switch ($PolEntryType)
    {
        ([TJX.PolFileEditor.PolEntryType]::REG_NONE)
        {
            return [Microsoft.Win32.RegistryValueKind]::None
        }

        ([TJX.PolFileEditor.PolEntryType]::REG_DWORD)
        {
            return [Microsoft.Win32.RegistryValueKind]::DWord
        }

        ([TJX.PolFileEditor.PolEntryType]::REG_DWORD_BIG_ENDIAN)
        {
            return [Microsoft.Win32.RegistryValueKind]::DWord
        }

        ([TJX.PolFileEditor.PolEntryType]::REG_BINARY)
        {
            return [Microsoft.Win32.RegistryValueKind]::Binary
        }

        ([TJX.PolFileEditor.PolEntryType]::REG_EXPAND_SZ)
        {
            return [Microsoft.Win32.RegistryValueKind]::ExpandString
        }

        ([TJX.PolFileEditor.PolEntryType]::REG_MULTI_SZ)
        {
            return [Microsoft.Win32.RegistryValueKind]::MultiString
        }

        ([TJX.PolFileEditor.PolEntryType]::REG_QWORD)
        {
            return [Microsoft.Win32.RegistryValueKind]::QWord
        }

        ([TJX.PolFileEditor.PolEntryType]::REG_SZ)
        {
            return [Microsoft.Win32.RegistryValueKind]::String
        }
    }
}
