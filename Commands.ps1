#requires -Version 2.0

$scriptRoot = Split-Path $MyInvocation.MyCommand.Path
. "$scriptRoot\Common.ps1"

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
    $existingEntry = $policyFile.GetValue($key, $ValueName)

    if ($null -ne $existingEntry -and $Type -eq (PolEntryTypeToRegistryValueKind $existingEntry.Type))
    {
        $existingData = GetEntryData -Entry $existingEntry -Type $Type
        if (DataIsEqual $Data $existingData -Type $Type)
        {
            Write-Verbose 'Specified policy setting is already configured.  No changes were made.'
            return
        }
    }

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

    if ($null -eq $entry)
    {
        Write-Verbose 'Specified policy setting already does not exist.  No changes were made.'
        return
    }

    $policyFile.DeleteValue($Key, $ValueName)
    $doUpdateGptIni = -not $NoGptIniUpdate
    SavePolicyFile -PolicyFile $policyFile -UpdateGptIni:$doUpdateGptIni -ErrorAction Stop
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
