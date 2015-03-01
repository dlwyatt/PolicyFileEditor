Add-Type -Path $PSScriptRoot\..\..\PolFileEditor.dll -ErrorAction Stop
. "$PSScriptRoot\..\..\Commands.ps1"

function Get-TargetResource
{
    [OutputType([hashtable])]
    param (
        [Parameter(Mandatory)]
        [ValidateSet('Machine', 'User', 'Administrators', 'NonAdministrators')]
        [string] $PolicyType,

        [Parameter(Mandatory)]
        [string] $KeyValueName
    )

    $configuration = @{
        PolicyType   = $PolicyType
        KeyValueName = $KeyValueName
        Ensure       = 'Absent'
        Data         = $null
        Type         = [Microsoft.Win32.RegistryValueKind]::Unknown
    }

    $path = GetPolFilePath -PolicyType $PolicyType
    if (Test-Path -LiteralPath $path -PathType Leaf)
    {
        $key, $valueName = ParseKeyValueName $KeyValueName
        $entry = Get-PolicyFileEntry -Path $path -Key $key -ValueName $valueName

        if ($entry)
        {
            $configuration['Ensure'] = 'Present'
            $configuration['Type']   = $entry.Type
            $configuration['Data']   = $entry.Data
        }
    }

    return $configuration
}

function Set-TargetResource
{
    param (
        [Parameter(Mandatory)]
        [ValidateSet('Machine', 'User', 'Administrators', 'NonAdministrators')]
        [string] $PolicyType,

        [Parameter(Mandatory)]
        [string] $KeyValueName,

        [ValidateSet('Present', 'Absent')]
        [string] $Ensure = 'Present',

        [string[]] $Data,

        [Microsoft.Win32.RegistryValueKind] $Type = [Microsoft.Win32.RegistryValueKind]::String
    )

    $path = GetPolFilePath -PolicyType $PolicyType
    $key, $valueName = ParseKeyValueName $KeyValueName

    if ($Type -eq [Microsoft.Win32.RegistryValueKind]::MultiString -or
        $Type -eq [Microsoft.Win32.RegistryValueKind]::Binary)
    {
        $dataToSet = $Data
    }
    else
    {
        $dataToSet = $Data[0]
    }

    if ($Ensure -eq 'Present')
    {
        Set-PolicyFileEntry -Path $path -Key $key -ValueName $valueName -Data $dataToSet -Type $Type
    }
    else
    {
        Remove-PolicyFileEntry -Path $path -Key $key -ValueName $valueName
    }
}

function Test-TargetResource
{
    [OutputType([bool])]
    param (
        [Parameter(Mandatory)]
        [ValidateSet('Machine', 'User', 'Administrators', 'NonAdministrators')]
        [string] $PolicyType,

        [Parameter(Mandatory)]
        [string] $KeyValueName,

        [ValidateSet('Present', 'Absent')]
        [string] $Ensure = 'Present',

        [string[]] $Data,

        [Microsoft.Win32.RegistryValueKind] $Type = [Microsoft.Win32.RegistryValueKind]::String
    )

    $path = GetPolFilePath -PolicyType $PolicyType
    $key, $valueName = ParseKeyValueName $KeyValueName

    $fileExists = Test-Path -LiteralPath $path -PathType Leaf

    if ($Ensure -eq 'Present')
    {
        if (-not $fileExists) { return $false }
        $entry = Get-PolicyFileEntry -Path $path -Key $key -ValueName $valueName

        return $null -ne $entry -and $Type -eq $entry.Type -and (DataIsEqual $entry.Data $Data -Type $Type)
    }
    else # Ensure is 'Absent'
    {
        if (-not $fileExists) { return $true }
        $entry = Get-PolicyFileEntry -Path $path -Key $key -ValueName $valueName

        return $null -eq $entry
    }
}

Export-ModuleMember Get-TargetResource, Test-TargetResource, Set-TargetResource
