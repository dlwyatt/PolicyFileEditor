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

    try
    {
        $path = GetPolFilePath -PolicyType $PolicyType
        return GetTargetResourceCommon -Path $path -KeyValueName $KeyValueName
    }
    catch
    {
        Write-Error -ErrorRecord $_
        return
    }
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

    try
    {
        $path = GetPolFilePath -PolicyType $PolicyType
        SetTargetResourceCommon -Path $path -KeyValueName $KeyValueName -Ensure $Ensure -Data $Data -Type $Type
    }
    catch
    {
        Write-Error -ErrorRecord $_
        return
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

    if ($null -eq $Data) { $Data = @() }

    try
    {
        Assert-ValidDataAndType -Data $Data -Type $Type
    }
    catch
    {
        Write-Error -ErrorRecord $_
        return
    }

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

function Assert-ValidDataAndType
{
    param (
        [string[]] $Data,
        [Microsoft.Win32.RegistryValueKind] $Type
    )

    if ($Type -ne [Microsoft.Win32.RegistryValueKind]::MultiString -and
        $Type -ne [Microsoft.Win32.RegistryValueKind]::Binary -and
        $Data.Count -gt 1)
    {
        throw 'Do not pass arrays with multiple values to the -Data parameter when -Type is not set to either Binary or MultiString.'
    }

}

Export-ModuleMember Get-TargetResource, Test-TargetResource, Set-TargetResource
