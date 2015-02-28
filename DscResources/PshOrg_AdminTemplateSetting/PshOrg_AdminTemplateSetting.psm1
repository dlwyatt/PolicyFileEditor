Import-Module PolicyFileEditor -ErrorAction Stop

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

        return $null -ne $entry -and $Type -eq $entry.Type -and (CompareData $entry.Data $Data -Type $Type)
    }
    else # Ensure is 'Absent'
    {
        if (-not $fileExists) { return $true }
        $entry = Get-PolicyFileEntry -Path $path -Key $key -ValueName $valueName

        return $null -eq $entry
    }
}

function GetPolFilePath
{
    param (
        [string] $PolicyType
    )

    switch ($PolicyType)
    {
        'Machine'
        {
            return Join-Path $env:SystemRoot System32\GroupPolicy\Machine\registry.pol
        }
        
        'User'
        {
            return Join-Path $env:SystemRoot System32\GroupPolicy\User\registry.pol
        }

        'Administrators'
        {
            # BUILTIN\Administrators well-known SID
            return Join-Path $env:SystemRoot System32\GroupPolicyUsers\S-1-5-32-544\User\registry.pol
        }

        'NonAdministrators'
        {
            # BUILTIN\Users well-known SID
            return Join-Path $env:SystemRoot System32\GroupPolicyUsers\S-1-5-32-545\User\registry.pol
        }
    }
}

function CompareData
{
    param (
        [object] $First,
        [object] $Second,
        [Microsoft.Win32.RegistryValueKind] $Type
    )

    if ($Type -eq [Microsoft.Win32.RegistryValueKind]::String -or
        $Type -eq [Microsoft.Win32.RegistryValueKind]::ExpandString -or
        $Type -eq [Microsoft.Win32.RegistryValueKind]::DWord -or
        $Type -eq [Microsoft.Win32.RegistryValueKind]::QWord)
    {
        return @($First)[0] -ceq @($Second)[0]
    }

    # If we get here, $Type is either MultiString or Binary, both of which need to compare arrays.
    # The PolicyFileEditor module never returns type Unknown or None.

    if ($First.Count -ne $Second.Count) { return $false }

    $count = $first.Count
    for ($i = 0; $i -lt $count; $i++)
    {
        if ($First[$i] -cne $Second[$i]) { return $false }
    }

    return $true
}

function ParseKeyValueName
{
    param ([string] $KeyValueName)

    if ($KeyValueName.EndsWith('\'))
    {
        $key       = $KeyValueName -replace '\\$'
        $valueName = ''
    }
    else
    {
        $key       = Split-Path $KeyValueName -Parent
        $valueName = Split-Path $KeyValueName -Leaf
    }

    return $key, $valueName
}

Export-ModuleMember Get-TargetResource, Test-TargetResource, Set-TargetResource
