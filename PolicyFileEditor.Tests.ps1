Remove-Module [P]olicyFileEditor
$scriptRoot = Split-Path $MyInvocation.MyCommand.Path -Parent
$psd1Path = Join-Path $scriptRoot PolicyFileEditor.psd1

$module = $null

function CreateDefaultGpo($Path)
{
    $paths = @(
        $Path
        Join-Path $Path Machine
        Join-Path $Path User
    )

    foreach ($p in $paths)
    {
        if (-not (Test-Path $p -PathType Container))
        {
            New-Item -Path $p -ItemType Directory -ErrorAction Stop
        }
    }

    $content = @'
[General]
gPCMachineExtensionNames=[{35378EAC-683F-11D2-A89A-00C04FBBCFA2}{D02B1F72-3407-48AE-BA88-E8213C6761F1}]
Version=65537
gPCUserExtensionNames=[{35378EAC-683F-11D2-A89A-00C04FBBCFA2}{D02B1F73-3407-48AE-BA88-E8213C6761F1}]
'@

    $gptIniPath = Join-Path $Path gpt.ini
    Set-Content -Path $gptIniPath -ErrorAction Stop -Encoding Ascii -Value $content

    Get-ChildItem -Path $Path -Include registry.pol -Force | Remove-Item -Force
}

function GetGptIniVersion($Path)
{
    foreach ($result in Select-String -Path $Path -Pattern '^\s*Version\s*=\s*(\d+)\s*$')
    {
        foreach ($match in $result.Matches)
        {
            $match.Groups[1].Value
        }
    }
}

try
{
    $module = Import-Module $psd1Path -ErrorAction Stop -PassThru -Force
    $gpoPath = 'TestDrive:\TestGpo'
    $gptIniPath = "$gpoPath\gpt.ini"

    Describe 'KeyValueName parsing' {
        InModuleScope PolicyFileEditor {
            $testCases = @(
                @{
                    KeyValueName  = 'Left\Right'
                    ExpectedKey   = 'Left'
                    ExpectedValue = 'Right'
                    Description   = 'Simple'
                }

                @{
                    KeyValueName  = 'Left\\Right'
                    ExpectedKey   = 'Left'
                    ExpectedValue = 'Right'
                    Description   = 'Multiple consecutive separators'
                }

                @{
                    KeyValueName  = '\Left\Right'
                    ExpectedKey   = 'Left'
                    ExpectedValue = 'Right'
                    Description   = 'Leading separator'
                }

                @{
                    KeyValueName  = 'Left\Right\'
                    ExpectedKey   = 'Left\Right'
                    ExpectedValue = ''
                    Description   = 'Trailing separator'
                }

                @{
                    KeyValueName  = '\\\Left\\\\Right\\\\\'
                    ExpectedKey   = 'Left\Right'
                    ExpectedValue = ''
                    Description   = 'Ridiculous with trailing separator'
                }

                @{
                    KeyValueName  = '\\\\\\\\Left\\\\\\\Right'
                    ExpectedKey   = 'Left'
                    ExpectedValue = 'Right'
                    Description   = 'Ridiculous with no trailing separator'
                }
            )

            It -TestCases $testCases 'Properly parses KeyValueName with <Description>' {
                param ($KeyValueName, $ExpectedKey, $ExpectedValue)
                $key, $valueName = ParseKeyValueName $KeyValueName

                $key | Should Be $ExpectedKey
                $valueName | Should Be $ExpectedValue
            }
        }
    }

    Describe 'Happy Path' {
        BeforeEach {
            CreateDefaultGpo -Path $gpoPath
        }

        Context 'Incrementing GPT.Ini version' {
            # User version is the high 16 bits, Machine version is the low 16 bits.
            # Reference:  http://blogs.technet.com/b/grouppolicy/archive/2007/12/14/understanding-the-gpo-version-number.aspx

            # Default value set in our CreateDefaultGpo function is 65537, or (1 -shl 16) + 1 ; Machine and User version both set to 1.
            # Decimal values ard hard-coded here so we can run the tests on PowerShell v2, which didn't have the -shl / -shr operators.
            # This puts the module's internal code which replaces these operators through a test as well.

            $testCases = @(
                @{
                    PolicyType = 'Machine'
                    Expected   = '65538' # (1 -shl 16) + 2
                }

                @{
                    PolicyType = 'User'
                    Expected   = '131073' # (2 -shl 16) + 1
                }

                @{
                    PolicyType = 'Machine', 'User'
                    Expected   = '131074' # (2 -shl 16) + 2
                }
            )

            It 'Sets the correct value for <PolicyType> updates' -TestCases $testCases {
                param ($PolicyType, $Expected)

                Update-GptIniVersion -Path $gptIniPath -PolicyType $PolicyType
                $version = @(GetGptIniVersion -Path $gptIniPath)

                $version.Count | Should Be 1
                $actual = $version[0]

                $actual | Should Be $Expected
            }
        }

        Context 'Automated modification of gpt.ini' {
            # These tests incidentally also cover the happy path functionality of
            # Set-PolicyFileEntry and Remove-PolicyFileEntry.  We'll cover errors
            # in a different section.

            $testCases = @(
                @{
                    PolicyType       = 'Machine'
                    ExpectedVersions = '65538', '65539' # (1 -shl 16) + 2, (1 -shl 16) + 2
                    NoGptIniUpdate   = $false
                    Count            = 1
                }

                @{
                    PolicyType       = 'User'
                    ExpectedVersions = '131073', '196609' # (2 -shl 16) + 1, (3 -shl 16) + 1
                    NoGptIniUpdate   = $false
                    Count            = 1
                }

                @{
                    PolicyType       = 'Machine'
                    ExpectedVersions = '65537', '65537' # (1 -shl 16) + 1, (1 -shl 16) + 1
                    NoGptIniUpdate   = $true
                    Count            = 1
                }

                @{
                    PolicyType       = 'User'
                    ExpectedVersions = '65537', '65537' # (1 -shl 16) + 1, (1 -shl 16) + 1
                    NoGptIniUpdate   = $true
                    Count            = 1
                }

                @{
                    PolicyType       = 'User'
                    ExpectedVersions = '131073', '196609' # (2 -shl 16) + 1, (3 -shl 16) + 1
                    NoGptIniUpdate   = $false
                    Count            = 2
                    EntriesToModify  = @(
                        New-Object psobject -Property @{
                            Key       = 'Software\Testing'
                            ValueName = 'Value1'
                            Type      = 'String'
                            Data      = 'Data'
                        }

                        New-Object psobject -Property @{
                            Key       = 'Software\Testing'
                            ValueName = 'Value2'
                            Type      = 'MultiString'
                            Data      = 'Multi', 'String', 'Data'
                        }
                    )
                }
            )

            It 'Behaves properly modifying <Count> entries in a <PolicyType> registry.pol file and NoGptIniUpdate is <NoGptIniUpdate>' -TestCases $testCases {
                param ($PolicyType, [string[]] $ExpectedVersions, [switch] $NoGptIniUpdate, [object[]] $EntriesToModify)

                if (-not $PSBoundParameters.ContainsKey('EntriesToModify'))
                {
                    $EntriesToModify = @(
                        New-Object psobject -Property @{
                            Key       = 'Software\Testing'
                            ValueName = 'TestValue'
                            Data      = 1
                            Type      = 'DWord'
                        }
                    )
                }

                $polPath = Join-Path $gpoPath $PolicyType\registry.pol

                $scriptBlock = {
                    $EntriesToModify | Set-PolicyFileEntry -Path $polPath -NoGptIniUpdate:$NoGptIniUpdate
                }

                # We do this next block of code twice to ensure that when "setting" a value that is already present in the
                # GPO, the version of gpt.ini is not updated.

                # Code is deliberately duplicated (rather then refactored into a loop) so that if we get failures,
                # the line numbers will tell us whether it was on the first or second execution of the duplicated
                # parts.

                $scriptBlock | Should Not Throw

                $expected = $ExpectedVersions[0]
                $version = @(GetGptIniVersion -Path $gptIniPath)

                $version.Count | Should Be 1
                $actual = $version[0]

                $actual | Should Be $expected

                $entries = @(Get-PolicyFileEntry -Path $PolPath -All)

                $entries.Count | Should Be $EntriesToModify.Count

                $count = $entries.Count
                for ($i = 0; $i -lt $count; $i++)
                {
                    $matchingEntry = $EntriesToModify | Where-Object { $_.Key -eq $entries[$i].Key -and $_.ValueName -eq $entries[$i].ValueName }

                    $entries[$i].ValueName | Should Be $matchingEntry.ValueName
                    $entries[$i].Key | Should Be $matchingEntry.Key
                    $entries[$i].Data | Should Be $matchingEntry.Data
                    $entries[$i].Type | Should Be $matchingEntry.Type
                }

                $scriptBlock | Should Not Throw

                $expected = $ExpectedVersions[0]
                $version = @(GetGptIniVersion -Path $gptIniPath)

                $version.Count | Should Be 1
                $actual = $version[0]

                $actual | Should Be $expected

                $entries = @(Get-PolicyFileEntry -Path $polPath -All)

                $entries.Count | Should Be $EntriesToModify.Count

                $count = $entries.Count
                for ($i = 0; $i -lt $count; $i++)
                {
                    $matchingEntry = $EntriesToModify | Where-Object { $_.Key -eq $entries[$i].Key -and $_.ValueName -eq $entries[$i].ValueName }

                    $entries[$i].ValueName | Should Be $matchingEntry.ValueName
                    $entries[$i].Key | Should Be $matchingEntry.Key
                    $entries[$i].Data | Should Be $matchingEntry.Data
                    $entries[$i].Type | Should Be $matchingEntry.Type
                }

                # End of duplicated bits; now we make sure that removing the entry
                # works, and still updates the gpt.ini version (if appropriate.)

                $scriptBlock = {
                    $EntriesToModify | Remove-PolicyFileEntry -Path $polPath -NoGptIniUpdate:$NoGptIniUpdate
                }

                $scriptBlock | Should Not Throw

                $expected = $ExpectedVersions[1]
                $version = @(GetGptIniVersion -Path $gptIniPath)

                $version.Count | Should Be 1
                $actual = $version[0]

                $actual | Should Be $expected

                $entries = @(Get-PolicyFileEntry -Path $polPath -All)

                $entries.Count | Should Be 0

                # Duplicate the Remove block for the same reasons; make sure the ini file isn't incremented
                # when the value is already missing.

                $scriptBlock | Should Not Throw

                $expected = $ExpectedVersions[1]
                $version = @(GetGptIniVersion -Path $gptIniPath)

                $version.Count | Should Be 1
                $actual = $version[0]

                $actual | Should Be $expected

                $entries = @(Get-PolicyFileEntry -Path $polPath -All)

                $entries.Count | Should Be 0
            }
        }

        Context 'Get/Set parity' {
            $testCases = @(
                @{
                    TestName = 'Creates a DWord value properly'
                    Type     = [Microsoft.Win32.RegistryValueKind]::DWord
                    Data     = @([UInt32]1)
                }

                @{
                    TestName = 'Creates a QWord value properly'
                    Type     = [Microsoft.Win32.RegistryValueKind]::QWord
                    Data     = @([UInt64]0x100000000L)
                }

                @{
                    TestName = 'Creates a String value properly'
                    Type     = [Microsoft.Win32.RegistryValueKind]::String
                    Data     = @('I am a string')
                }

                @{
                    TestName = 'Creates an ExpandString value properly'
                    Type     = [Microsoft.Win32.RegistryValueKind]::ExpandString
                    Data     = @('My temp path is %TEMP%')
                }

                @{
                    TestName = 'Creates a MultiString value properly'
                    Type     = [Microsoft.Win32.RegistryValueKind]::MultiString
                    Data     = [string[]]('I', 'am', 'a', 'multi', 'string')
                }

                @{
                    TestName = 'Creates a Binary value properly'
                    Type     = [Microsoft.Win32.RegistryValueKind]::Binary
                    Data     = [byte[]](1..32)
                }

                @{
                    TestName     = 'Allows hex strings to be assigned to DWord values'
                    Type         = [Microsoft.Win32.RegistryValueKind]::DWord
                    Data         = @('0x12345')
                    ExpectedData = [UInt32]0x12345
                }

                @{
                    TestName     = 'Allows hex strings to be assigned to QWord values'
                    Type         = [Microsoft.Win32.RegistryValueKind]::QWord
                    Data         = @('0x12345789')
                    ExpectedData = [Uint64]0x123456789L
                }

                @{
                    TestName     = 'Allows hex strings to be assigned to Binary types'
                    Type         = [Microsoft.Win32.RegistryValueKind]::Binary
                    Data         = '0x1', '0xFF', '0x12'
                    ExpectedData = [byte[]](0x1,0xFF,0x12)
                }

                @{
                    TestName     = 'Allows non-string data to be assigned to String values'
                    Type         = [Microsoft.Win32.RegistryValueKind]::String
                    Data         = @(12345)
                    ExpectedData = '12345'
                }

                @{
                    TestName     = 'Allows non-string data to be assigned to ExpandString values'
                    Type         = [Microsoft.Win32.RegistryValueKind]::ExpandString
                    Data         = @(12345)
                    ExpectedData = '12345'
                }

                @{
                    TestName     = 'Allows non-string data to be assigned to MultiString values'
                    Type         = [Microsoft.Win32.RegistryValueKind]::MultiString
                    Data         = 1..5
                    ExpectedData = '1', '2', '3', '4', '5'
                }
            )

            It '<TestName>' -TestCases $testCases {
                param ($TestName, $Type, $Data, $ExpectedData)

                $polPath = Join-Path $gpoPath Machine\registry.pol

                if (-not $PSBoundParameters.ContainsKey('ExpectedData'))
                {
                    $ExpectedData = $Data
                }

                $scriptBlock = {
                    Set-PolicyFileEntry -Path      $polPath `
                                        -Key       Software\Testing `
                                        -ValueName TestValue `
                                        -Data      $Data `
                                        -Type      $Type
                }

                $scriptBlock | Should Not Throw

                $entries = @(Get-PolicyFileEntry -Path $polPath -All)

                $entries.Count | Should Be 1

                $entries[0].ValueName | Should Be TestValue
                $entries[0].Key | Should Be Software\Testing
                $entries[0].Type | Should Be $Type

                $newData = @($entries[0].Data)
                $Data = @($Data)

                $Data.Count | Should Be $newData.Count

                $count = $Data.Count
                for ($i = 0; $i -lt $count; $i++)
                {
                    $Data[$i] | Should BeExactly $newData[$i]
                }

            }

            It 'Gets values by Key and PropertyName successfully' {
                $polPath   = Join-Path $gpoPath Machine\registry.pol
                $key       = 'Software\Testing'
                $valueName = 'TestValue'
                $data      = 'I am a string'
                $type      = ([Microsoft.Win32.RegistryValueKind]::String)

                $scriptBlock = {
                    Set-PolicyFileEntry -Path      $polPath `
                                        -Key       $key `
                                        -ValueName $valueName `
                                        -Data      $data `
                                        -Type      $type
                }

                $scriptBlock | Should Not Throw

                $entry = Get-PolicyFileEntry -Path $polPath -Key $key -ValueName $valueName

                $entry | Should Not Be $null
                $entry.ValueName | Should Be $valueName
                $entry.Key | Should Be $key
                $entry.Type | Should Be $type
                $entry.Data | Should Be $data
            }
        }

        Context 'Automatic creation of gpt.ini' {
            It 'Creates a gpt.ini file if one is not found' {
                Remove-Item $gptIniPath

                $path = Join-Path $gpoPath Machine\registry.pol

                Set-PolicyFileEntry -Path $path -Key 'Whatever' -ValueName 'Whatever' -Data 'Whatever' -Type String

                $gptIniPath | Should Exist
                GetGptIniVersion -Path $gptIniPath | Should Be 1
            }
        }
    }

    Describe 'Not-so-happy Path' {
        BeforeEach {
            CreateDefaultGpo -Path $gpoPath
        }

        $testCases = @(
            @{
                Type = [Microsoft.Win32.RegistryValueKind]::DWord
                ExpectedMessage = 'When -Type is set to DWord, -Data must be passed a valid UInt32 value.'
            }

            @{
                Type = [Microsoft.Win32.RegistryValueKind]::QWord
                ExpectedMessage = 'When -Type is set to QWord, -Data must be passed a valid UInt64 value.'
            }

            @{
                Type = [Microsoft.Win32.RegistryValueKind]::Binary
                ExpectedMessage = 'When -Type is set to Binary, -Data must be passed a Byte[] array.'
            }
        )

        It 'Gives a reasonable error when non-numeric data is passed to <Type> values' -TestCases $testCases {
            param ($Type, $ExpectedMessage)

            $scriptBlock = {
                Set-PolicyFileEntry -Path        $gpoPath\Machine\registry.pol `
                                    -Key         Software\Testing `
                                    -ValueName   TestValue `
                                    -Type        $Type `
                                    -Data        'I am not a number' `
                                    -ErrorAction Stop
            }

            $scriptBlock | Should Throw $ExpectedMessage
        }
    }
}
finally
{
    if ($null -ne $module)
    {
        Remove-Module -ModuleInfo $module -Force
    }
}
