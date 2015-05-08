$modulePath = $PSCommandPath -replace '\.Tests\.ps1$', '.psm1'
$prefix = [guid]::NewGuid().Guid -replace '[^a-f\d]'

$module = $null

try
{
    $module = Import-Module $modulePath -PassThru -Prefix $prefix -ErrorAction Stop

    InModuleScope $module.Name {
        Describe 'Get-TargetResource' {
            Mock Test-Path { $true } -ParameterFilter { $LiteralPath -like "$env:SystemRoot\system32\GroupPolicy\*\registry.pol" }

            Context 'When the value is present' {
                $key       = 'Software\Testing'
                $valueName = 'TestValue'
                $data      = [uint32]12345
                $type      = [Microsoft.Win32.RegistryValueKind]::DWord

                Mock Get-PolicyFileEntry {
                    return New-Object psobject -Property @{
                        ValueName = $valueName
                        Key       = $key
                        Data      = $data
                        Type      = $type
                    }
                }

                It 'Returns the proper state' {
                    $hashtable = Get-TargetResource -PolicyType Machine -KeyValueName "$key\$valueName"

                    $hashtable.PSBase.Count | Should Be 5
                    $hashtable['PolicyType'] | Should Be Machine
                    $hashtable['Ensure'] | Should Be 'Present'
                    $hashtable['KeyValueName'] | Should Be "$key\$valueName"
                    $hashtable['Data'] | Should Be $data
                    $hashtable['Type'] | Should Be $type
                }
            }

            Context 'When the value is absent' {
                $key       = 'Software\Testing'
                $valueName = 'TestValue'

                Mock Get-PolicyFileEntry { }

                It 'Returns the proper state' {
                    $hashtable = Get-TargetResource -PolicyType Machine -KeyValueName "$key\$valueName"

                    $hashtable.PSBase.Count | Should Be 5
                    $hashtable['PolicyType'] | Should Be Machine
                    $hashtable['Ensure'] | Should Be 'Absent'
                    $hashtable['KeyValueName'] | Should Be "$key\$valueName"
                    $hashtable['Data'] | Should Be $null
                    $hashtable['Type'] | Should Be ([Microsoft.Win32.RegistryValueKind]::Unknown)
                }
            }
        }

        Describe 'Test-TargetResource' {
            # Using a hashtable here to avoid variable naming collisions with the function parameter names
            # (Sometimes a danger when using InModuleScope)

            $mockValues = @{
                key       = 'Software\Testing'
                valueName = 'TestValue'
                data      = [uint32]12345
                type      = [Microsoft.Win32.RegistryValueKind]::DWord
            }

            Mock Test-Path { $true } -ParameterFilter { $LiteralPath -like "$env:SystemRoot\system32\GroupPolicy\*\registry.pol" }

            Mock Get-PolicyFileEntry {
                return New-Object psobject -Property @{
                    ValueName = $mockValues.valueName
                    Key       = $mockValues.key
                    Data      = $mockValues.data
                    Type      = $mockValues.type
                }
            }

            It 'Returns true when the system is in the proper state' {
                $result = Test-TargetResource -PolicyType Machine `
                                              -KeyValueName "$($mockValues.key)\$($mockValues.valueName)" `
                                              -Ensure 'Present' `
                                              -Data $mockValues.data `
                                              -Type $mockValues.type

                $result | Should Be $true
            }

            It 'Returns false if the Data does not match' {
                $result = Test-TargetResource -PolicyType Machine `
                                              -KeyValueName "$($mockValues.key)\$($mockValues.valueName)" `
                                              -Ensure 'Present' `
                                              -Data BogusData `
                                              -Type $mockValues.type

                $result | Should Be $false
            }

            It 'Returns false if the Type does not match' {
                $result = Test-TargetResource -PolicyType Machine `
                                              -KeyValueName "$($mockValues.key)\$($mockValues.valueName)" `
                                              -Ensure 'Present' `
                                              -Data $mockValues.data `
                                              -Type ([Microsoft.Win32.RegistryValueKind]::QWord)

                $result | Should Be $false
            }

            Mock Get-PolicyFileEntry { }

            It 'Returns false when the entry is not found' {
                $result = Test-TargetResource -PolicyType Machine `
                                              -KeyValueName "$($mockValues.key)\$($mockValues.valueName)" `
                                              -Ensure 'Present' `
                                              -Data $mockValues.data `
                                              -Type $mockValues.type

                $result | Should Be $false
            }
        }

        Describe 'Set-TargetResource' {
            # Using a hashtable here to avoid variable naming collisions with the function parameter names
            # (Sometimes a danger when using InModuleScope)

            $values = @{
                key       = 'Software\Testing'
                valueName = 'TestValue'
                data      = [uint32]12345
                type      = [Microsoft.Win32.RegistryValueKind]::DWord
            }

            Mock Test-Path { $true } -ParameterFilter { $LiteralPath -like "$env:SystemRoot\system32\GroupPolicy\*\registry.pol" }
            Mock Set-PolicyFileEntry
            Mock Remove-PolicyFileEntry

            It 'Calls Set-PolicyFileEntry when Ensure is set to Present' {
                Set-TargetResource -PolicyType Machine `
                                   -KeyValueName "$($values.key)\$($values.valueName)" `
                                   -Ensure 'Present' `
                                   -Data $values.data `
                                   -Type $values.type

                Assert-MockCalled Set-PolicyFileEntry -Scope It -ParameterFilter {
                    $Key -eq $values.key -and
                    $ValueName -eq $values.valueName -and
                    $Data -eq $values.data -and
                    $Type -eq $values.type
                }

                Assert-MockCalled Remove-PolicyFileEntry -Scope It -Times 0
            }

            It 'Calls Remove-PolicyFileEntry when Ensure is set to Absent' {
                Set-TargetResource -PolicyType Machine `
                                   -KeyValueName "$($values.key)\$($values.valueName)" `
                                   -Ensure 'Absent'

                Assert-MockCalled Remove-PolicyFileEntry -Scope It -ParameterFilter {
                    $Key -eq $values.key -and
                    $ValueName -eq $values.valueName
                }

                Assert-MockCalled Set-PolicyFileEntry -Scope It -Times 0
            }
        }
    }
}
finally
{
    if ($module) { Remove-Module -ModuleInfo $module }
}
