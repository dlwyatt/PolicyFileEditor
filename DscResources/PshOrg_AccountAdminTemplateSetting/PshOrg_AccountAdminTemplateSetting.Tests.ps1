<#
Most of the tests we need are already taken care of in PshOrg_AdminTemplateSetting.
The only difference between the two resources is that AccountAdminTemplateSetting calls
GetPolFilePath with the -Account parameter instead of the -PolicyType parameter.

Rather than going through the motions of testing the *-TargetResource methods here,
we'll just unit test the GetPolFilePath function instead.
#>

. "$PSScriptRoot\..\..\Common.ps1"

Describe 'GetPolFilePath with -Account parameter' {
    It 'Returns the proper path for a well-known SID' {
        $actual = GetPolFilePath -Account BUILTIN\Administrators
        $expected = Join-Path $env:SystemRoot System32\GroupPolicyUsers\S-1-5-32-544\User\registry.pol
        $actual | Should Be $expected
    }

    It 'Returns the proper path for a local account' {
        $computer = [adsi]"WinNT://$env:COMPUTERNAME"
        $user = $computer.Children |
                Where SchemaClassName -eq User |
                Select -First 1

        $sid = New-Object System.Security.Principal.SecurityIdentifier($user.objectSid[0], 0)

        $actual = GetPolFilePath -Account $user.Name[0]
        $expected = Join-Path $env:SystemRoot System32\GroupPolicyUsers\$($sid.Value)\User\registry.pol

        $actual | Should Be $expected
    }

    It 'Allows a SID string to be passed directly' {
        $actual = GetPolFilePath -Account S-1-5-32-544
        $expected = Join-Path $env:SystemRoot System32\GroupPolicyUsers\S-1-5-32-544\User\registry.pol
        $actual | Should Be $expected
    }
}
