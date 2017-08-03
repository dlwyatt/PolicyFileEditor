#requires -Version 3.0

Task default -depends Build,Sign

Properties {
    $source               = $psake.build_script_dir
    $buildTarget          = "$home\Documents\WindowsPowerShell\Modules\PolicyFileEditor"
    $signerCertThumbprint = '572DD922FB11FD90D47EE466EDF4191019FB19FA'
    $signerTimestampUrl   = 'http://timestamp.digicert.com'

    $filesToExclude = @(
        'README.md'
        '*.Tests.ps1'
        'build.psake.ps1'
    )
}

Task Test {
    $result = Invoke-Pester -Path $source -PassThru
    $failed = $result.FailedCount

    if ($failed -gt 0)
    {
        throw "$failed unit tests failed; build aborting."
    }
}

Task Build -depends Test {
    function Get-RelativePath
    {
        param ( [string] $Path, [string] $RelativeTo )
        return $Path -replace "^$([regex]::Escape($RelativeTo))\\?"
    }

    if (Test-Path -Path $buildTarget -PathType Container)
    {
        Remove-Item -Path $buildTarget -Recurse -Force -ErrorAction Stop
    }

    $null = New-Item -Path $buildTarget -ItemType Directory -ErrorAction Stop

    Get-ChildItem -Path $source\* -File -Recurse -ErrorAction Stop |
    Where { $file = $_; -not ($filesToExclude | Where { $file.Name -like $_ }) } |
    ForEach {
        $sourceFile = $_
        $relativePath = Get-RelativePath -Path $sourceFile.FullName -RelativeTo $source
        $targetPath = Join-Path $buildTarget $relativePath
        $parent = Split-Path $targetPath -Parent

        if (-not (Test-Path -LiteralPath $parent -PathType Container))
        {
            $null = New-Item -Path $parent -ItemType Directory -ErrorAction Stop
        }

        Copy-Item -LiteralPath $sourceFile.FullName -Destination $targetPath -ErrorAction Stop
    }
}

Task Sign {
    if (-not $signerCertThumbprint)
    {
        throw 'Sign task cannot run without a value in the signerCertThumbprint property.'
    }

    $paths = @(
        "Cert:\CurrentUser\My\$signerCertThumbprint"
        "Cert:\LocalMachine\My\$signerCertThumbprint"
    )

    $cert = Get-ChildItem -Path $paths |
            Where-Object { $_.PrivateKey -is [System.Security.Cryptography.RSACryptoServiceProvider] } |
            Select-Object -First 1

    if ($cert -eq $null) {
        throw "Code signing certificate with thumbprint '$signerCertThumbprint' was not found, or did not have a usable private key."
    }

    $properties = @(
        @{ Label = 'Name'; Expression = { Split-Path -Path $_.Path -Leaf } }
        'Status'
        @{ Label = 'SignerCertificate'; Expression = { $_.SignerCertificate.Thumbprint } }
        @{ Label = 'TimeStamperCertificate'; Expression = { $_.TimeStamperCertificate.Thumbprint } }
    )

    $splat = @{
        Certificate   = $cert
        IncludeChain  = 'All'
        Force         = $true
        HashAlgorithm = 'SHA256'
    }

    if ($signerTimestampUrl) { $splat['TimestampServer'] = $signerTimestampUrl }

    Get-ChildItem -Path $buildTarget -Recurse -File -Include *.ps1, *.psm1, *.psd1 |
    Set-AuthenticodeSignature @splat -ErrorAction Stop |
    Format-Table -Property $properties -AutoSize
}
