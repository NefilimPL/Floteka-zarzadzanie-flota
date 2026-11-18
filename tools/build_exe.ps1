param(
    [string]$OutputDirectory = 'dist/Floteka',
    [string]$Tag = $env:FLOTEKA_BUILD_TAG
)

$ErrorActionPreference = 'Stop'
Set-Location -LiteralPath (Split-Path -Parent $PSScriptRoot)
$flotekaOutput = if ([IO.Path]::IsPathRooted($OutputDirectory)) { $OutputDirectory } else { Join-Path (Get-Location).Path $OutputDirectory }
$flotekaOutput = [IO.Path]::GetFullPath($flotekaOutput)
New-Item -ItemType Directory -Path $flotekaOutput -Force | Out-Null

$flotekaVersionArgs = @()
if ($Tag) { $flotekaVersionArgs = @('--tag', $Tag) }
python -m tools.release_version @flotekaVersionArgs
if ($LASTEXITCODE -ne 0) { throw 'Version preparation failed.' }
$flotekaVersion = Get-Content -LiteralPath 'build/version.json' -Raw -Encoding UTF8 | ConvertFrom-Json
python -m PyInstaller --noconfirm --clean --windowed --onefile --name Floteka --version-file build/windows-version.txt --distpath $flotekaOutput --workpath build/pyinstaller main.py
if ($LASTEXITCODE -ne 0) { throw 'PyInstaller build failed.' }

$flotekaExe = Join-Path $flotekaOutput 'Floteka.exe'
# Nowy, pusty katalog testowy zapobiega odczytaniu raportu poprzedniego EXE.
$flotekaSmoke = Join-Path (Get-Location) ('build/smoke-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $flotekaSmoke | Out-Null
$flotekaReport = Join-Path $flotekaSmoke 'exe-test.json'
# Sciezka jest cytowana osobno: katalog repo moze zawierac spacje.
$flotekaArguments = '--self-test "{0}"' -f $flotekaReport
$flotekaProcess = Start-Process -FilePath $flotekaExe -ArgumentList $flotekaArguments -WindowStyle Hidden -PassThru
if (-not $flotekaProcess.WaitForExit(60000)) {
    Stop-Process -Id $flotekaProcess.Id -Force -ErrorAction SilentlyContinue
    throw 'EXE self-test timed out.'
}
$flotekaProcess.Refresh()
if ($flotekaProcess.ExitCode -ne 0) { throw "EXE self-test failed: $($flotekaProcess.ExitCode)" }
if (-not (Test-Path -LiteralPath $flotekaReport)) { throw 'EXE did not produce a fresh test report.' }
$flotekaResult = Get-Content -LiteralPath $flotekaReport -Raw -Encoding UTF8 | ConvertFrom-Json
if (-not $flotekaResult.ok -or -not $flotekaResult.frozen -or $flotekaResult.app -ne 'Floteka' -or $flotekaResult.version -ne $flotekaVersion.version) {
    throw 'EXE report does not confirm a working, packaged Floteka application.'
}
Copy-Item -LiteralPath $flotekaReport -Destination (Join-Path $flotekaOutput 'exe-test.json') -Force
Copy-Item -LiteralPath 'README.md' -Destination (Join-Path $flotekaOutput 'INSTRUKCJA.md') -Force
$flotekaHash = (Get-FileHash -LiteralPath $flotekaExe -Algorithm SHA256).Hash.ToLowerInvariant()
Set-Content -LiteralPath (Join-Path $flotekaOutput 'SUMY-SHA256.txt') -Value ($flotekaHash + '  Floteka.exe') -Encoding UTF8
$flotekaBuild = [ordered]@{
    app = 'Floteka'
    version = $flotekaResult.version
    tag = $flotekaVersion.tag
    source_ref = $flotekaVersion.source_ref
    commit = $flotekaVersion.commit
    event_commit = $env:GITHUB_SHA
    run_id = $env:GITHUB_RUN_ID
    sha256 = $flotekaHash
    self_test = $true
}
$flotekaBuild | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $flotekaOutput 'build-info.json') -Encoding UTF8
Write-Output "Built and tested: $flotekaExe"
