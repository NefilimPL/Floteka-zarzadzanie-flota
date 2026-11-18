param([string]$OutputDirectory = 'dist/Floteka')

$ErrorActionPreference = 'Stop'
Set-Location -LiteralPath (Split-Path -Parent $PSScriptRoot)
$flotaReleaseOutput = if ([IO.Path]::IsPathRooted($OutputDirectory)) { $OutputDirectory } else { Join-Path (Get-Location).Path $OutputDirectory }
$flotaReleaseOutput = [IO.Path]::GetFullPath($flotaReleaseOutput)
$flotaTestRoot = Join-Path $env:TEMP ('floteka-tests-' + [guid]::NewGuid().ToString('N'))
python -m pytest tests --ignore=tests/test_excel_layout.py -q -p no:cacheprovider --basetemp $flotaTestRoot
if ($LASTEXITCODE -ne 0) { throw 'Application tests failed.' }
python tools/build_excel.py --output $flotaReleaseOutput
if ($LASTEXITCODE -ne 0) { throw 'XLSM build failed.' }
$env:FLOTA_EXCEL_FILE = Join-Path $flotaReleaseOutput 'Floteka.xlsm'
python -m pytest tests/test_excel_layout.py -q -p no:cacheprovider
if ($LASTEXITCODE -ne 0) { throw 'XLSM layout tests failed.' }
python tools/test_excel.py
if ($LASTEXITCODE -ne 0) { throw 'XLSM integration tests failed.' }
python tools/test_excel_regressions.py
if ($LASTEXITCODE -ne 0) { throw 'XLSM regression tests failed.' }
python tools/test_deadlines_excel.py
if ($LASTEXITCODE -ne 0) { throw 'XLSM deadline color tests failed.' }
python tools/test_optional_excel.py
if ($LASTEXITCODE -ne 0) { throw 'XLSM optional data tests failed.' }
python tools/check_excel_filters_speed.py $env:FLOTA_EXCEL_FILE
if ($LASTEXITCODE -ne 0) { throw 'XLSM filter tests failed.' }
& "$PSScriptRoot/build_exe.ps1" -OutputDirectory $flotaReleaseOutput
Copy-Item -LiteralPath (Join-Path $flotaReleaseOutput 'exe-test.json') -Destination 'docs/exe-test.json' -Force
python tools/test_export_excel.py --exe (Join-Path $flotaReleaseOutput 'Floteka.exe')
if ($LASTEXITCODE -ne 0) { throw 'XLSX export compatibility tests failed.' }
python tools/make_demo.py --output $flotaReleaseOutput
if ($LASTEXITCODE -ne 0) { throw 'Demo creation failed.' }
python tools/package_release.py --output $flotaReleaseOutput
if ($LASTEXITCODE -ne 0) { throw 'Release packaging failed.' }
