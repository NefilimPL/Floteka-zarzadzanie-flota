param(
    [string]$OutputDirectory = '',
    [string]$Tag = $env:FLOTEKA_BUILD_TAG
)

$ErrorActionPreference = 'Stop'
$flotekaRoot = Split-Path -Parent $PSScriptRoot
Set-Location -LiteralPath $flotekaRoot
$flotekaTranscriptStarted = $false
$flotekaOriginalPath = $env:PATH
$flotekaOriginalExcelFile = $env:FLOTA_EXCEL_FILE
$flotekaOriginalVersionDir = $env:FLOTEKA_EXCEL_VERSION_DIR
$flotekaOriginalPythonPath = $env:PYTHONPATH
$flotekaOriginalPythonHome = $env:PYTHONHOME

function Invoke-FlotekaPython {
    param([string[]]$Arguments)
    & $script:flotekaPython @Arguments
    if ($LASTEXITCODE -ne 0) { throw "Python step failed (exit $LASTEXITCODE): $($Arguments -join ' ')" }
}

try {
    if (-not $OutputDirectory) {
        $OutputDirectory = Join-Path $flotekaRoot ('dist/Floteka-lokalnie/' + (Get-Date -Format 'yyyyMMdd-HHmmss-fff'))
    }
    $flotekaOutput = if ([IO.Path]::IsPathRooted($OutputDirectory)) { $OutputDirectory } else { Join-Path $flotekaRoot $OutputDirectory }
    $flotekaOutput = [IO.Path]::GetFullPath($flotekaOutput)
    if (Test-Path -LiteralPath $flotekaOutput) {
        if (-not (Test-Path -LiteralPath $flotekaOutput -PathType Container) -or @(Get-ChildItem -LiteralPath $flotekaOutput -Force).Count -gt 0) {
            throw "Output directory must be empty: $flotekaOutput"
        }
    } else {
        New-Item -ItemType Directory -Path $flotekaOutput | Out-Null
    }
    Start-Transcript -LiteralPath (Join-Path $flotekaOutput 'build.log') -NoClobber | Out-Null
    $flotekaTranscriptStarted = $true
    $env:PYTHONPATH = $null
    $env:PYTHONHOME = $null

    # Probe both launchers: the Windows py launcher may be installed without
    # a registered interpreter even when a working python.exe is on PATH.
    $flotekaBasePython = $null
    $flotekaProbe = 'import sys,tkinter; assert sys.version_info[:2] == (3,13) and sys.maxsize > 2**32; r=tkinter.Tk(); r.withdraw(); r.update(); r.destroy(); print(sys.executable)'
    foreach ($flotekaLauncher in @('py', 'python')) {
        if (Get-Command $flotekaLauncher -ErrorAction SilentlyContinue) {
            $flotekaLauncherArgs = if ($flotekaLauncher -eq 'py') { @('-3.13', '-c', $flotekaProbe) } else { @('-c', $flotekaProbe) }
            try {
                $flotekaProbeOutput = & $flotekaLauncher @flotekaLauncherArgs 2>$null
                if ($LASTEXITCODE -eq 0 -and $flotekaProbeOutput) { $flotekaBasePython = @($flotekaProbeOutput)[-1]; break }
            } catch { continue }
        }
    }
    if (-not $flotekaBasePython) { throw 'Install Python 3.13 x64 with Tcl/Tk and pip, add it to PATH, and run from a Windows desktop session.' }
    $flotekaEnvironment = Join-Path $flotekaRoot 'build/local-venv'
    $flotekaPython = Join-Path $flotekaEnvironment 'Scripts/python.exe'
    if (-not (Test-Path -LiteralPath $flotekaPython)) {
        & $flotekaBasePython -m venv $flotekaEnvironment
        if ($LASTEXITCODE -ne 0) { throw 'Could not create the isolated build environment.' }
    }
    Invoke-FlotekaPython -Arguments @('-c', $flotekaProbe)
    $env:PATH = (Join-Path $flotekaEnvironment 'Scripts') + [IO.Path]::PathSeparator + $flotekaOriginalPath
    Invoke-FlotekaPython -Arguments @('-m', 'pip', 'install', '-r', 'requirements-build.txt')
    $flotekaTagArgs = @()
    if ($Tag) { $flotekaTagArgs = @('--tag', $Tag) }
    Invoke-FlotekaPython -Arguments (@('-m', 'tools.generate_excel', '--output', $flotekaOutput) + $flotekaTagArgs)
    $env:FLOTA_EXCEL_FILE = Join-Path $flotekaOutput 'Floteka.xlsm'
    $env:FLOTEKA_EXCEL_VERSION_DIR = $flotekaOutput
    $flotekaTestTemp = Join-Path $flotekaRoot ('build/pytest-local-' + [guid]::NewGuid().ToString('N'))
    Invoke-FlotekaPython -Arguments @('-m', 'pytest', 'tests', '--capture=sys', '-q', '-p', 'no:cacheprovider', '--basetemp', $flotekaTestTemp, '--junitxml', (Join-Path $flotekaOutput 'test-results.xml'))
    # build_exe.ps1 invokes python via PATH; the isolated environment comes first.
    & (Join-Path $PSScriptRoot 'build_exe.ps1') -OutputDirectory $flotekaOutput -Tag $Tag
    Invoke-FlotekaPython -Arguments @('-m', 'tools.make_demo', '--output', $flotekaOutput)
    Invoke-FlotekaPython -Arguments @('-m', 'tools.package_local', '--output', $flotekaOutput)
    Write-Output "Gotowy pakiet: $(Join-Path $flotekaOutput 'Floteka-pakiet.zip')"
    Write-Output 'XLSM: zachowano makra szablonow; lokalna budowa nie uruchamiala makr w Excelu.'
} catch {
    Write-Error -Message $_ -ErrorAction Continue
    exit 1
} finally {
    $env:PATH = $flotekaOriginalPath
    $env:FLOTA_EXCEL_FILE = $flotekaOriginalExcelFile
    $env:FLOTEKA_EXCEL_VERSION_DIR = $flotekaOriginalVersionDir
    $env:PYTHONPATH = $flotekaOriginalPythonPath
    $env:PYTHONHOME = $flotekaOriginalPythonHome
    if ($flotekaTranscriptStarted) { Stop-Transcript | Out-Null }
}
