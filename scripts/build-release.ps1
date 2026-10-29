param(
    [switch]$Installer,
    [string]$InnoCompiler
)

$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $PSScriptRoot
$pythonPath = Join-Path $projectRoot 'venv\Scripts\python.exe'
if (-not (Test-Path -LiteralPath $pythonPath)) {
    throw 'Create venv and install requirements-dev.txt first. See README.md.'
}
$versionLine = Select-String -LiteralPath (Join-Path $projectRoot 'app.py') -Pattern '^APP_VERSION = "([0-9.]+)"$'
if (-not $versionLine) { throw 'APP_VERSION was not found in app.py.' }
$releaseVersion = $versionLine.Matches[0].Groups[1].Value
$releaseDirectory = Join-Path $projectRoot "build\release-v$releaseVersion"

if ($Installer) {
    if (-not $InnoCompiler) {
        $compilerCommand = Get-Command ISCC.exe -ErrorAction SilentlyContinue
        if ($compilerCommand) { $InnoCompiler = $compilerCommand.Source }
    }
    if (-not $InnoCompiler -or -not (Test-Path -LiteralPath $InnoCompiler)) {
        throw 'Inno Setup compiler is required. Supply -InnoCompiler with the full ISCC.exe path.'
    }
}

Push-Location $projectRoot
try {
    & $pythonPath -m PyInstaller --noconfirm --distpath (Join-Path $releaseDirectory 'app') --workpath (Join-Path $releaseDirectory 'work') app.spec
    if ($LASTEXITCODE -ne 0) { throw 'PyInstaller build failed.' }
    $appExe = Join-Path $releaseDirectory 'app\app.exe'
    Copy-Item -LiteralPath $appExe -Destination (Join-Path $releaseDirectory 'ppt-txt2img-portable.exe')
    if ($Installer) {
        & $InnoCompiler "/DMyAppSource=$appExe" "/O$releaseDirectory" '/Fppt-txt2img' inno-setup.iss
        if ($LASTEXITCODE -ne 0) { throw 'Installer compilation failed.' }
    }
    $artifacts = @(Join-Path $releaseDirectory 'ppt-txt2img-portable.exe')
    if ($Installer) { $artifacts += Join-Path $releaseDirectory 'ppt-txt2img.exe' }
    $checksums = foreach ($artifact in $artifacts) {
        $hash = Get-FileHash -LiteralPath $artifact -Algorithm SHA256
        '{0}  {1}' -f $hash.Hash.ToLowerInvariant(), (Split-Path -Leaf $artifact)
    }
    $checksums | Set-Content -LiteralPath (Join-Path $releaseDirectory 'SHA256SUMS.txt') -Encoding utf8
    Write-Output "Release files: $releaseDirectory"
} finally {
    Pop-Location
}
