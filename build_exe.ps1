$ErrorActionPreference = "Stop"

Set-Location -LiteralPath $PSScriptRoot

$buildDir = ".\build_clean"
$distDir = ".\dist_clean"
$releaseDir = ".\release_bundle"
$releaseProgramDir = Join-Path $releaseDir "program"
$settingsPath = Join-Path $PSScriptRoot "settings.json"
$settingsBackupPath = Join-Path $env:TEMP ("yeehe_settings_backup_" + [guid]::NewGuid().ToString("N") + ".json")

try {
    if (Test-Path -LiteralPath $settingsPath) {
        Copy-Item -LiteralPath $settingsPath -Destination $settingsBackupPath -Force
    }

    # The published build uses fresh defaults and never contains local API keys.
    python -X utf8 -c "import json; from term_extractor_app.storage import build_default_settings; open('settings.json', 'w', encoding='utf-8').write(json.dumps(build_default_settings().to_dict(), ensure_ascii=False, indent=2))"

    Write-Host "Cleaning previous build artifacts..."
    Remove-Item -LiteralPath $buildDir -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $distDir -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $releaseDir -Recurse -Force -ErrorAction SilentlyContinue

    Write-Host "Building WebUI package with PyInstaller..."
    pyinstaller ".\yeehe_toolkit_suite.spec" --noconfirm --clean --workpath $buildDir --distpath $distDir

    Write-Host "Copying packaged program into release bundle..."
    New-Item -ItemType Directory -Path $releaseProgramDir -Force | Out-Null
    Get-ChildItem -Path (Join-Path $distDir "Yeehe_Toolkit_Suite") -Force | ForEach-Object {
        Copy-Item -LiteralPath $_.FullName -Destination $releaseProgramDir -Recurse -Force
    }

    Copy-Item -LiteralPath ".\start_webui.bat" -Destination $releaseDir -Force
    Copy-Item -LiteralPath ".\logo.png" -Destination $releaseDir -Force
    Write-Host "Build completed."
}
finally {
    if (Test-Path -LiteralPath $settingsBackupPath) {
        Move-Item -LiteralPath $settingsBackupPath -Destination $settingsPath -Force
    }
}
