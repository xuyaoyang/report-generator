param(
    [switch]$Clean
)

$ErrorActionPreference = "Stop"
$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $ProjectRoot

if ($Clean) {
    Remove-Item -Recurse -Force -LiteralPath ".\build", ".\dist" -ErrorAction SilentlyContinue
}

python -m PyInstaller .\report_generator.spec --noconfirm

Write-Host ""
$iscc = Get-Command iscc.exe -ErrorAction SilentlyContinue
if (-not $iscc) {
    $candidate = Get-ChildItem "$env:LOCALAPPDATA\Programs", "C:\Program Files*", "C:\Program Files (x86)" `
        -Recurse -Filter ISCC.exe -ErrorAction SilentlyContinue |
        Select-Object -First 1
    if ($candidate) {
        $iscc = $candidate
    }
}

if ($iscc) {
    $isccPath = if ($iscc.Source) { $iscc.Source } else { $iscc.FullName }
    & $isccPath ".\installer.iss"
    Write-Host "Installer created: $ProjectRoot\dist\融海报告生成安装包.exe"
} else {
    Write-Host "Inno Setup not found. Portable folder created: $ProjectRoot\dist\融海报告生成"
    Write-Host "Install Inno Setup and rerun this script to create a setup EXE."
}

Write-Host "First launch asks for a work directory and initializes Excel templates, material library, and report output there."
