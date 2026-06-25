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
Write-Host "打包完成：$ProjectRoot\dist\融海报告生成"
Write-Host "首次启动会初始化：文档\报告生成工作目录"
