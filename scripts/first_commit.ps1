param(
    [string]$Message = "Initialize project baseline"
)

$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $PSScriptRoot
Set-Location -LiteralPath $ProjectRoot

if (-not (Test-Path -LiteralPath ".git")) {
    git init
}

git branch -M main
git add .
git commit -m $Message
git status --short --branch
