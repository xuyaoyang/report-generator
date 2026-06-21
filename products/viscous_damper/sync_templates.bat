@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ========================================
echo   黏滞阻尼器 — 模板同步工具
echo ========================================
echo.
echo 将根据 template.docx 重新生成 template_prepared.docx
echo.

python prepare_template.py

echo.
echo 同步完成！按任意键退出...
pause >nul
