@echo off
setlocal
set "CONVERTER_MSIX=%~dp0ChuyenDoiTaiLieu.msix"
powershell.exe -NoProfile -Command "$ErrorActionPreference = 'Stop'; Add-AppxPackage -Path $env:CONVERTER_MSIX -AllowUnsigned"
if errorlevel 1 (
    echo Installation failed. Run this file as administrator.
    pause
    exit /b 1
)
echo Installed. Right-click a supported file in File Explorer.
pause
