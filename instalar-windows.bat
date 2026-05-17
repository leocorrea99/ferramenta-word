@echo off
title Ferramenta Word — Instalador
cls
echo.
echo  +--------------------------------------------+
echo  ^|         Ferramenta Word                    ^|
echo  ^|         Instalador para Windows            ^|
echo  +--------------------------------------------+
echo.
echo  Aguarde, instalando...
echo.

powershell -ExecutionPolicy Bypass -NoProfile -Command "$t=$env:TEMP+'\fw_inst.ps1'; (New-Object Net.WebClient).DownloadFile('https://leocorrea99.github.io/ferramenta-word/instalar-windows.ps1',$t); & $t; Remove-Item $t -EA 0"

echo.
pause
