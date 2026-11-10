@echo off
chcp 65001 >nul
cd /d "%~dp0"
where node >nul 2>nul
if %errorlevel% equ 0 (
    echo Iniciando servidor local do Dashboard em http://localhost:3000 ...
    start "" "http://localhost:3000"
    node server.local.js
) else (
    start "" "%~dp0index.html"
)
