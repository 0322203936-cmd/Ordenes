@echo off
title Sistema Excel - Ordenes Wholesale (Pacifica Farms)
cd /d "%~dp0"
echo Iniciando servidor en http://localhost:3050 ...
start "" "http://localhost:3050"
node server.js
pause
