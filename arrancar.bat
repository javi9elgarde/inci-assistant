@echo off
title Asistente de Incidencias - EasyWEB
echo.
echo  ====================================
echo   Asistente de Incidencias - EasyWEB
echo  ====================================
echo.
echo  Iniciando servidor...
echo  Abre tu navegador en: http://localhost:5000
echo.
start "" "http://localhost:5000"
"C:\Users\JAVIERGARDERUIZ\AppData\Local\Programs\Python\Python312\python.exe" "%~dp0app.py"
pause
