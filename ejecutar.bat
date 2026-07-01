@echo off
echo ==============================================
echo Detectando instalacion de Python...
echo ==============================================

REM Intentar detectar Python en el PATH del sistema
where python >nul 2>&1
if %errorlevel% == 0 (
    set PYTHON_EXE=python
    goto :found
)

REM Buscar Python en ubicaciones comunes de instalacion por usuario
for %%P in (
    "%LOCALAPPDATA%\Python\pythoncore-3.14-64\python.exe"
    "%LOCALAPPDATA%\Python\pythoncore-3.13-64\python.exe"
    "%LOCALAPPDATA%\Python\pythoncore-3.12-64\python.exe"
    "%LOCALAPPDATA%\Python\pythoncore-3.11-64\python.exe"
    "%LOCALAPPDATA%\Programs\Python\Python314\python.exe"
    "%LOCALAPPDATA%\Programs\Python\Python313\python.exe"
    "%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
    "%LOCALAPPDATA%\Programs\Python\Python311\python.exe"
    "C:\Python314\python.exe"
    "C:\Python313\python.exe"
    "C:\Python312\python.exe"
    "C:\Python311\python.exe"
    "C:\Users\ucaquiroga\AppData\Local\Python\pythoncore-3.14-64\python.exe"
) do (
    if exist %%P (
        set PYTHON_EXE=%%P
        goto :found
    )
)

echo.
echo ERROR: No se encontro Python instalado en este equipo.
echo.
echo Por favor instala Python desde: https://www.python.org/downloads/
echo Al instalar, marca la opcion "Add Python to PATH"
echo.
pause
exit /b 1

:found
echo Python encontrado: %PYTHON_EXE%
echo.
echo ==============================================
echo Instalando dependencias necesarias (pywin32)...
echo ==============================================
%PYTHON_EXE% -m pip install pywin32 --quiet

echo.
echo ==============================================
echo Ejecutando el Agente de Captura SAP...
echo ==============================================
%PYTHON_EXE% captura_sap.py

echo.
pause
