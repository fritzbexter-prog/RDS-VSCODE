@echo off
setlocal

echo ============================================================
echo   DragDropBox - NAV 2018 Control Add-in Build
echo ============================================================
echo.

:: --- Compiler suchen ---
set CSC=
for /d %%D in ("%SystemRoot%\Microsoft.NET\Framework\v4.0*") do (
    if exist "%%D\csc.exe" set CSC=%%D\csc.exe
)
if "%CSC%"=="" (
    echo FEHLER: csc.exe nicht gefunden.
    echo         .NET Framework 4.x muss installiert sein.
    goto :eof
)
echo Compiler: %CSC%

:: --- NAV DLL pruefen ---
set NAVDLL=C:\Program Files (x86)\Microsoft Dynamics NAV\110\RoleTailored Client\Microsoft.Dynamics.Framework.UI.Extensibility.dll
if not exist "%NAVDLL%" (
    echo FEHLER: NAV Extensibility DLL nicht gefunden:
    echo         %NAVDLL%
    echo.
    echo Bitte den Pfad in dieser Datei anpassen.
    goto :eof
)
echo NAV DLL:  %NAVDLL%

:: --- Strong Name Key erzeugen falls noetig ---
if not exist DragDropBox.snk (
    echo.
    echo Erzeuge Strong Name Key...
    set SNEXE=
    for /d %%D in ("%ProgramFiles(x86)%\Microsoft SDKs\Windows\v*") do (
        if exist "%%D\bin\sn.exe" set SNEXE=%%D\bin\sn.exe
        if exist "%%D\bin\NETFX 4.8 Tools\sn.exe" set SNEXE=%%D\bin\NETFX 4.8 Tools\sn.exe
        if exist "%%D\bin\NETFX 4.7.2 Tools\sn.exe" set SNEXE=%%D\bin\NETFX 4.7.2 Tools\sn.exe
        if exist "%%D\bin\NETFX 4.6.2 Tools\sn.exe" set SNEXE=%%D\bin\NETFX 4.6.2 Tools\sn.exe
    )
    if "%SNEXE%"=="" (
        echo FEHLER: sn.exe nicht gefunden. Windows SDK benoetigt.
        goto :eof
    )
    "%SNEXE%" -k DragDropBox.snk
    echo Key erzeugt: DragDropBox.snk
)

:: --- Kompilieren ---
echo.
echo Kompiliere DragDropBox.dll ...
"%CSC%" /target:library /out:DragDropBox.dll /keyfile:DragDropBox.snk /reference:"%NAVDLL%" /reference:System.Windows.Forms.dll /reference:System.Drawing.dll /reference:System.dll DragDropBoxControl.cs

if errorlevel 1 (
    echo.
    echo FEHLER beim Kompilieren!
    goto :eof
)

echo.
echo ============================================================
echo   Build erfolgreich: DragDropBox.dll
echo ============================================================

:: --- Public Key Token anzeigen ---
set SNEXE=
for /d %%D in ("%ProgramFiles(x86)%\Microsoft SDKs\Windows\v*") do (
    if exist "%%D\bin\sn.exe" set SNEXE=%%D\bin\sn.exe
    if exist "%%D\bin\NETFX 4.8 Tools\sn.exe" set SNEXE=%%D\bin\NETFX 4.8 Tools\sn.exe
    if exist "%%D\bin\NETFX 4.7.2 Tools\sn.exe" set SNEXE=%%D\bin\NETFX 4.7.2 Tools\sn.exe
    if exist "%%D\bin\NETFX 4.6.2 Tools\sn.exe" set SNEXE=%%D\bin\NETFX 4.6.2 Tools\sn.exe
)
if not "%SNEXE%"=="" (
    echo.
    echo Public Key Token:
    "%SNEXE%" -T DragDropBox.dll
)

echo.
echo Naechste Schritte:
echo   1. DragDropBox.dll kopieren nach:
echo      C:\Program Files (x86)\Microsoft Dynamics NAV\110\RoleTailored Client\Add-ins\
echo   2. In NAV unter Client Add-ins registrieren:
echo      - Control Add-in Name: DragDropBox
echo      - Public Key Token: siehe oben
echo      - Version: 1.0.0.0
echo   3. In einer NAV Form als Control verwenden
echo.

endlocal
