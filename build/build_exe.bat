@echo off
setlocal EnableExtensions
cd /d "%~dp0"

title Rapportage PDF Generator - Build

echo ==========================================
echo   Rapportage PDF Generator - EXE BUILD
echo ==========================================
echo.

REM ==================================================
REM Mappenstructuur
REM ==================================================
REM
REM Rapportage PDF Generator\
REM |
REM +-- app.py
REM +-- README.txt
REM +-- requirements.txt
REM |
REM +-- config\
REM |   +-- config.ini
REM |
REM +-- assets\
REM |   +-- Rapportage_PDF_Generator.ico
REM |
REM +-- build\
REM |   +-- build_exe.bat
REM |   +-- build_version.py
REM |
REM +-- dist\
REM |   +-- Rapportage Merger TOOL.exe
REM |
REM +-- temp\
REM     +-- build\
REM     +-- version_info.txt
REM     +-- Rapportage Merger TOOL.spec
REM
REM ==================================================

REM De map waarin deze BAT staat
set "BUILD_DIR=%~dp0"

REM Hoofdmap = één niveau boven build\
set "ROOT_DIR=%~dp0.."

REM Overige mappen
set "ASSETS_DIR=%ROOT_DIR%\assets"
set "TEMP_DIR=%ROOT_DIR%\temp"
set "DIST_DIR=%ROOT_DIR%\dist"

echo Hoofdmap:
echo %ROOT_DIR%
echo.

echo [1/5] Controle bestanden...
echo.

REM --------------------------------------------------
REM Controle app.py
REM --------------------------------------------------

if not exist "%ROOT_DIR%\app.py" (
    echo FOUT: app.py ontbreekt.
    echo.
    echo Verwacht op:
    echo %ROOT_DIR%\app.py
    echo.
    pause
    exit /b 1
)

REM --------------------------------------------------
REM Controle build_version.py
REM --------------------------------------------------

if not exist "%BUILD_DIR%build_version.py" (
    echo FOUT: build_version.py ontbreekt.
    echo.
    echo Verwacht op:
    echo %BUILD_DIR%build_version.py
    echo.
    pause
    exit /b 1
)

REM --------------------------------------------------
REM Controle ico
REM --------------------------------------------------

if not exist "%ASSETS_DIR%\Rapportage_PDF_Generator.ico" (
    echo FOUT: Rapportage_PDF_Generator.ico ontbreekt.
    echo.
    echo Verwacht op:
    echo %ASSETS_DIR%\Rapportage_PDF_Generator.ico
    echo.
    pause
    exit /b 1
)

echo Alle benodigde bestanden gevonden.
echo.

REM ==================================================
REM 2. Versie ophalen uit app.py
REM ==================================================

echo [2/5] Versie ophalen uit app.py...
echo.

echo APP:
echo %ROOT_DIR%\app.py
echo.

echo Build script:
echo %BUILD_DIR%build_version.py
echo.

python "%BUILD_DIR%build_version.py" "%ROOT_DIR%\app.py"

if errorlevel 1 (
    echo.
    echo ==========================================
    echo FOUT BIJ build_version.py
    echo ==========================================
    echo.
    pause
    exit /b 1
)

REM --------------------------------------------------
REM build_version.py maakt version_info.txt
REM in de huidige werkmap.
REM Deze BAT draait vanuit build\, dus:
REM build\version_info.txt
REM --------------------------------------------------

if not exist "%BUILD_DIR%version_info.txt" (
    echo.
    echo FOUT: version_info.txt is niet aangemaakt.
    echo.
    echo Verwacht op:
    echo %BUILD_DIR%version_info.txt
    echo.
    pause
    exit /b 1
)

REM Maak temp aan indien nodig
if not exist "%TEMP_DIR%" mkdir "%TEMP_DIR%"

REM Oude version_info verwijderen
if exist "%TEMP_DIR%\version_info.txt" (
    del /q "%TEMP_DIR%\version_info.txt"
)

REM Verplaats nieuwe version_info naar temp
move /y "%BUILD_DIR%version_info.txt" "%TEMP_DIR%\version_info.txt" >nul

if not exist "%TEMP_DIR%\version_info.txt" (
    echo.
    echo FOUT: version_info.txt kon niet naar temp worden verplaatst.
    pause
    exit /b 1
)

echo.
echo Versie-informatie succesvol aangemaakt.
echo.

REM ==================================================
REM 3. Oude buildbestanden verwijderen
REM ==================================================

echo [3/5] Oude buildbestanden verwijderen...
echo.

if exist "%TEMP_DIR%\build" (
    rmdir /s /q "%TEMP_DIR%\build"
)

if exist "%DIST_DIR%" (
    rmdir /s /q "%DIST_DIR%"
)

REM Maak dist opnieuw aan
if not exist "%DIST_DIR%" (
    mkdir "%DIST_DIR%"
)

REM Maak temp opnieuw aan
if not exist "%TEMP_DIR%" (
    mkdir "%TEMP_DIR%"
)

echo Oude buildbestanden verwijderd.
echo.

REM ==================================================
REM 4. EXE bouwen
REM ==================================================

echo [4/5] EXE bouwen...
echo.
echo Dit kan enige tijd duren.
echo.

python -m PyInstaller ^
    --onefile ^
    --windowed ^
    --name "Rapportage Merger TOOL" ^
    --distpath "%DIST_DIR%" ^
    --workpath "%TEMP_DIR%\build" ^
    --specpath "%TEMP_DIR%" ^
    --version-file "%TEMP_DIR%\version_info.txt" ^
    --collect-all pywin32 ^
    --hidden-import win32timezone ^
    --icon "%ASSETS_DIR%\Rapportage_PDF_Generator.ico" ^
    "%ROOT_DIR%\app.py"

if errorlevel 1 (
    echo.
    echo ==========================================
    echo BUILD MISLUKT
    echo ==========================================
    echo.
    pause
    exit /b 1
)

REM ==================================================
REM 5. Resultaat controleren
REM ==================================================

echo.
echo [5/5] Build controleren...
echo.

if not exist "%DIST_DIR%\Rapportage Merger TOOL.exe" (
    echo.
    echo FOUT: EXE niet gevonden.
    echo.
    echo Verwacht op:
    echo %DIST_DIR%\Rapportage Merger TOOL.exe
    echo.
    pause
    exit /b 1
)

echo.
echo ==========================================
echo       BUILD GESLAAGD
echo ==========================================
echo.
echo EXE:
echo %DIST_DIR%\Rapportage Merger TOOL.exe
echo.
echo Versie-informatie:
echo %TEMP_DIR%\version_info.txt
echo.
echo Icoon:
echo %ASSETS_DIR%\Rapportage_PDF_Generator.ico
echo.
echo ==========================================
echo.

pause