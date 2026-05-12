@echo off
setlocal enabledelayedexpansion
title MailWeave Build

echo.
echo  ==========================================================
echo    MailWeave ^| Windows Installer Build
echo  ==========================================================
echo.

cd /d "%~dp0"

:: ── Step 1 : Ensure Python dependencies are installed ────────────────────────
echo [1/5] Checking Python dependencies...
python -m pip install --quiet --upgrade ^
    tkinterdnd2 extract-msg python-docx reportlab pillow pyinstaller pywin32 ^
    pypdf docx2pdf
if errorlevel 1 (
    echo  ERROR: pip install failed.
    goto :fail
)
echo       Done.
echo.

:: ── Step 2 : Generate application icon ───────────────────────────────────────
echo [2/5] Generating mailweave.ico...
python create_icon.py
if errorlevel 1 (
    echo  ERROR: Icon generation failed.
    goto :fail
)

:: ── Step 2b : Sync version stamps (version_info.txt + _version.iss) ──────────
echo        Syncing version stamps from version.py...
python sync_version.py
if errorlevel 1 (
    echo  ERROR: Version sync failed.
    goto :fail
)

:: ── Step 2c : Regenerate UI icons ────────────────────────────────────────────
echo        Regenerating UI icons...
python generate_ui_icons.py
if errorlevel 1 (
    echo  WARN: Icon regeneration failed; using existing icon_*.png files.
)
echo.

:: ── Step 3 : Build with PyInstaller ──────────────────────────────────────────
echo [3/5] Running PyInstaller (this takes a minute)...

if exist "build\MailWeave"  rmdir /s /q "build\MailWeave"
if exist "dist\MailWeave"   rmdir /s /q "dist\MailWeave"

python -m PyInstaller MailWeave.spec --clean --noconfirm
if errorlevel 1 (
    echo.
    echo  ERROR: PyInstaller build failed.  See output above.
    goto :fail
)

if not exist "dist\MailWeave\MailWeave.exe" (
    echo  ERROR: dist\MailWeave\MailWeave.exe not found after build.
    goto :fail
)

echo.
echo       PyInstaller build succeeded.
echo       Output: dist\MailWeave\
echo.

:: ── Step 4 : Sign the application executable ─────────────────────────────────
echo [4/5] Signing dist\MailWeave\MailWeave.exe...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0sign.ps1" "dist\MailWeave\MailWeave.exe"
if errorlevel 1 (
    echo.
    echo  WARN: Application exe was not signed. Continuing with unsigned build.
    echo        Set MAILWEAVE_SIGN_THUMBPRINT to override the cert thumbprint.
    echo.
)

:: ── Step 5 : Create Windows installer with Inno Setup ────────────────────────
echo [5/5] Creating Windows installer...

set "ISCC="
for %%P in (
    "%LocalAppData%\Programs\Inno Setup 6\ISCC.exe"
    "C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
    "C:\Program Files\Inno Setup 6\ISCC.exe"
    "C:\Program Files (x86)\Inno Setup 5\ISCC.exe"
    "C:\Program Files\Inno Setup 5\ISCC.exe"
) do (
    if exist %%P (
        set "ISCC=%%P"
        goto :found_iscc
    )
)

echo.
echo  Inno Setup was not found on this machine.
echo.
echo  To create the installer (.exe), download Inno Setup (free) from:
echo    https://jrsoftware.org/isdl.php
echo.
echo  Or install it via winget:
echo    winget install --id JRSoftware.InnoSetup
echo.
echo  The un-packaged application is ready to run right now at:
echo    dist\MailWeave\MailWeave.exe
echo.
goto :done_no_installer

:found_iscc
echo       Found Inno Setup at: !ISCC!
echo.

if not exist "installer_output" mkdir "installer_output"

!ISCC! installer.iss
if errorlevel 1 (
    echo.
    echo  ERROR: Inno Setup compilation failed.
    goto :fail
)

:: Sign the installer too
if exist "installer_output\MailWeave-Setup-1.0.exe" (
    echo.
    echo       Signing installer_output\MailWeave-Setup-1.0.exe...
    powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0sign.ps1" "installer_output\MailWeave-Setup-1.0.exe"
    if errorlevel 1 (
        echo  WARN: Installer was not signed.
    )
)

echo.
echo  ============================================================
echo    BUILD COMPLETE
echo  ============================================================
echo.
echo    Installer : installer_output\MailWeave-Setup-1.0.exe
echo    Portable  : dist\MailWeave\MailWeave.exe
echo.
goto :end

:done_no_installer
echo  ============================================================
echo    BUILD COMPLETE  (no installer — Inno Setup not found)
echo  ============================================================
echo.
echo    Portable app : dist\MailWeave\MailWeave.exe
echo.
goto :end

:fail
echo.
echo  ============================================================
echo    BUILD FAILED
echo  ============================================================
echo.
pause
exit /b 1

:end
pause
