@echo off
setlocal

echo ========================================================
echo        USTANOVKA NADSTROYKI SPECHELPER DLYA EXCEL
echo            (Prava administratora NE TREBUYUTSYA)
echo ========================================================
echo.

set "TARGET_DIR=%LOCALAPPDATA%\SpecHelper"
set "BIN_DIR=%~dp0bin"
set "DB_DIR=%~dp0databases"

echo [1/3] Kopirovanie faylov v: %TARGET_DIR%...
if not exist "%TARGET_DIR%" mkdir "%TARGET_DIR%"
if not exist "%TARGET_DIR%\databases" mkdir "%TARGET_DIR%\databases"

copy /y "%BIN_DIR%\*.dll" "%TARGET_DIR%\" >nul
copy /y "%DB_DIR%\*.xlsx" "%TARGET_DIR%\databases\" >nul

echo [2/3] Registratsiya v reestre HKCU...

set "CLSID_ADDIN={A19845E6-5F24-4B78-8316-D4E03C1247AA}"
set "CLSID_CTRL={D86F2C74-3208-4CE0-A4F7-1B8C512F9E82}"

set "SLASH_DIR=%TARGET_DIR:\=/%"
set "CODEBASE=file:///%SLASH_DIR%/SpecHelper.dll"

reg add "HKCU\Software\Classes\SpecHelper.AddIn" /ve /d "SpecHelper.AddIn" /f >nul
reg add "HKCU\Software\Classes\SpecHelper.AddIn\CLSID" /ve /d "%CLSID_ADDIN%" /f >nul

set "K_ADDIN=HKCU\Software\Classes\CLSID\%CLSID_ADDIN%"
reg add "%K_ADDIN%" /ve /d "SpecHelper.AddIn" /f >nul
reg add "%K_ADDIN%\ProgId" /ve /d "SpecHelper.AddIn" /f >nul
reg add "%K_ADDIN%\InprocServer32" /ve /d "mscoree.dll" /f >nul
reg add "%K_ADDIN%\InprocServer32" /v "ThreadingModel" /d "Both" /f >nul
reg add "%K_ADDIN%\InprocServer32" /v "Class" /d "SpecHelper.AddIn" /f >nul
reg add "%K_ADDIN%\InprocServer32" /v "Assembly" /d "SpecHelper, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null" /f >nul
reg add "%K_ADDIN%\InprocServer32" /v "RuntimeVersion" /d "v4.0.30319" /f >nul
reg add "%K_ADDIN%\InprocServer32" /v "CodeBase" /d "%CODEBASE%" /f >nul

reg add "%K_ADDIN%\InprocServer32\1.0.0.0" /v "Class" /d "SpecHelper.AddIn" /f >nul
reg add "%K_ADDIN%\InprocServer32\1.0.0.0" /v "Assembly" /d "SpecHelper, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null" /f >nul
reg add "%K_ADDIN%\InprocServer32\1.0.0.0" /v "RuntimeVersion" /d "v4.0.30319" /f >nul
reg add "%K_ADDIN%\InprocServer32\1.0.0.0" /v "CodeBase" /d "%CODEBASE%" /f >nul

reg add "HKCU\Software\Classes\SpecHelper.UI.SpecPaneControl" /ve /d "SpecHelper.UI.SpecPaneControl" /f >nul
reg add "HKCU\Software\Classes\SpecHelper.UI.SpecPaneControl\CLSID" /ve /d "%CLSID_CTRL%" /f >nul

set "K_CTRL=HKCU\Software\Classes\CLSID\%CLSID_CTRL%"
reg add "%K_CTRL%" /ve /d "SpecHelper.UI.SpecPaneControl" /f >nul
reg add "%K_CTRL%\ProgId" /ve /d "SpecHelper.UI.SpecPaneControl" /f >nul
reg add "%K_CTRL%\InprocServer32" /ve /d "mscoree.dll" /f >nul
reg add "%K_CTRL%\InprocServer32" /v "ThreadingModel" /d "Both" /f >nul
reg add "%K_CTRL%\InprocServer32" /v "Class" /d "SpecHelper.UI.SpecPaneControl" /f >nul
reg add "%K_CTRL%\InprocServer32" /v "Assembly" /d "SpecHelper, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null" /f >nul
reg add "%K_CTRL%\InprocServer32" /v "RuntimeVersion" /d "v4.0.30319" /f >nul
reg add "%K_CTRL%\InprocServer32" /v "CodeBase" /d "%CODEBASE%" /f >nul

reg add "%K_CTRL%\InprocServer32\1.0.0.0" /v "Class" /d "SpecHelper.UI.SpecPaneControl" /f >nul
reg add "%K_CTRL%\InprocServer32\1.0.0.0" /v "Assembly" /d "SpecHelper, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null" /f >nul
reg add "%K_CTRL%\InprocServer32\1.0.0.0" /v "RuntimeVersion" /d "v4.0.30319" /f >nul
reg add "%K_CTRL%\InprocServer32\1.0.0.0" /v "CodeBase" /d "%CODEBASE%" /f >nul

echo [3/3] Aktivatsiya v Excel...
set "K_EXCEL=HKCU\Software\Microsoft\Office\Excel\Addins\SpecHelper.AddIn"
reg add "%K_EXCEL%" /v "FriendlyName" /d "SpecHelper" /f >nul
reg add "%K_EXCEL%" /v "Description" /d "SpecHelper Excel AddIn" /f >nul
reg add "%K_EXCEL%" /v "LoadBehavior" /t REG_DWORD /d 3 /f >nul
reg add "%K_EXCEL%" /v "CommandLineSafe" /t REG_DWORD /d 0 /f >nul

echo.
echo ========================================================
echo        USTANOVKA USPESHNO ZAVERSHENA!
echo ========================================================
echo.
echo Otkroyte Excel - nadstroyka gotova k rabote.
echo.
pause