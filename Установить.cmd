@echo off
setlocal

echo ========================================================
echo        USTANOVKA NADSTROYKI SPECHELPER DLYA EXCEL
echo            (Prava administratora NE TREBUYUTSYA)
echo ========================================================
echo.

set "BIN_DIR=%~dp0bin"
set "DB_DIR=%~dp0databases"
set "TARGET_DIR=%LOCALAPPDATA%\SpecHelper"

rem 1. Proverka nalichiya sobrannyh bibliotek
if not exist "%BIN_DIR%\SpecHelper.dll" (
    echo OSHIBKA: Sobrannye fayly nadstroyki ne naydeny!
    echo Fayl "%BIN_DIR%\SpecHelper.dll" otsutstvuet.
    echo Ubedites, chto papka "bin" nahoditsya ryadom s ustanovshchikom.
    echo.
    pause
    exit /b 1
)

echo [1/3] Kopirovanie faylov v: %TARGET_DIR%...
if not exist "%TARGET_DIR%" mkdir "%TARGET_DIR%"
if not exist "%TARGET_DIR%\databases" mkdir "%TARGET_DIR%\databases"

rem Udalenie vozmozhnyh konfliktuyushchih sistemnyh bibliotek .NET
if exist "%TARGET_DIR%\mscorlib.dll" del /f /q "%TARGET_DIR%\mscorlib.dll" >nul 2>&1
if exist "%TARGET_DIR%\System.*.dll" del /f /q "%TARGET_DIR%\System.*.dll" >nul 2>&1

rem Kopirovanie tolko razreshennyh bibliotek nadstroyki
if exist "%BIN_DIR%\SpecHelper.dll" copy /y "%BIN_DIR%\SpecHelper.dll" "%TARGET_DIR%\" >nul
if exist "%BIN_DIR%\ExcelDataReader.dll" copy /y "%BIN_DIR%\ExcelDataReader.dll" "%TARGET_DIR%\" >nul
if exist "%BIN_DIR%\ExcelDataReader.DataSet.dll" copy /y "%BIN_DIR%\ExcelDataReader.DataSet.dll" "%TARGET_DIR%\" >nul
if exist "%BIN_DIR%\Newtonsoft.Json.dll" copy /y "%BIN_DIR%\Newtonsoft.Json.dll" "%TARGET_DIR%\" >nul
if exist "%BIN_DIR%\OFFICE.DLL" copy /y "%BIN_DIR%\OFFICE.DLL" "%TARGET_DIR%\" >nul
if exist "%BIN_DIR%\Microsoft.Office.Interop.Excel.dll" copy /y "%BIN_DIR%\Microsoft.Office.Interop.Excel.dll" "%TARGET_DIR%\" >nul
if exist "%BIN_DIR%\extensibility.dll" copy /y "%BIN_DIR%\extensibility.dll" "%TARGET_DIR%\" >nul

if exist "%DB_DIR%" copy /y "%DB_DIR%\*.xlsx" "%TARGET_DIR%\databases\" >nul

echo [2/3] Registratsiya v reestre HKCU (32-bit i 64-bit Office)...

set "CLSID_ADDIN={A19845E6-5F24-4B78-8316-D4E03C1247AA}"
set "CLSID_CTRL={D86F2C74-3208-4CE0-A4F7-1B8C512F9E82}"
set "SLASH_DIR=%TARGET_DIR:\=/%"
set "CODEBASE=file:///%SLASH_DIR%/SpecHelper.dll"

rem Registratsiya ProgID dlya 64-bit i 32-bit
for %%B in ("HKCU\Software\Classes" "HKCU\Software\Classes\WOW6432Node") do (
    reg add "%%~B\SpecHelper.AddIn" /ve /d "SpecHelper.AddIn" /f >nul
    reg add "%%~B\SpecHelper.AddIn\CLSID" /ve /d "%CLSID_ADDIN%" /f >nul
    reg add "%%~B\SpecHelper.UI.SpecPaneControl" /ve /d "SpecHelper.UI.SpecPaneControl" /f >nul
    reg add "%%~B\SpecHelper.UI.SpecPaneControl\CLSID" /ve /d "%CLSID_CTRL%" /f >nul
)

rem Registratsiya CLSID AddIn v obe vetki
for %%B in ("HKCU\Software\Classes\CLSID" "HKCU\Software\Classes\WOW6432Node\CLSID") do (
    reg add "%%~B\%CLSID_ADDIN%" /ve /d "SpecHelper.AddIn" /f >nul
    reg add "%%~B\%CLSID_ADDIN%\ProgId" /ve /d "SpecHelper.AddIn" /f >nul
    reg add "%%~B\%CLSID_ADDIN%\InprocServer32" /ve /d "mscoree.dll" /f >nul
    reg add "%%~B\%CLSID_ADDIN%\InprocServer32" /v "ThreadingModel" /d "Both" /f >nul
    reg add "%%~B\%CLSID_ADDIN%\InprocServer32" /v "Class" /d "SpecHelper.AddIn" /f >nul
    reg add "%%~B\%CLSID_ADDIN%\InprocServer32" /v "Assembly" /d "SpecHelper, Version=1.1.0.0, Culture=neutral, PublicKeyToken=null" /f >nul
    reg add "%%~B\%CLSID_ADDIN%\InprocServer32" /v "RuntimeVersion" /d "v4.0.30319" /f >nul
    reg add "%%~B\%CLSID_ADDIN%\InprocServer32" /v "CodeBase" /d "%CODEBASE%" /f >nul
    reg add "%%~B\%CLSID_ADDIN%\InprocServer32\1.0.0.0" /v "Class" /d "SpecHelper.AddIn" /f >nul
    reg add "%%~B\%CLSID_ADDIN%\InprocServer32\1.0.0.0" /v "Assembly" /d "SpecHelper, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null" /f >nul
    reg add "%%~B\%CLSID_ADDIN%\InprocServer32\1.0.0.0" /v "RuntimeVersion" /d "v4.0.30319" /f >nul
    reg add "%%~B\%CLSID_ADDIN%\InprocServer32\1.0.0.0" /v "CodeBase" /d "%CODEBASE%" /f >nul
    reg add "%%~B\%CLSID_ADDIN%\InprocServer32\1.1.0.0" /v "Class" /d "SpecHelper.AddIn" /f >nul
    reg add "%%~B\%CLSID_ADDIN%\InprocServer32\1.1.0.0" /v "Assembly" /d "SpecHelper, Version=1.1.0.0, Culture=neutral, PublicKeyToken=null" /f >nul
    reg add "%%~B\%CLSID_ADDIN%\InprocServer32\1.1.0.0" /v "RuntimeVersion" /d "v4.0.30319" /f >nul
    reg add "%%~B\%CLSID_ADDIN%\InprocServer32\1.1.0.0" /v "CodeBase" /d "%CODEBASE%" /f >nul
)

rem Registratsiya CLSID SpecPaneControl s obyazatelnymi markerami ActiveX v obe vetki
for %%B in ("HKCU\Software\Classes\CLSID" "HKCU\Software\Classes\WOW6432Node\CLSID") do (
    reg add "%%~B\%CLSID_CTRL%" /ve /d "SpecHelper.UI.SpecPaneControl" /f >nul
    reg add "%%~B\%CLSID_CTRL%\ProgId" /ve /d "SpecHelper.UI.SpecPaneControl" /f >nul
    reg add "%%~B\%CLSID_CTRL%\InprocServer32" /ve /d "mscoree.dll" /f >nul
    reg add "%%~B\%CLSID_CTRL%\InprocServer32" /v "ThreadingModel" /d "Both" /f >nul
    reg add "%%~B\%CLSID_CTRL%\InprocServer32" /v "Class" /d "SpecHelper.UI.SpecPaneControl" /f >nul
    reg add "%%~B\%CLSID_CTRL%\InprocServer32" /v "Assembly" /d "SpecHelper, Version=1.1.0.0, Culture=neutral, PublicKeyToken=null" /f >nul
    reg add "%%~B\%CLSID_CTRL%\InprocServer32" /v "RuntimeVersion" /d "v4.0.30319" /f >nul
    reg add "%%~B\%CLSID_CTRL%\InprocServer32" /v "CodeBase" /d "%CODEBASE%" /f >nul
    reg add "%%~B\%CLSID_CTRL%\InprocServer32\1.0.0.0" /v "Class" /d "SpecHelper.UI.SpecPaneControl" /f >nul
    reg add "%%~B\%CLSID_CTRL%\InprocServer32\1.0.0.0" /v "Assembly" /d "SpecHelper, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null" /f >nul
    reg add "%%~B\%CLSID_CTRL%\InprocServer32\1.0.0.0" /v "RuntimeVersion" /d "v4.0.30319" /f >nul
    reg add "%%~B\%CLSID_CTRL%\InprocServer32\1.0.0.0" /v "CodeBase" /d "%CODEBASE%" /f >nul
    reg add "%%~B\%CLSID_CTRL%\InprocServer32\1.1.0.0" /v "Class" /d "SpecHelper.UI.SpecPaneControl" /f >nul
    reg add "%%~B\%CLSID_CTRL%\InprocServer32\1.1.0.0" /v "Assembly" /d "SpecHelper, Version=1.1.0.0, Culture=neutral, PublicKeyToken=null" /f >nul
    reg add "%%~B\%CLSID_CTRL%\InprocServer32\1.1.0.0" /v "RuntimeVersion" /d "v4.0.30319" /f >nul
    reg add "%%~B\%CLSID_CTRL%\InprocServer32\1.1.0.0" /v "CodeBase" /d "%CODEBASE%" /f >nul
    reg add "%%~B\%CLSID_CTRL%\Control" /f >nul
    reg add "%%~B\%CLSID_CTRL%\Implemented Categories\{40FC6ED4-2438-11CF-A3DB-080036F12502}" /f >nul
    reg add "%%~B\%CLSID_CTRL%\Implemented Categories\{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}" /f >nul
)

echo [3/3] Aktivatsiya v Excel (LoadBehavior = 3)...
for %%K in ("HKCU\Software\Microsoft\Office\Excel\Addins\SpecHelper.AddIn" "HKCU\Software\Wow6432Node\Microsoft\Office\Excel\Addins\SpecHelper.AddIn") do (
    reg add "%%~K" /v "FriendlyName" /d "SpecHelper" /f >nul
    reg add "%%~K" /v "Description" /d "SpecHelper Excel AddIn" /f >nul
    reg add "%%~K" /v "LoadBehavior" /t REG_DWORD /d 3 /f >nul
    reg add "%%~K" /v "CommandLineSafe" /t REG_DWORD /d 0 /f >nul
)

echo.
echo ========================================================
echo        USTANOVKA USPESHNO ZAVERSHENA!
echo ========================================================
echo.
echo Otkroyte Excel - nadstroyka gotova k rabote.
echo.
pause
