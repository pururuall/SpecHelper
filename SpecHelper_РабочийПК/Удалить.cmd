@echo off
setlocal

echo ========================================================
echo        UDALENIE NADSTROYKI SPECHELPER IZ EXCEL
echo ========================================================
echo.

set "CLSID_ADDIN={A19845E6-5F24-4B78-8316-D4E03C1247AA}"
set "CLSID_CTRL={D86F2C74-3208-4CE0-A4F7-1B8C512F9E82}"

echo Otkluchenie nadstroyki v Excel...
reg delete "HKCU\Software\Microsoft\Office\Excel\Addins\SpecHelper.AddIn" /f >nul 2>&1
reg delete "HKCU\Software\Wow6432Node\Microsoft\Office\Excel\Addins\SpecHelper.AddIn" /f >nul 2>&1

echo Ochistka reestra COM (32-bit i 64-bit)...
for %%B in ("HKCU\Software\Classes" "HKCU\Software\Classes\WOW6432Node") do (
    reg delete "%%~B\SpecHelper.AddIn" /f >nul 2>&1
    reg delete "%%~B\SpecHelper.UI.SpecPaneControl" /f >nul 2>&1
)

for %%B in ("HKCU\Software\Classes\CLSID" "HKCU\Software\Classes\WOW6432Node\CLSID") do (
    reg delete "%%~B\%CLSID_ADDIN%" /f >nul 2>&1
    reg delete "%%~B\%CLSID_CTRL%" /f >nul 2>&1
)

echo Nadstroyka uspeshno udalena iz reestra Excel!
echo.
pause
