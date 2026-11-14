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
reg delete "HKCU\Software\Classes\SpecHelper.AddIn" /f >nul 2>&1
reg delete "HKCU\Software\Classes\CLSID\%CLSID_ADDIN%" /f >nul 2>&1
reg delete "HKCU\Software\Classes\SpecHelper.UI.SpecPaneControl" /f >nul 2>&1
reg delete "HKCU\Software\Classes\CLSID\%CLSID_CTRL%" /f >nul 2>&1

echo Nadstroyka uspeshno udalena iz reestra Excel!
echo.
pause