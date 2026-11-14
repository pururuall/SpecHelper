# Скрипт удаления регистрации надстройки SpecHelper
$ErrorActionPreference = "SilentlyContinue"

Write-Host "Удаление регистрации SpecHelper Add-In..." -ForegroundColor Yellow

$clsidAddIn = "{A19845E6-5F24-4B78-8316-D4E03C1247AA}"
$clsidControl = "{D86F2C74-3208-4CE0-A4F7-1B8C512F9E82}"

Remove-Item -Path "HKCU:\Software\Microsoft\Office\Excel\Addins\SpecHelper.AddIn" -Recurse -Force
Remove-Item -Path "HKCU:\Software\Classes\SpecHelper.AddIn" -Recurse -Force
Remove-Item -Path "HKCU:\Software\Classes\CLSID\$clsidAddIn" -Recurse -Force
Remove-Item -Path "HKCU:\Software\Classes\SpecHelper.UI.SpecPaneControl" -Recurse -Force
Remove-Item -Path "HKCU:\Software\Classes\CLSID\$clsidControl" -Recurse -Force

Write-Host "Надстройка SpecHelper успешно удалена из реестра." -ForegroundColor Green
