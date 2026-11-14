# Скрипт регистрации COM надстройки SpecHelper для Excel в HKCU (без необходимости прав администратора)
$ErrorActionPreference = "Stop"

$dllPath = "D:\EXCEL mibil\src\SpecHelper\bin\Release\SpecHelper.dll"
if (-not (Test-Path $dllPath)) {
    throw "Файл надстройки не найден: $dllPath. Сначала выполните сборку проекта."
}

Write-Host "1. Генерация regasm .reg файла..." -ForegroundColor Cyan
$regasm = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe"
$tempReg = "D:\EXCEL mibil\scripts\temp_classes.reg"
$hkcuReg = "D:\EXCEL mibil\scripts\hkcu_classes.reg"

& $regasm $dllPath /codebase /regfile:$tempReg | Out-Null

if (-not (Test-Path $tempReg)) {
    throw "Не удалось сгенерировать файл регистрации."
}

Write-Host "2. Адаптация путей реестра для HKEY_CURRENT_USER..." -ForegroundColor Cyan
$regContent = Get-Content $tempReg -Raw -Encoding Default
# Заменяем HKEY_CLASSES_ROOT на HKEY_CURRENT_USER\Software\Classes
$hkcuContent = $regContent -replace '\[HKEY_CLASSES_ROOT\\', '[HKEY_CURRENT_USER\Software\Classes\'

# Добавляем регистрацию в Excel Addins
$excelAddinReg = @"

[HKEY_CURRENT_USER\Software\Microsoft\Office\Excel\Addins\SpecHelper.AddIn]
"FriendlyName"="Справочник спецификаций (SpecHelper)"
"Description"="Вставка позиций спецификации из баз данных Excel"
"LoadBehavior"=dword:00000003
"CommandLineSafe"=dword:00000000

"@

$finalContent = $hkcuContent + "`r`n" + $excelAddinReg
[System.IO.File]::WriteAllText($hkcuReg, $finalContent, [System.Text.Encoding]::Unicode)

Write-Host "3. Импорт в реестр текущего пользователя..." -ForegroundColor Cyan
reg import $hkcuReg

Remove-Item $tempReg -Force -ErrorAction SilentlyContinue

Write-Host "Надстройка SpecHelper успешно зарегистрирована в реестре текущего пользователя!" -ForegroundColor Green
