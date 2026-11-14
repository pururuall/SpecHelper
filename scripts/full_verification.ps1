# Полная комплексная проверка готовности SpecHelper
$ErrorActionPreference = "Stop"
$sw = [System.Diagnostics.Stopwatch]::StartNew()

Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "        КОМПЛЕКСНОЕ ТЕСТИРОВАНИЕ НАДСТРОЙКИ SPECHELPER" -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan

$passCount = 0
$failCount = 0

function Assert-Test($condition, $testName, $detail = "") {
    if ($condition) {
        $script:passCount++
        Write-Host "  [OK] $testName" -ForegroundColor Green
        if ($detail) { Write-Host "       -> $detail" -ForegroundColor Gray }
    } else {
        $script:failCount++
        Write-Host "  [FAIL] $testName" -ForegroundColor Red
        if ($detail) { Write-Host "       -> $detail" -ForegroundColor Yellow }
    }
}

# ----------------------------------------------------------------------
Write-Host "`n1. Проверка пакета для рабочего ПК (SpecHelper_РабочийПК)..." -ForegroundColor Yellow
# ----------------------------------------------------------------------
$pkgDir = "D:\EXCEL mibil\SpecHelper_РабочийПК"
Assert-Test (Test-Path $pkgDir) "Папка пакета существует: $pkgDir"

$reqDlls = @("SpecHelper.dll", "ExcelDataReader.dll", "ExcelDataReader.DataSet.dll", "Newtonsoft.Json.dll")
foreach ($dll in $reqDlls) {
    $dllPath = Join-Path $pkgDir "bin\$dll"
    Assert-Test (Test-Path $dllPath) "Библиотека $dll присутствует в bin\"
}

$reqDbs = @("Материалы.xlsx", "Метизы.xlsx", "Комплектующие.xlsx")
foreach ($db in $reqDbs) {
    $dbPath = Join-Path $pkgDir "databases\$db"
    Assert-Test (Test-Path $dbPath) "Файл базы данных $db присутствует в databases\"
}

Assert-Test (Test-Path "$pkgDir\Установить.cmd") "Скрипт установки Установить.cmd присутствует"
Assert-Test (Test-Path "$pkgDir\Удалить.cmd") "Скрипт удаления Удалить.cmd присутствует"

# ----------------------------------------------------------------------
Write-Host "`n2. Тестирование скрипта установки Установить.cmd (без прав админа)..." -ForegroundColor Yellow
# ----------------------------------------------------------------------
$installProc = Start-Process -FilePath "cmd.exe" -ArgumentList "/c `"$pkgDir\Установить.cmd`" < nul" -Wait -PassThru -NoNewWindow
Assert-Test ($installProc.ExitCode -eq 0) "Скрипт Установить.cmd выполнен без ошибок (ExitCode 0)"

$targetDir = "$env:LOCALAPPDATA\SpecHelper"
Assert-Test (Test-Path "$targetDir\SpecHelper.dll") "Файлы скопированы в профиль пользователя: $targetDir"
Assert-Test (Test-Path "$targetDir\databases\Материалы.xlsx") "Базы скопированы в $targetDir\databases"

# Проверка ключей реестра HKCU
$regAddin = Get-ItemProperty -Path "HKCU:\Software\Microsoft\Office\Excel\Addins\SpecHelper.AddIn" -ErrorAction SilentlyContinue
Assert-Test ($regAddin -ne $null -and $regAddin.LoadBehavior -eq 3) "Регистрация в HKCU Excel Addins: LoadBehavior = 3"

$clsidAddIn = "{A19845E6-5F24-4B78-8316-D4E03C1247AA}"
$regClsid = Get-ItemProperty -Path "HKCU:\Software\Classes\CLSID\$clsidAddIn\InprocServer32" -ErrorAction SilentlyContinue
Assert-Test ($regClsid -ne $null -and $regClsid.CodeBase -like "*SpecHelper.dll*") "COM-регистрация SpecHelper.AddIn в HKCU Classes"

$clsidCtrl = "{D86F2C74-3208-4CE0-A4F7-1B8C512F9E82}"
$regCtrl = Get-ItemProperty -Path "HKCU:\Software\Classes\CLSID\$clsidCtrl\InprocServer32" -ErrorAction SilentlyContinue
Assert-Test ($regCtrl -ne $null -and $regCtrl.CodeBase -like "*SpecHelper.dll*") "COM-регистрация SpecPaneControl в HKCU Classes"

# ----------------------------------------------------------------------
Write-Host "`n3. Тестирование загрузки надстройки в Excel (COM Automation)..." -ForegroundColor Yellow
# ----------------------------------------------------------------------
$logPath = "$env:APPDATA\SpecHelper\addin_debug.log"
if (Test-Path $logPath) { Remove-Item $logPath -Force }

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$addin = $excel.COMAddIns.Item("SpecHelper.AddIn")
Assert-Test ($addin -ne $null) "Надстройка SpecHelper.AddIn найдена в COMAddIns Excel"

$addin.Connect = $true
Assert-Test ($addin.Connect -eq $true) "Надстройка успешно подключена (Connect = True)"

# Проверка лога инициализации
$logContent = if (Test-Path $logPath) { Get-Content $logPath -Raw } else { "" }
Assert-Test ($logContent -like "*CreateCTP succeeded!*") "CustomTaskPane успешно создан фабрикой Excel"
Assert-Test ($logContent -like "*SpecPaneControl initialized successfully*") "Элемент SpecPaneControl успешно инициализирован"
Assert-Test ($logContent -like "*_taskPane.Visible is now true*") "Панель задач переведена в состояние Visible = True"
Assert-Test ($logContent -notlike "*FATAL*") "В логе инициализации отсутствуют критические ошибки (FATAL)"

# ----------------------------------------------------------------------
Write-Host "`n4. Тестирование движка чтения данных (ExcelDataReader)..." -ForegroundColor Yellow
# ----------------------------------------------------------------------
Add-Type -Path "$targetDir\SpecHelper.dll"

# База 1: Материалы
$matRes = [SpecHelper.Services.ExcelCatalogReader]::LoadCatalog("$targetDir\databases\Материалы.xlsx", [SpecHelper.Models.CatalogType]::Materials)
Assert-Test ($matRes.Item2.Count -ge 14) "База 'Материалы': прочитано $($matRes.Item2.Count) позиций"
Assert-Test ($matRes.Item1.Children.Count -ge 2) "База 'Материалы': корень содержит $($matRes.Item1.Children.Count) основных раздела"

# База 2: Метизы
$hwRes = [SpecHelper.Services.ExcelCatalogReader]::LoadCatalog("$targetDir\databases\Метизы.xlsx", [SpecHelper.Models.CatalogType]::Hardware)
Assert-Test ($hwRes.Item2.Count -ge 17) "База 'Метизы': прочитано $($hwRes.Item2.Count) позиций"
$hasFasteners = ($hwRes.Item1.Children | Where-Object { $_.Name -like "*Крепеж*" }) -ne $null
Assert-Test $hasFasteners "База 'Метизы': присутствует раздел 'Крепеж' с вложенными подразделами"

# База 3: Комплектующие
$compRes = [SpecHelper.Services.ExcelCatalogReader]::LoadCatalog("$targetDir\databases\Комплектующие.xlsx", [SpecHelper.Models.CatalogType]::Components)
Assert-Test ($compRes.Item2.Count -ge 9) "База 'Комплектующие': прочитано $($compRes.Item2.Count) позиций"

# ----------------------------------------------------------------------
Write-Host "`n5. Тестирование операций вставки в лист Excel..." -ForegroundColor Yellow
# ----------------------------------------------------------------------
$wb = $excel.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)

# Тест 5.1: Обычная вставка в ячейку B5
$ws.Range("B5").Select() | Out-Null
$testBolt = $hwRes.Item2 | Where-Object { $_.Designation -like "*М6*20*" } | Select-Object -First 1
$config = [SpecHelper.Services.ConfigManager]::Load()
$config.AutoAdvanceRow = $true

$resNormal = [SpecHelper.Services.ExcelInsertionService]::InsertItem($excel, $testBolt, $config)
Assert-Test ($resNormal.Item1 -eq $true) "Вставка позиции завершилась успехом" "Сообщение: $($resNormal.Item2)"

$b5Val = $ws.Range("B5").Value2
$c5Val = $ws.Range("C5").Value2
$activeAfter = $excel.ActiveCell.Address()

Assert-Test ($b5Val -eq $testBolt.Designation) "Ячейка B5 получила Обозначение: $b5Val"
Assert-Test ($c5Val -eq $testBolt.Name) "Ячейка C5 получила Наименование: $c5Val"
Assert-Test ($activeAfter -eq '$B$6') "Курсор автоматически перешел на строку ниже (B6)"

# Тест 5.2: Штатная отмена через Undo (Ctrl + Z)
$excel.Undo()
$b5AfterUndo = $ws.Range("B5").Value2
$c5AfterUndo = $ws.Range("C5").Value2
Assert-Test ($b5AfterUndo -eq $null -and $c5AfterUndo -eq $null) "Штатная отмена (Undo / Ctrl+Z) успешно откатила вставку" "B5='$b5AfterUndo', C5='$c5AfterUndo'"

# Повторная вставка для проверки защиты листа
$ws.Range("B5").Select() | Out-Null
[SpecHelper.Services.ExcelInsertionService]::InsertItem($excel, $testBolt, $config) | Out-Null

# Тест 5.3: Защита от защищенных листов (ProtectContents)
$ws.Protect("test_password") | Out-Null
$resProtected = [SpecHelper.Services.ExcelInsertionService]::InsertItem($excel, $testBolt, $config)
Assert-Test ($resProtected.Item1 -eq $false -and $resProtected.Item2 -like "*защищен*") "Корректная обработка защищенного листа (без падения)" "Сообщение: $($resProtected.Item2)"

$wb.Close($false)
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ws) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null

# ----------------------------------------------------------------------
Write-Host "`n6. Тестирование скрипта удаления Удалить.cmd..." -ForegroundColor Yellow
# ----------------------------------------------------------------------
$uninstProc = Start-Process -FilePath "cmd.exe" -ArgumentList "/c `"$pkgDir\Удалить.cmd`" < nul" -Wait -PassThru -NoNewWindow
Assert-Test ($uninstProc.ExitCode -eq 0) "Скрипт Удалить.cmd выполнен успешно"

$regAfterUninst = Get-ItemProperty -Path "HKCU:\Software\Microsoft\Office\Excel\Addins\SpecHelper.AddIn" -ErrorAction SilentlyContinue
Assert-Test ($regAfterUninst -eq $null) "Регистрация надстройки удалена из реестра Excel"

# Повторно накатываем Установить.cmd, чтобы оставить систему в готовом рабочем состоянии
$reinstallProc = Start-Process -FilePath "cmd.exe" -ArgumentList "/c `"$pkgDir\Установить.cmd`" < nul" -Wait -PassThru -NoNewWindow
Assert-Test ($reinstallProc.ExitCode -eq 0) "Повторный запуск Установить.cmd восстановил готовность системы"

$sw.Stop()

# ----------------------------------------------------------------------
Write-Host "`n================================================================" -ForegroundColor Cyan
Write-Host "ИТОГ ПРОВЕРКИ:" -ForegroundColor Cyan
Write-Host "  Пройдено проверок: $passCount" -ForegroundColor Green
Write-Host "  Провалено проверок: $failCount" -ForegroundColor $(if ($failCount -eq 0) { "Green" } else { "Red" })
Write-Host "  Общее время тестирования: $($sw.Elapsed.TotalSeconds.ToString('F2')) сек" -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
