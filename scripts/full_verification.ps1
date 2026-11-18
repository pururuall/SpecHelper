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

$rootDir = (Resolve-Path "$PSScriptRoot\..").Path
$pkgDir = Join-Path $rootDir "SpecHelper_РабочийПК"
$targetDir = "$env:LOCALAPPDATA\SpecHelper"

# ----------------------------------------------------------------------
Write-Host "`n1. Проверка структуры автономных дистрибутивов..." -ForegroundColor Yellow
# ----------------------------------------------------------------------
Assert-Test (Test-Path $pkgDir) "Папка пакета существует: $pkgDir"
Assert-Test (Test-Path (Join-Path $rootDir "bin")) "Корневая папка bin существует: $rootDir\bin"

$reqDlls = @(
    "SpecHelper.dll",
    "ExcelDataReader.dll",
    "ExcelDataReader.DataSet.dll",
    "Newtonsoft.Json.dll",
    "OFFICE.DLL",
    "Microsoft.Office.Interop.Excel.dll",
    "extensibility.dll"
)

foreach ($dll in $reqDlls) {
    $pDll = Join-Path $pkgDir "bin\$dll"
    $rDll = Join-Path $rootDir "bin\$dll"
    Assert-Test (Test-Path $pDll) "Библиотека $dll присутствует в SpecHelper_РабочийПК\bin"
    Assert-Test (Test-Path $rDll) "Библиотека $dll присутствует в корневой bin\"
}

# Проверка отсутствия системных библиотек .NET (mscorlib.dll, System.*.dll)
$bannedP = Get-ChildItem -Path (Join-Path $pkgDir "bin") -Include "mscorlib.dll", "System.*.dll" -Recurse
$bannedR = Get-ChildItem -Path (Join-Path $rootDir "bin") -Include "mscorlib.dll", "System.*.dll" -Recurse
Assert-Test ($bannedP.Count -eq 0) "В SpecHelper_РабочийПК\bin отсутствуют mscorlib.dll и System.*.dll"
Assert-Test ($bannedR.Count -eq 0) "В корневой bin\ отсутствуют mscorlib.dll и System.*.dll"

$reqDbs = @("Материалы.xlsx", "Метизы.xlsx", "Комплектующие.xlsx", "Шаблоны.xlsx")
foreach ($db in $reqDbs) {
    $dbPath = Join-Path $pkgDir "databases\$db"
    Assert-Test (Test-Path $dbPath) "Файл базы данных $db присутствует в databases\"
}

Assert-Test (Test-Path "$pkgDir\Установить.cmd") "Скрипт установки Установить.cmd присутствует в SpecHelper_РабочийПК"
Assert-Test (Test-Path "$pkgDir\Удалить.cmd") "Скрипт удаления Удалить.cmd присутствует в SpecHelper_РабочийПК"
Assert-Test (Test-Path "$rootDir\Установить.cmd") "Скрипт установки Установить.cmd присутствует в корне проекта"
Assert-Test (Test-Path "$rootDir\Удалить.cmd") "Скрипт удаления Удалить.cmd присутствует в корне проекта"

# ----------------------------------------------------------------------
Write-Host "`n2. Тестирование скрипта установки Установить.cmd (без прав админа)..." -ForegroundColor Yellow
# ----------------------------------------------------------------------
$installProc = Start-Process -FilePath "cmd.exe" -ArgumentList "/c `"`"$pkgDir\Установить.cmd`" < nul`"" -Wait -PassThru -NoNewWindow
Assert-Test ($installProc.ExitCode -eq 0) "Скрипт Установить.cmd выполнен без ошибок (ExitCode 0)"

Assert-Test (Test-Path "$targetDir\SpecHelper.dll") "Файлы скопированы в профиль пользователя: $targetDir"
Assert-Test (Test-Path "$targetDir\databases\Материалы.xlsx") "Базы скопированы в $targetDir\databases"

# Проверка отсутствия mscorlib.dll в целевой папке
$bannedTarget = Get-ChildItem -Path $targetDir -Include "mscorlib.dll", "System.*.dll" -Recurse
Assert-Test ($bannedTarget.Count -eq 0) "В целевой папке установки (%LOCALAPPDATA%\SpecHelper) отсутствуют mscorlib.dll и System.*.dll"

# Проверка ключей реестра HKCU (32-bit и 64-bit ветки)
$regAddin = Get-ItemProperty -Path "HKCU:\Software\Microsoft\Office\Excel\Addins\SpecHelper.AddIn" -ErrorAction SilentlyContinue
Assert-Test ($regAddin -ne $null -and $regAddin.LoadBehavior -eq 3) "Регистрация в HKCU Excel Addins: LoadBehavior = 3"

$regAddinWow = Get-ItemProperty -Path "HKCU:\Software\Wow6432Node\Microsoft\Office\Excel\Addins\SpecHelper.AddIn" -ErrorAction SilentlyContinue
Assert-Test ($regAddinWow -ne $null -and $regAddinWow.LoadBehavior -eq 3) "Регистрация в HKCU Wow6432Node Excel Addins: LoadBehavior = 3"

$clsidAddIn = "{A19845E6-5F24-4B78-8316-D4E03C1247AA}"
$regClsidAddin = Get-ItemProperty -Path "HKCU:\Software\Classes\CLSID\$clsidAddIn\InprocServer32" -ErrorAction SilentlyContinue
Assert-Test ($regClsidAddin -ne $null -and $regClsidAddin.CodeBase -like "*SpecHelper.dll*") "COM-регистрация AddIn в HKCU\Software\Classes\CLSID"

$regClsidAddinWow = Get-ItemProperty -Path "HKCU:\Software\Classes\WOW6432Node\CLSID\$clsidAddIn\InprocServer32" -ErrorAction SilentlyContinue
Assert-Test ($regClsidAddinWow -ne $null -and $regClsidAddinWow.CodeBase -like "*SpecHelper.dll*") "COM-регистрация AddIn в HKCU\Software\Classes\WOW6432Node\CLSID"

$clsidCtrl = "{D86F2C74-3208-4CE0-A4F7-1B8C512F9E82}"
$regCtrl = Get-ItemProperty -Path "HKCU:\Software\Classes\CLSID\$clsidCtrl\InprocServer32" -ErrorAction SilentlyContinue
Assert-Test ($regCtrl -ne $null -and $regCtrl.CodeBase -like "*SpecHelper.dll*") "COM-регистрация SpecPaneControl в HKCU\Software\Classes\CLSID"

$regCtrlWow = Get-ItemProperty -Path "HKCU:\Software\Classes\WOW6432Node\CLSID\$clsidCtrl\InprocServer32" -ErrorAction SilentlyContinue
Assert-Test ($regCtrlWow -ne $null -and $regCtrlWow.CodeBase -like "*SpecHelper.dll*") "COM-регистрация SpecPaneControl в HKCU\Software\Classes\WOW6432Node\CLSID"

# Проверка обязательных ключей ActiveX-контрола (Control и Implemented Categories)
$ctrlKey = "HKCU:\Software\Classes\CLSID\$clsidCtrl\Control"
$cat1Key = "HKCU:\Software\Classes\CLSID\$clsidCtrl\Implemented Categories\{40FC6ED4-2438-11CF-A3DB-080036F12502}"
$cat2Key = "HKCU:\Software\Classes\CLSID\$clsidCtrl\Implemented Categories\{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}"
Assert-Test (Test-Path $ctrlKey) "Присутствует маркер ActiveX Control в CLSID"
Assert-Test (Test-Path $cat1Key) "Присутствует маркер Implemented Categories {40FC6ED4-...} в CLSID"
Assert-Test (Test-Path $cat2Key) "Присутствует маркер Implemented Categories {62C8FE65-...} в CLSID"

$ctrlKeyWow = "HKCU:\Software\Classes\WOW6432Node\CLSID\$clsidCtrl\Control"
$cat1KeyWow = "HKCU:\Software\Classes\WOW6432Node\CLSID\$clsidCtrl\Implemented Categories\{40FC6ED4-2438-11CF-A3DB-080036F12502}"
$cat2KeyWow = "HKCU:\Software\Classes\WOW6432Node\CLSID\$clsidCtrl\Implemented Categories\{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}"
Assert-Test (Test-Path $ctrlKeyWow) "Присутствует маркер ActiveX Control в WOW6432Node\CLSID"
Assert-Test (Test-Path $cat1KeyWow) "Присутствует маркер Implemented Categories {40FC6ED4-...} в WOW6432Node\CLSID"
Assert-Test (Test-Path $cat2KeyWow) "Присутствует маркер Implemented Categories {62C8FE65-...} в WOW6432Node\CLSID"

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
Assert-Test ($logContent -like "*CreateCTP succeeded!*" -or $logContent -like "*Fallback floating window*") "Панель или резервное плавающее окно успешно инициализировано"
Assert-Test ($logContent -like "*SpecPaneControl initialized successfully*" -or $logContent -like "*Fallback floating window created successfully*") "SpecPaneControl успешно инициализирован"
Assert-Test ($logContent -notlike "*FATAL*") "В логе инициализации отсутствуют критические ошибки (FATAL)"

# ----------------------------------------------------------------------
Write-Host "`n4. Тестирование движка чтения данных (ExcelDataReader)..." -ForegroundColor Yellow
# ----------------------------------------------------------------------
Add-Type -Path "$targetDir\SpecHelper.dll"

# База 1: Материалы
$matRes = [SpecHelper.Services.ExcelCatalogReader]::LoadCatalog("$targetDir\databases\Материалы.xlsx", [SpecHelper.Models.CatalogType]::Materials)
Assert-Test ($matRes.Item2.Count -ge 14) "База 'Материалы': прочитано $($matRes.Item2.Count) позиций"
Assert-Test ($matRes.Item1.Children.Count -ge 1) "База 'Материалы': дерево каталогов успешно сформировано"

# База 2: Метизы
$hwRes = [SpecHelper.Services.ExcelCatalogReader]::LoadCatalog("$targetDir\databases\Метизы.xlsx", [SpecHelper.Models.CatalogType]::Hardware)
Assert-Test ($hwRes.Item2.Count -ge 17) "База 'Метизы': прочитано $($hwRes.Item2.Count) позиций"
$hasFasteners = ($hwRes.Item1.Children | Where-Object { $_.Name -like "*Креп*" }) -ne $null
Assert-Test $hasFasteners "База 'Метизы': присутствует раздел 'Крепёж' / 'Крепеж' с вложенными подразделами"

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
$uninstProc = Start-Process -FilePath "cmd.exe" -ArgumentList "/c `"`"$pkgDir\Удалить.cmd`" < nul`"" -Wait -PassThru -NoNewWindow
Assert-Test ($uninstProc.ExitCode -eq 0) "Скрипт Удалить.cmd выполнен успешно"

$regAfterUninst = Get-ItemProperty -Path "HKCU:\Software\Microsoft\Office\Excel\Addins\SpecHelper.AddIn" -ErrorAction SilentlyContinue
Assert-Test ($regAfterUninst -eq $null) "Регистрация надстройки удалена из реестра Excel (HKCU)"

$regAfterUninstWow = Get-ItemProperty -Path "HKCU:\Software\Wow6432Node\Microsoft\Office\Excel\Addins\SpecHelper.AddIn" -ErrorAction SilentlyContinue
Assert-Test ($regAfterUninstWow -eq $null) "Регистрация надстройки удалена из реестра Excel (WOW6432Node)"

$regCtrlAfter = Get-ItemProperty -Path "HKCU:\Software\Classes\CLSID\$clsidCtrl" -ErrorAction SilentlyContinue
Assert-Test ($regCtrlAfter -eq $null) "CLSID SpecPaneControl удален из реестра (CLSID)"

$regCtrlWowAfter = Get-ItemProperty -Path "HKCU:\Software\Classes\WOW6432Node\CLSID\$clsidCtrl" -ErrorAction SilentlyContinue
Assert-Test ($regCtrlWowAfter -eq $null) "CLSID SpecPaneControl удален из реестра (WOW6432Node\CLSID)"

# Повторно накатываем Установить.cmd, чтобы оставить систему в готовом рабочем состоянии
$reinstallProc = Start-Process -FilePath "cmd.exe" -ArgumentList "/c `"`"$pkgDir\Установить.cmd`" < nul`"" -Wait -PassThru -NoNewWindow
Assert-Test ($reinstallProc.ExitCode -eq 0) "Повторный запуск Установить.cmd восстановил готовность системы"

# ----------------------------------------------------------------------
Write-Host "`n7. Тестирование службы автообновления (UpdateService)..." -ForegroundColor Yellow
# ----------------------------------------------------------------------
$asm = [System.Reflection.Assembly]::LoadFrom("$targetDir\SpecHelper.dll")
$asmVer = $asm.GetName().Version.ToString()
Assert-Test ($asmVer -eq "1.0.0.0") "AssemblyVersion остался 1.0.0.0 для обратной совместимости с COM в реестре"

$fileVerAttr = [Attribute]::GetCustomAttribute($asm, [System.Reflection.AssemblyFileVersionAttribute])
$fileVer = if ($fileVerAttr) { $fileVerAttr.Version } else { "" }
Assert-Test ($fileVer -eq "1.0.0.0") "AssemblyFileVersion равен 1.0.0.0"

$currServiceVer = [SpecHelper.Services.UpdateService]::GetCurrentAssemblyFileVersion()
Assert-Test ($currServiceVer.ToString() -eq "1.0.0.0") "UpdateService.GetCurrentAssemblyFileVersion() вернул $currServiceVer"

$swOffline = [System.Diagnostics.Stopwatch]::StartNew()
$offlineCheck = [SpecHelper.Services.UpdateService]::CheckForUpdates($null, "https://192.0.2.1/version.json")
$swOffline.Stop()
Assert-Test ($offlineCheck -eq $false -and $swOffline.Elapsed.TotalSeconds -le 7) "Офлайн-проверка (таймаут 5 сек) завершается тихо без исключений ($($swOffline.Elapsed.TotalSeconds.ToString('F2')) сек)"

# ----------------------------------------------------------------------
Write-Host "`n8. Тестирование модуля шаблонов (TemplateService и вкладка 'Шаблоны')..." -ForegroundColor Yellow
# ----------------------------------------------------------------------
$templatesPath = [SpecHelper.Services.TemplateService]::GetTemplatesFilePath($null)
Assert-Test (Test-Path $templatesPath) "Файл шаблонов успешно обнаружен: $templatesPath"

$templatesList = [SpecHelper.Services.TemplateService]::LoadTemplates($templatesPath)
Assert-Test ($templatesList.Count -eq 3) "Считано ровно 3 демонстрационных шаблона (фактически: $($templatesList.Count))"

$tpl1 = $templatesList | Where-Object { $_.Name -eq "Коробка распашная" }
Assert-Test ($tpl1 -ne $null -and $tpl1.StartRow -eq 6 -and $tpl1.EndRow -eq 10 -and $tpl1.RowCount -eq 5) "Шаблон 'Коробка распашная' распознан (строки 6..10, 5 строк)"

$tpl2 = $templatesList | Where-Object { $_.Name -eq "Тумба подвесная 600" }
Assert-Test ($tpl2 -ne $null -and $tpl2.StartRow -eq 16 -and $tpl2.EndRow -eq 20 -and $tpl2.RowCount -eq 5) "Шаблон 'Тумба подвесная 600' распознан (строки 16..20, 5 строк)"

$tpl3 = $templatesList | Where-Object { $_.Name -eq "Ящик выдвижной 500" }
Assert-Test ($tpl3 -ne $null -and $tpl3.StartRow -eq 26 -and $tpl3.EndRow -eq 30 -and $tpl3.RowCount -eq 5) "Шаблон 'Ящик выдвижной 500' распознан (строки 26..30, 5 строк)"

# Тестирование нативной вставки шаблона в рабочий лист Excel
$excelTpl = New-Object -ComObject Excel.Application
$excelTpl.Visible = $false
$excelTpl.DisplayAlerts = $false
$wbTest = $excelTpl.Workbooks.Add()
$wsTest = $wbTest.Worksheets.Item(1)
$wsTest.Cells.Item(1, 1).Value2 = "Обозначение"
$wsTest.Cells.Item(1, 2).Value2 = "Наименование"
$wsTest.Cells.Item(2, 1).Value2 = "Существующая строка внизу"

# Выбираем ячейку A2 (вставка должна сдвинуть существующую строку вниз)
$targetCell = $wsTest.Cells.Item(2, 1)
$targetCell.Select() | Out-Null

$insertResult = [SpecHelper.Services.TemplateService]::InsertTemplate($excelTpl, $templatesPath, $tpl1)
Assert-Test ($insertResult.Item1 -eq $true) "Вставка шаблона в Excel завершилась успехом" "Сообщение: $($insertResult.Item2)"

$cellName = $wsTest.Cells.Item(2, 2).Value2
Assert-Test ($cellName -eq "Боковина короба 720х300") "Ячейка B2 содержит наименование детали: '$cellName'"

$formulaMult = $wsTest.Cells.Item(2, 5).Formula
Assert-Test ($formulaMult -eq "=C2*D2") "Формула умножения перенесена и скорректирована: '$formulaMult'"

$formulaSum = $wsTest.Cells.Item(6, 5).Formula
Assert-Test ($formulaSum -eq "=SUM(E2:E5)") "Итоговая формула суммы перенесена и скорректирована: '$formulaSum'"

$shiftedVal = $wsTest.Cells.Item(7, 1).Value2
Assert-Test ($shiftedVal -eq "Существующая строка внизу") "Существующие данные сдвинуты вниз на строку 7: '$shiftedVal'"

# Тест 8.2: Повторная мгновенная вставка из кэша (No Re-open Lag)
$swFast = [System.Diagnostics.Stopwatch]::StartNew()
$insertResult2 = [SpecHelper.Services.TemplateService]::InsertTemplate($excelTpl, $templatesPath, $tpl2)
$swFast.Stop()
Assert-Test ($insertResult2.Item1 -eq $true -and $swFast.ElapsedMilliseconds -lt 500) "Повторная вставка выполнена без задержки открытия с диска ($($swFast.ElapsedMilliseconds) мс)"

# Тест 8.3: Безопасное закрытие фоновой книги
[SpecHelper.Services.TemplateService]::CloseBackgroundWorkbook()
Assert-Test ($true) "Метод CloseBackgroundWorkbook() успешно освободил фоновую книгу без ошибок"

$wbTest.Close($false)
$excelTpl.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excelTpl) | Out-Null

$sw.Stop()

# ----------------------------------------------------------------------
Write-Host "`n================================================================" -ForegroundColor Cyan
Write-Host "ИТОГ ПРОВЕРКИ:" -ForegroundColor Cyan
Write-Host "  Пройдено проверок: $passCount" -ForegroundColor Green
Write-Host "  Провалено проверок: $failCount" -ForegroundColor $(if ($failCount -eq 0) { "Green" } else { "Red" })
Write-Host "  Общее время тестирования: $($sw.Elapsed.TotalSeconds.ToString('F2')) сек" -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
