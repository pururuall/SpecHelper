# Тестирование бизнес-логики SpecHelper
$ErrorActionPreference = "Stop"

$dllPath = "D:\EXCEL mibil\src\SpecHelper\bin\Release\SpecHelper.dll"
Add-Type -Path $dllPath

Write-Host "=== ТЕСТ 1: Чтение базы Материалов ===" -ForegroundColor Cyan
$matPath = "D:\EXCEL mibil\databases\Материалы.xlsx"
$matResult = [SpecHelper.Services.ExcelCatalogReader]::LoadCatalog($matPath, [SpecHelper.Models.CatalogType]::Materials)
$matRoot = $matResult.Item1
$matItems = $matResult.Item2

Write-Host "Всего позиций загружено: $($matItems.Count)"
Write-Host "Количество разделов в корне: $($matRoot.Children.Count)"
foreach ($child in $matRoot.Children) {
    Write-Host "  - Раздел: $($child.Name) (всего позиций в ветке: $($child.TotalItemCount))"
    foreach ($sub in $child.Children) {
        Write-Host "      * Подраздел: $($sub.Name) (позиций: $($sub.TotalItemCount))"
    }
}

Write-Host "`n=== ТЕСТ 2: Чтение базы Метизов ===" -ForegroundColor Cyan
$hwPath = "D:\EXCEL mibil\databases\Метизы.xlsx"
$hwResult = [SpecHelper.Services.ExcelCatalogReader]::LoadCatalog($hwPath, [SpecHelper.Models.CatalogType]::Hardware)
$hwRoot = $hwResult.Item1
$hwItems = $hwResult.Item2

Write-Host "Всего позиций загружено: $($hwItems.Count)"
foreach ($child in $hwRoot.Children) {
    Write-Host "  - Раздел: $($child.Name) (всего позиций в ветке: $($child.TotalItemCount))"
    foreach ($subChild in $child.Children) {
        Write-Host "      * Подраздел: $($subChild.Name) (позиций: $($subChild.TotalItemCount))"
    }
}

Write-Host "`n=== ТЕСТ 3: Безопасная вставка в ячейку Excel ===" -ForegroundColor Cyan
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$wb = $excel.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)

# Выбираем ячейку B5
$targetCell = $ws.Range("B5")
$targetCell.Select() | Out-Null

# Выбираем позицию для вставки (например, Болт М8х25)
$testItem = $hwItems | Where-Object { $_.Designation -like "*М8*25*" } | Select-Object -First 1
Write-Host "Выбрана позиция: $($testItem.Designation) | $($testItem.Name)"

$config = [SpecHelper.Services.ConfigManager]::Load()
$config.AutoAdvanceRow = $true

$insertRes = [SpecHelper.Services.ExcelInsertionService]::InsertItem($excel, $testItem, $config)
Write-Host "Результат вставки: Успех=$($insertRes.Item1), Сообщение=$($insertRes.Item2)"

$valB5 = $ws.Range("B5").Value2
$valC5 = $ws.Range("C5").Value2
$activeAfter = $excel.ActiveCell.Address()

Write-Host "Значение в B5 (Обозначение): $valB5"
Write-Host "Значение в C5 (Наименование): $valC5"
Write-Host "Новая активная ячейка после автосдвига: $activeAfter (ожидалась `$B`$6)"

$wb.Close($false)
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ws) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null

if ($valB5 -eq $testItem.Designation -and $valC5 -eq $testItem.Name -and $activeAfter -eq '$B$6') {
    Write-Host "`n>>> ВСЕ ТЕСТЫ ПРОЙДЕНЫ НА 100%! <<<" -ForegroundColor Green
} else {
    Write-Host "`nОШИБКА В ТЕСТАХ!" -ForegroundColor Red
}
