# Скрипт полной автономной сборки проекта SpecHelper
$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$rootDir = (Resolve-Path "$scriptDir\..").Path
$solutionPath = Join-Path $rootDir "src\SpecHelper.sln"

Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "             СБОРКА АВТОНОМНОГО РЕЛИЗА SPECHELPER" -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan

# 1. Поиск nuget.exe и восстановление пакетов
$nugetPath = Join-Path $rootDir ".tools\nuget.exe"
if (Test-Path $nugetPath) {
    Write-Host "`n[1/4] Восстановление NuGet пакетов..." -ForegroundColor Yellow
    & $nugetPath restore $solutionPath
} else {
    Write-Host "`n[1/4] nuget.exe не найден, используем имеющиеся пакеты..." -ForegroundColor Gray
}

# 2. Поиск MSBuild
$msbuildCandidates = @(
    "C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\MSBuild.exe",
    "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe",
    "C:\Program Files\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe",
    "C:\Program Files\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\MSBuild.exe",
    "C:\Program Files (x86)\Microsoft Visual Studio\2019\BuildTools\MSBuild\Current\Bin\MSBuild.exe",
    "C:\Program Files (x86)\Microsoft Visual Studio\2019\Community\MSBuild\Current\Bin\MSBuild.exe"
)

$msbuild = $msbuildCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $msbuild) {
    $vswhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
    if (Test-Path $vswhere) {
        $vsInstallPath = & $vswhere -latest -products * -requires Microsoft.Component.MSBuild -property installationPath
        if ($vsInstallPath) {
            $candidate = Join-Path $vsInstallPath "MSBuild\Current\Bin\MSBuild.exe"
            if (Test-Path $candidate) { $msbuild = $candidate }
        }
    }
}

if (-not $msbuild) {
    throw "MSBuild.exe не найден. Установите Visual Studio или Build Tools."
}

Write-Host "`n[2/4] Сборка проекта SpecHelper.sln (Release)..." -ForegroundColor Yellow
& $msbuild $solutionPath /p:Configuration=Release /t:Rebuild /v:minimal
if ($LASTEXITCODE -ne 0) {
    throw "Ошибка компиляции проекта SpecHelper."
}

# 3. Список обязательных библиотек (строго БЕЗ mscorlib.dll и System.*.dll)
$allowedDlls = @(
    "SpecHelper.dll",
    "ExcelDataReader.dll",
    "ExcelDataReader.DataSet.dll",
    "Newtonsoft.Json.dll",
    "OFFICE.DLL",
    "Microsoft.Office.Interop.Excel.dll",
    "extensibility.dll"
)

$releaseBin = Join-Path $rootDir "src\SpecHelper\bin\Release"
$libDir = Join-Path $rootDir "lib"
$packagesDir = Join-Path $rootDir "packages"

# 4. Формирование целевых автономных папок bin/
$targetBinDirs = @(
    (Join-Path $rootDir "bin"),
    (Join-Path $rootDir "SpecHelper_РабочийПК\bin")
)

Write-Host "`n[3/4] Формирование автономных папок bin/..." -ForegroundColor Yellow
foreach ($tb in $targetBinDirs) {
    if (-not (Test-Path $tb)) {
        New-Item -ItemType Directory -Path $tb -Force | Out-Null
    }

    # Очистка от запрещенных системных библиотек, которые ломают загрузку .NET
    Get-ChildItem -Path $tb -Filter "mscorlib.dll" -ErrorAction SilentlyContinue | Remove-Item -Force
    Get-ChildItem -Path $tb -Filter "System.*.dll" -ErrorAction SilentlyContinue | Remove-Item -Force

    # Копирование разрешенных библиотек
    foreach ($dll in $allowedDlls) {
        $source = $null
        if (Test-Path (Join-Path $releaseBin $dll)) {
            $source = Join-Path $releaseBin $dll
        } elseif (Test-Path (Join-Path $libDir $dll)) {
            $source = Join-Path $libDir $dll
        }

        if ($source) {
            Copy-Item -Path $source -Destination (Join-Path $tb $dll) -Force
            Write-Host "  -> скопирован: $dll" -ForegroundColor Gray
        } else {
            Write-Warning "Библиотека $dll не найдена ни в Release, ни в lib!"
        }
    }

    # Копирование отладочного символа pdb надстройки при наличии
    $pdb = Join-Path $releaseBin "SpecHelper.pdb"
    if (Test-Path $pdb) {
        Copy-Item -Path $pdb -Destination (Join-Path $tb "SpecHelper.pdb") -Force
    }

    Write-Host "Папка готова: $tb" -ForegroundColor Green
}

# Копирование баз данных в пакет SpecHelper_РабочийПК\databases
Write-Host "`n[4/5] Синхронизация каталогов баз данных..." -ForegroundColor Yellow
$srcDb = Join-Path $rootDir "databases"
$pkgDb = Join-Path $rootDir "SpecHelper_РабочийПК\databases"
if (Test-Path $srcDb) {
    if (-not (Test-Path $pkgDb)) { New-Item -ItemType Directory -Path $pkgDb -Force | Out-Null }
    Copy-Item -Path "$srcDb\*.xlsx" -Destination $pkgDb -Force
    Write-Host "Базы данных скопированы в $pkgDb" -ForegroundColor Green
}

# 5. Упаковка zip-архива для GitHub Releases
Write-Host "`n[5/5] Создание архива для GitHub Releases (SpecHelper_Update.zip)..." -ForegroundColor Yellow
Add-Type -AssemblyName System.IO.Compression.FileSystem
$zipOut = Join-Path $rootDir "SpecHelper_Update.zip"
if (Test-Path $zipOut) { Remove-Item $zipOut -Force }
[System.IO.Compression.ZipFile]::CreateFromDirectory((Join-Path $rootDir "bin"), $zipOut)
Write-Host "Архив готов: $zipOut" -ForegroundColor Green

Write-Host "`n================================================================" -ForegroundColor Cyan
Write-Host "СБОРКА УСПЕШНО ЗАВЕРШЕНА!" -ForegroundColor Green
Write-Host "Автономный пакет готов в: SpecHelper_РабочийПК" -ForegroundColor Green
Write-Host "Корневая папка bin готова: $rootDir\bin" -ForegroundColor Green
Write-Host "Архив для релиза готов: $zipOut" -ForegroundColor Green
Write-Host "================================================================" -ForegroundColor Cyan
