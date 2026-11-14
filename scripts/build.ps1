# Скрипт сборки проекта SpecHelper
$ErrorActionPreference = "Stop"

$msbuild = "C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\MSBuild.exe"
if (-not (Test-Path $msbuild)) {
    throw "MSBuild не найден по пути $msbuild"
}

$solutionPath = "D:\EXCEL mibil\src\SpecHelper.sln"

Write-Host "Сборка SpecHelper.sln (Release)..." -ForegroundColor Cyan
& $msbuild $solutionPath /p:Configuration=Release /v:minimal

if ($LASTEXITCODE -eq 0) {
    Write-Host "Сборка успешно завершена!" -ForegroundColor Green
} else {
    Write-Host "Ошибка при сборке проекта." -ForegroundColor Red
}
