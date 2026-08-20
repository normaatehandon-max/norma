$ErrorActionPreference = "Stop"
Write-Host ""
Write-Host "=== COMPILANDO CEDULA_INGRESOS.exe AUTONOMO ===" -ForegroundColor Cyan
Write-Host ""

dotnet restore
dotnet publish -c Release -r win-x64 --self-contained true `
  -p:PublishSingleFile=true `
  -p:PublishTrimmed=false `
  -p:EnableCompressionInSingleFile=true

Write-Host ""
Write-Host "LISTO." -ForegroundColor Green
Write-Host "El ejecutable estara en:"
Write-Host ".\bin\Release\net8.0-windows\win-x64\publish\CEDULA_INGRESOS.exe"
