$ErrorActionPreference = "Stop"
$ManifestUrl = "https://vvihsesoftware.github.io/dc-expert/manifest.xml"
$LocalFolder = "C:\DCExpertManifest"
$ShareName = "DCExpert"

Write-Host "Установка DC Expert..." -ForegroundColor Cyan
if (!(Test-Path $LocalFolder)) { New-Item -ItemType Directory -Force -Path $LocalFolder | Out-Null }

Invoke-WebRequest -Uri $ManifestUrl -OutFile "$LocalFolder\manifest.xml"

if (!(Get-SmbShare -Name $ShareName -ErrorAction SilentlyContinue)) {
    New-SmbShare -Name $ShareName -Path $LocalFolder -ReadAccess "Everyone" | Out-Null
}

$RegistryPath = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{-DC-Expert-Trust-}"
if (!(Test-Path $RegistryPath)) { New-Item -Path $RegistryPath -Force | Out-Null }
New-ItemProperty -Path $RegistryPath -Name "Id" -Value "{-DC-Expert-Trust-}" -PropertyType String -Force | Out-Null
New-ItemProperty -Path $RegistryPath -Name "Url" -Value "\\$env:COMPUTERNAME\$ShareName" -PropertyType String -Force | Out-Null
New-ItemProperty -Path $RegistryPath -Name "Flags" -Value 1 -PropertyType DWord -Force | Out-Null

Write-Host "Готово! Откройте Excel -> Вставка -> Получить надстройки -> Общие папки (Shared Folder)." -ForegroundColor Green
Pause