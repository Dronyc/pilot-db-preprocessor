<#
.SYNOPSIS
Создаёт резервную копию файла Препроцессор_структуры_БД_Pilot.xlsm.
Хранит последние 10 копий.
#>
param(
    [string]$SourceFile = "Препроцессор_структуры_БД_Pilot.xlsm",
    [string]$BackupDir = "_Backup",
    [int]$MaxBackups = 10
)

$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$fullSourcePath = Join-Path $scriptDir $SourceFile

if (-not (Test-Path $fullSourcePath)) {
    Write-Host "ОШИБКА: Файл не найден: $fullSourcePath" -ForegroundColor Red
    Read-Host "Нажмите Enter для выхода"
    exit 1
}

$fullBackupDir = Join-Path $scriptDir $BackupDir
if (-not (Test-Path $fullBackupDir)) {
    New-Item -ItemType Directory -Path $fullBackupDir | Out-Null
}

$timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$backupFile = "Препроцессор_${timestamp}.xlsm"
$fullBackupPath = Join-Path $fullBackupDir $backupFile

Copy-Item -Path $fullSourcePath -Destination $fullBackupPath
Write-Host "Бэкап создан: $backupFile" -ForegroundColor Green

# Лог
$logFile = Join-Path $fullBackupDir "backup.log"
$logEntry = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Создан бэкап: $backupFile ($([math]::Round((Get-Item $fullBackupPath).Length / 1MB, 2)) MB)"
Add-Content -Path $logFile -Value $logEntry

# Ротация (удаление старых)
$backups = Get-ChildItem -Path $fullBackupDir -Filter "Препроцессор_*.xlsm" | Sort-Object LastWriteTime -Descending
if ($backups.Count -gt $MaxBackups) {
    $toDelete = $backups | Select-Object -Skip $MaxBackups
    foreach ($file in $toDelete) {
        Remove-Item $file.FullName -Force
        Add-Content -Path $logFile -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Удалён старый бэкап: $($file.Name)"
        Write-Host "Удалён старый: $($file.Name)" -ForegroundColor Yellow
    }
}

Write-Host "Готово. Сохранено копий: $($backups.Count)"
Read-Host "Нажмите Enter для выхода"