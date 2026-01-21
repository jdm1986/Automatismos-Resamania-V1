# Backup diario de PostgreSQL (resamania) en Windows
# Uso: boton derecho > "Run with PowerShell"

$PgBin = "C:\Program Files\PostgreSQL\18\bin"
$PgHost = "localhost"
$Port = "5432"
$DbName = "resamania"
$DbUser = "resamania_user"
$BackupDir = "C:\AUTOMATIZACIONES\backups"
$OneDriveDir = "C:\Users\FP Villalobos\VILLALOBOS CLUB\OneDrive - UpGyms Iberia\BACKUP AUTOMATISMOS"
$RetentionDays = 14

if (!(Test-Path $BackupDir)) {
    New-Item -ItemType Directory -Path $BackupDir | Out-Null
}
if ($OneDriveDir -and !(Test-Path $OneDriveDir)) {
    New-Item -ItemType Directory -Path $OneDriveDir | Out-Null
}

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$backupFile = Join-Path $BackupDir "$DbName`_$timestamp.dump"
$logFile = Join-Path $BackupDir "backup_log.txt"

# Opcion simple: escribir la password aqui (o usar .pgpass).
# $env:PGPASSWORD = "TU_PASSWORD_FUERTE"
$env:PGPASSWORD = ""

$pgDump = Join-Path $PgBin "pg_dump.exe"

if (!(Test-Path $pgDump)) {
    Write-Host "No se encuentra pg_dump en: $pgDump"
    exit 1
}

Write-Host "Iniciando backup a: $backupFile"

& $pgDump -h $PgHost -p $Port -U $DbUser -F c -b -f $backupFile $DbName 2>> $logFile

if ($LASTEXITCODE -ne 0) {
    Write-Host "Error en backup. Revisa $logFile"
    exit 1
}

# Copiar a OneDrive si esta disponible
if ($OneDriveDir -and (Test-Path $OneDriveDir)) {
    $destFile = Join-Path $OneDriveDir (Split-Path $backupFile -Leaf)
    Copy-Item -Path $backupFile -Destination $destFile -Force
}

# Limpieza de backups antiguos
Get-ChildItem $BackupDir -Filter "$DbName`_*.dump" | Where-Object {
    $_.LastWriteTime -lt (Get-Date).AddDays(-$RetentionDays)
} | Remove-Item -Force

if ($OneDriveDir -and (Test-Path $OneDriveDir)) {
    Get-ChildItem $OneDriveDir -Filter "$DbName`_*.dump" | Where-Object {
        $_.LastWriteTime -lt (Get-Date).AddDays(-$RetentionDays)
    } | Remove-Item -Force
}

Write-Host "Backup completado."
