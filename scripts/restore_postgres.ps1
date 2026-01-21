# Restaurar backup de PostgreSQL (resamania) en Windows
# Uso: Edita las variables y ejecuta en PowerShell

$PgBin = "C:\Program Files\PostgreSQL\18\bin"
$Host = "localhost"
$Port = "5432"
$DbName = "resamania"
$DbUser = "resamania_user"
$BackupFile = "C:\AUTOMATIZACIONES\backups\resamania_YYYYMMDD_HHMMSS.dump"

# Opcion simple: escribir la password aqui (o usar .pgpass).
# $env:PGPASSWORD = "TU_PASSWORD_FUERTE"
$env:PGPASSWORD = ""

$pgRestore = Join-Path $PgBin "pg_restore.exe"

if (!(Test-Path $pgRestore)) {
    Write-Host "No se encuentra pg_restore en: $pgRestore"
    exit 1
}

if (!(Test-Path $BackupFile)) {
    Write-Host "No se encuentra el backup: $BackupFile"
    exit 1
}

Write-Host "Restaurando backup: $BackupFile"

& $pgRestore -h $Host -p $Port -U $DbUser -d $DbName --clean --if-exists $BackupFile

if ($LASTEXITCODE -ne 0) {
    Write-Host "Error en restauracion."
    exit 1
}

Write-Host "Restauracion completada."
