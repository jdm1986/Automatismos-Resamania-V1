# Backups PostgreSQL (PC Club)

Objetivo: si se rompe el PC de recepcion, restaurar la base de datos en un PC nuevo en 10-15 minutos.

## 1) Preparar carpeta de backups
Crea la carpeta:
```
C:\AUTOMATIZACIONES\backups
```

## 2) Configurar el script de backup
Archivo:
```
scripts/backup_postgres.ps1
```

Editar variables:
- $PgBin: ruta a PostgreSQL (ej. C:\Program Files\PostgreSQL\18\bin)
- $Host: localhost
- $Port: 5432
- $DbName: resamania
- $DbUser: resamania_user
- $BackupDir: C:\AUTOMATIZACIONES\backups
- $OneDriveDir: ruta de OneDrive para copiar el backup (opcional). Ejemplo:
  C:\Users\FP Villalobos\VILLALOBOS CLUB\OneDrive - UpGyms Iberia\BACKUP AUTOMATISMOS
- $RetentionDays: 14
- $env:PGPASSWORD: password del usuario (opcional)

## 3) Probar backup manual
PowerShell (como admin o usuario normal):
```
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
.\scripts\backup_postgres.ps1
```
Debe crear un archivo .dump en C:\AUTOMATIZACIONES\backups

## 4) Crear tarea programada (diaria)
1. Abrir "Programador de tareas".
2. Crear tarea basica.
3. Nombre: Backup Resamania.
4. Trigger: Diario (ej. 23:30).
5. Accion: Iniciar un programa.
6. Programa o script:
```
powershell.exe
```
7. Agregar argumentos:
```
-ExecutionPolicy Bypass -File "C:\RUTA\A\TU\PROYECTO\scripts\backup_postgres.ps1"
```
8. Finalizar.

## 5) Restaurar en un PC nuevo
1. Instalar PostgreSQL.
2. Crear usuario y BD (ver INSTALACION_OTRO_CLUB.md).
3. Copiar el backup .dump al nuevo PC.
4. Editar:
```
scripts/restore_postgres.ps1
```
5. Ejecutar:
```
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
.\scripts\restore_postgres.ps1
```

## 6) Recomendaciones
- Copia la carpeta C:\AUTOMATIZACIONES\backups a un USB o NAS.
- No borres el ultimo backup valido.
