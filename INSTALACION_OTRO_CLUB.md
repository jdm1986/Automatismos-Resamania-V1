# Manual de instalacion - Otro Club (PostgreSQL + 2 PCs)

Escenario: intranet local, PostgreSQL en PC Club, PC Manager como cliente. Dos PCs trabajando a la vez.

## 0) Requisitos previos
- Ambos PCs en la misma red local (LAN).
- IP fija para el PC Club (recomendado).
- Ejecutable en carpeta local (no OneDrive).

## A) PC CLUB (Servidor)

### 1) Asignar IP fija (recomendado)
1. Panel de control > Centro de redes > Cambiar configuracion del adaptador.
2. Boton derecho en Ethernet > Propiedades.
3. Protocolo de Internet version 4 (TCP/IPv4) > Propiedades.
4. Asignar IP fija:
   - IP: 192.168.X.Y (ej. 192.168.10.12)
   - Mascara: 255.255.255.0
   - Puerta de enlace: 192.168.X.1
5. DNS: el del router o 8.8.8.8.

### 2) Instalar PostgreSQL
1. Instalar PostgreSQL 18 (o 16/17).
2. Puerto: 5432.
3. Contrasena del superusuario: fuerte y guardada.

### 3) Crear usuario y base de datos
Abrir SQL Shell (psql) y entrar con el usuario postgres.

Comandos:
```
CREATE USER resamania_user WITH PASSWORD 'TU_PASSWORD_FUERTE';
CREATE DATABASE resamania OWNER resamania_user;
GRANT ALL PRIVILEGES ON DATABASE resamania TO resamania_user;
\c resamania
GRANT ALL ON SCHEMA public TO resamania_user;
```

### 4) Permitir conexiones LAN
Archivo: postgresql.conf
Ruta tipica: C:\Program Files\PostgreSQL\18\data\postgresql.conf

Cambiar:
```
listen_addresses = '*'
```

Archivo: pg_hba.conf
Ruta tipica: C:\Program Files\PostgreSQL\18\data\pg_hba.conf

Anadir al final:
```
host    resamania    resamania_user    192.168.X.0/24    md5
```

### 5) Reiniciar servicio PostgreSQL
Abrir services.msc y reiniciar PostgreSQL 18.

### 6) Abrir firewall solo en red privada
1. Windows Defender Firewall > Reglas de entrada > Nueva regla.
2. Puerto TCP 5432.
3. Solo red privada.
4. Nombre: PostgreSQL 5432 LAN.

### 7) Verificar que escucha
PowerShell:
```
netstat -ano | findstr 5432
```
Debe salir LISTENING.

## B) PC MANAGER (Cliente)

### 1) Probar conexion
PowerShell:
```
Test-NetConnection 192.168.X.Y -Port 5432
```
Debe salir TcpTestSucceeded : True.

### 2) Copiar el ejecutable a local
Ejemplo:
```
C:\AUTOMATIZACIONES\AUTOMATISMOS_RESAMANIA.exe
```

## C) Configurar el programa (ambos PCs)
1. Abrir el programa.
2. Boton CONFIG BD:
   - Host: IP del PC Club (ej. 192.168.10.12)
   - Puerto: 5432
   - BD: resamania
   - Usuario: resamania_user
   - Password: la creada en PostgreSQL
3. Guardar.

## D) Cargar CSV (solo PC Club)
1. Si no hay CSV en BD, la app pedira la carpeta.
2. Seleccionar carpeta con los 4 CSV:
   - RESUMEN CLIENTE.csv
   - ACCESOS.csv
   - FACTURAS Y VALES.csv
   - IMPAGOS.csv
3. Pulsa RECARGAR BD para subirlos a PostgreSQL.

## E) Mapas e incidencias
1. Cargar mapa desde el PC Club.
2. El PC Manager vera todo desde la BD.

## F) Flujo diario recomendado
1. PC Club actualiza CSV diariamente.
2. Pulsa RECARGAR BD.
3. PC Manager solo abre y trabaja.

## G) Problemas tipicos
- No conecta: revisar firewall, pg_hba.conf y usar IP.
- No carga datos: verificar que los CSV se subieron desde PC Club.
- No aparece mapa: cargar primero en PC Club.

## H) Backups recomendados
- Backup diario local: C:\AUTOMATIZACIONES\backups
- Copia automatica a OneDrive (ejemplo):
  C:\Users\FP Villalobos\VILLALOBOS CLUB\OneDrive - UpGyms Iberia\BACKUP AUTOMATISMOS
- Ver scripts: scripts/backup_postgres.ps1 y scripts/restore_postgres.ps1
