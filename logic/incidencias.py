import os
import sqlite3
import time
from datetime import datetime

try:
    import psycopg
except Exception:
    psycopg = None


class IncidenciasDB:
    def __init__(self, db_path: str, db_config=None):
        self.db_path = db_path
        self.db_config = db_config or {}
        self.use_postgres = bool(self.db_config.get("host"))
        self.init_db()

    def _connect(self):
        if not self.use_postgres:
            return sqlite3.connect(self.db_path)
        if psycopg is None:
            raise RuntimeError("psycopg no esta instalado. Instala psycopg para usar PostgreSQL.")
        return psycopg.connect(
            host=self.db_config.get("host"),
            port=self.db_config.get("port"),
            dbname=self.db_config.get("name"),
            user=self.db_config.get("user"),
            password=self.db_config.get("password"),
            connect_timeout=5,
        )

    def _sql(self, sql: str) -> str:
        if self.use_postgres:
            return sql.replace("?", "%s")
        return sql

    def _lock_path(self):
        return f"{self.db_path}.lock"

    def _acquire_lock(self, stale_seconds=120, retries=25, wait_seconds=0.2):
        lock_path = self._lock_path()
        now = time.time()
        for _ in range(retries):
            try:
                fd = os.open(lock_path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
                with os.fdopen(fd, "w") as f:
                    f.write(f"{os.getpid()}|{now}\n")
                return True
            except FileExistsError:
                try:
                    mtime = os.path.getmtime(lock_path)
                    if now - mtime > stale_seconds:
                        os.remove(lock_path)
                        continue
                except Exception:
                    pass
                time.sleep(wait_seconds)
            except Exception:
                return False
        return False

    def _release_lock(self):
        try:
            os.remove(self._lock_path())
        except Exception:
            pass

    def _run_write(self, fn):
        if self.use_postgres:
            return fn()
        if not self._acquire_lock():
            raise RuntimeError("Base de datos ocupada. Intentalo de nuevo en unos segundos.")
        try:
            return fn()
        finally:
            self._release_lock()

    def init_db(self):
        os.makedirs(os.path.dirname(self.db_path), exist_ok=True)
        def _op():
            with self._connect() as conn:
                cur = conn.cursor()
                if self.use_postgres:
                    cur.execute(
                        """
                        CREATE TABLE IF NOT EXISTS inc_mapas (
                            id SERIAL PRIMARY KEY,
                            nombre TEXT,
                            ruta TEXT,
                            orden INTEGER,
                            ancho INTEGER,
                            alto INTEGER
                        )
                        """
                    )
                    cur.execute(
                        """
                        CREATE TABLE IF NOT EXISTS inc_areas (
                            id SERIAL PRIMARY KEY,
                            mapa_id INTEGER,
                            nombre TEXT,
                            x1 INTEGER, y1 INTEGER, x2 INTEGER, y2 INTEGER,
                            color TEXT,
                            FOREIGN KEY(mapa_id) REFERENCES inc_mapas(id)
                        )
                        """
                    )
                    cur.execute(
                        """
                        CREATE TABLE IF NOT EXISTS inc_maquinas (
                            id SERIAL PRIMARY KEY,
                            area_id INTEGER,
                            nombre TEXT,
                            serie TEXT,
                            numero_asignado TEXT,
                            x1 INTEGER, y1 INTEGER, x2 INTEGER, y2 INTEGER,
                            color TEXT,
                            FOREIGN KEY(area_id) REFERENCES inc_areas(id)
                        )
                        """
                    )
                    cur.execute(
                        """
                        CREATE TABLE IF NOT EXISTS inc_incidencias (
                            id SERIAL PRIMARY KEY,
                            mapa_id INTEGER,
                            area_id INTEGER,
                            maquina_id INTEGER,
                            fecha TIMESTAMP,
                            creador_nombre TEXT,
                            creador_apellido1 TEXT,
                            creador_apellido2 TEXT,
                            creador_movil TEXT,
                            creador_email TEXT,
                            elemento TEXT,
                            descripcion TEXT,
                            estado TEXT,
                            reporte_path TEXT,
                            FOREIGN KEY(mapa_id) REFERENCES inc_mapas(id),
                            FOREIGN KEY(area_id) REFERENCES inc_areas(id),
                            FOREIGN KEY(maquina_id) REFERENCES inc_maquinas(id)
                        )
                        """
                    )
                    cur.execute("ALTER TABLE inc_incidencias ADD COLUMN IF NOT EXISTS reporte_path TEXT")
                    cur.execute("ALTER TABLE inc_incidencias ADD COLUMN IF NOT EXISTS creador_nombre TEXT")
                    cur.execute("ALTER TABLE inc_incidencias ADD COLUMN IF NOT EXISTS creador_apellido1 TEXT")
                    cur.execute("ALTER TABLE inc_incidencias ADD COLUMN IF NOT EXISTS creador_apellido2 TEXT")
                    cur.execute("ALTER TABLE inc_incidencias ADD COLUMN IF NOT EXISTS creador_movil TEXT")
                    cur.execute("ALTER TABLE inc_incidencias ADD COLUMN IF NOT EXISTS creador_email TEXT")
                else:
                    cur.execute(
                        """
                        CREATE TABLE IF NOT EXISTS inc_mapas (
                            id INTEGER PRIMARY KEY AUTOINCREMENT,
                            nombre TEXT,
                            ruta TEXT,
                            orden INTEGER,
                            ancho INTEGER,
                            alto INTEGER
                        )
                        """
                    )
                    cur.execute(
                        """
                        CREATE TABLE IF NOT EXISTS inc_areas (
                            id INTEGER PRIMARY KEY AUTOINCREMENT,
                            mapa_id INTEGER,
                            nombre TEXT,
                            x1 INTEGER, y1 INTEGER, x2 INTEGER, y2 INTEGER,
                            color TEXT,
                            FOREIGN KEY(mapa_id) REFERENCES inc_mapas(id)
                        )
                        """
                    )
                    cur.execute(
                        """
                        CREATE TABLE IF NOT EXISTS inc_maquinas (
                            id INTEGER PRIMARY KEY AUTOINCREMENT,
                            area_id INTEGER,
                            nombre TEXT,
                            serie TEXT,
                            numero_asignado TEXT,
                            x1 INTEGER, y1 INTEGER, x2 INTEGER, y2 INTEGER,
                            color TEXT,
                            FOREIGN KEY(area_id) REFERENCES inc_areas(id)
                        )
                        """
                    )
                    cur.execute(
                        """
                        CREATE TABLE IF NOT EXISTS inc_incidencias (
                            id INTEGER PRIMARY KEY AUTOINCREMENT,
                            mapa_id INTEGER,
                            area_id INTEGER,
                            maquina_id INTEGER,
                            fecha TEXT,
                            creador_nombre TEXT,
                            creador_apellido1 TEXT,
                            creador_apellido2 TEXT,
                            creador_movil TEXT,
                            creador_email TEXT,
                            elemento TEXT,
                            descripcion TEXT,
                            estado TEXT,
                            reporte_path TEXT,
                            FOREIGN KEY(mapa_id) REFERENCES inc_mapas(id),
                            FOREIGN KEY(area_id) REFERENCES inc_areas(id),
                            FOREIGN KEY(maquina_id) REFERENCES inc_maquinas(id)
                        )
                        """
                    )
                    cur.execute("PRAGMA table_info(inc_incidencias)")
                    cols = [row[1] for row in cur.fetchall()]
                    if "reporte_path" not in cols:
                        cur.execute("ALTER TABLE inc_incidencias ADD COLUMN reporte_path TEXT")
                    if "creador_nombre" not in cols:
                        cur.execute("ALTER TABLE inc_incidencias ADD COLUMN creador_nombre TEXT")
                    if "creador_apellido1" not in cols:
                        cur.execute("ALTER TABLE inc_incidencias ADD COLUMN creador_apellido1 TEXT")
                    if "creador_apellido2" not in cols:
                        cur.execute("ALTER TABLE inc_incidencias ADD COLUMN creador_apellido2 TEXT")
                    if "creador_movil" not in cols:
                        cur.execute("ALTER TABLE inc_incidencias ADD COLUMN creador_movil TEXT")
                    if "creador_email" not in cols:
                        cur.execute("ALTER TABLE inc_incidencias ADD COLUMN creador_email TEXT")
                conn.commit()
        self._run_write(_op)

    def add_map(self, nombre, ruta, ancho, alto):
        def _op():
            with self._connect() as conn:
                cur = conn.cursor()
                cur.execute("SELECT COALESCE(MAX(orden), 0) + 1 FROM inc_mapas")
                orden = cur.fetchone()[0]
                cur.execute(
                    self._sql("INSERT INTO inc_mapas (nombre, ruta, orden, ancho, alto) VALUES (?, ?, ?, ?, ?)"),
                    (nombre, ruta, orden, ancho, alto),
                )
                conn.commit()
        self._run_write(_op)

    def list_maps(self):
        with self._connect() as conn:
            cur = conn.cursor()
            cur.execute("SELECT id, nombre, ruta, orden, ancho, alto FROM inc_mapas ORDER BY orden ASC, id ASC")
            return cur.fetchall()

    def delete_map(self, mapa_id):
        def _op():
            with self._connect() as conn:
                cur = conn.cursor()
                # borrar incidencias
                cur.execute(self._sql("DELETE FROM inc_incidencias WHERE mapa_id=?"), (mapa_id,))
                # borrar maquinas y areas
                cur.execute(
                    self._sql("DELETE FROM inc_maquinas WHERE area_id IN (SELECT id FROM inc_areas WHERE mapa_id=?)"),
                    (mapa_id,),
                )
                cur.execute(self._sql("DELETE FROM inc_areas WHERE mapa_id=?"), (mapa_id,))
                # borrar mapa
                cur.execute(self._sql("DELETE FROM inc_mapas WHERE id=?"), (mapa_id,))
                conn.commit()
        self._run_write(_op)

    def update_map_path(self, mapa_id, ruta):
        def _op():
            with self._connect() as conn:
                cur = conn.cursor()
                cur.execute(self._sql("UPDATE inc_mapas SET ruta=? WHERE id=?"), (ruta, mapa_id))
                conn.commit()
        self._run_write(_op)

    def add_area(self, mapa_id, nombre, x1, y1, x2, y2, color):
        def _op():
            with self._connect() as conn:
                cur = conn.cursor()
                cur.execute(
                    self._sql("INSERT INTO inc_areas (mapa_id, nombre, x1, y1, x2, y2, color) VALUES (?, ?, ?, ?, ?, ?, ?)"),
                    (mapa_id, nombre, x1, y1, x2, y2, color),
                )
                conn.commit()
        self._run_write(_op)

    def list_areas(self, mapa_id):
        with self._connect() as conn:
            cur = conn.cursor()
            cur.execute(
                self._sql("SELECT id, nombre, x1, y1, x2, y2, color FROM inc_areas WHERE mapa_id=?"),
                (mapa_id,),
            )
            return cur.fetchall()

    def update_area(self, area_id, nombre, x1, y1, x2, y2):
        def _op():
            with self._connect() as conn:
                cur = conn.cursor()
                cur.execute(
                    self._sql("UPDATE inc_areas SET nombre=?, x1=?, y1=?, x2=?, y2=? WHERE id=?"),
                    (nombre, x1, y1, x2, y2, area_id),
                )
                conn.commit()
        self._run_write(_op)

    def add_machine(self, area_id, nombre, serie, numero_asignado, x1, y1, x2, y2, color):
        def _op():
            with self._connect() as conn:
                cur = conn.cursor()
                cur.execute(
                    self._sql(
                        """
                        INSERT INTO inc_maquinas (area_id, nombre, serie, numero_asignado, x1, y1, x2, y2, color)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """
                    ),
                    (area_id, nombre, serie, numero_asignado, x1, y1, x2, y2, color),
                )
                conn.commit()
        self._run_write(_op)

    def list_machines(self, mapa_id):
        with self._connect() as conn:
            cur = conn.cursor()
            cur.execute(
                self._sql(
                    """
                    SELECT m.id, m.area_id, m.nombre, m.serie, m.numero_asignado, m.x1, m.y1, m.x2, m.y2, m.color,
                           a.nombre
                    FROM inc_maquinas m
                    JOIN inc_areas a ON a.id = m.area_id
                    WHERE a.mapa_id=?
                    """
                ),
                (mapa_id,),
            )
            return cur.fetchall()

    def update_machine(self, machine_id, nombre, serie, numero_asignado, x1, y1, x2, y2):
        def _op():
            with self._connect() as conn:
                cur = conn.cursor()
                cur.execute(
                    self._sql(
                        """
                        UPDATE inc_maquinas
                        SET nombre=?, serie=?, numero_asignado=?, x1=?, y1=?, x2=?, y2=?
                        WHERE id=?
                        """
                    ),
                    (nombre, serie, numero_asignado, x1, y1, x2, y2, machine_id),
                )
                conn.commit()
        self._run_write(_op)

    def delete_machine(self, machine_id):
        def _op():
            with self._connect() as conn:
                cur = conn.cursor()
                cur.execute(self._sql("DELETE FROM inc_incidencias WHERE maquina_id=?"), (machine_id,))
                cur.execute(self._sql("DELETE FROM inc_maquinas WHERE id=?"), (machine_id,))
                conn.commit()
        self._run_write(_op)

    def add_incident(
        self,
        mapa_id,
        area_id,
        maquina_id,
        elemento,
        descripcion,
        estado="PENDIENTE",
        reporte_path="",
        creador_nombre="",
        creador_apellido1="",
        creador_apellido2="",
        creador_movil="",
        creador_email="",
    ):
        def _op():
            with self._connect() as conn:
                cur = conn.cursor()
                cur.execute(
                    self._sql(
                        """
                        INSERT INTO inc_incidencias (
                            mapa_id, area_id, maquina_id, fecha,
                            creador_nombre, creador_apellido1, creador_apellido2, creador_movil, creador_email,
                            elemento, descripcion, estado, reporte_path
                        )
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """
                    ),
                    (
                        mapa_id,
                        area_id,
                        maquina_id,
                        datetime.now().strftime("%Y-%m-%d %H:%M"),
                        creador_nombre,
                        creador_apellido1,
                        creador_apellido2,
                        creador_movil,
                        creador_email,
                        elemento,
                        descripcion,
                        estado,
                        reporte_path,
                    ),
                )
                conn.commit()
        self._run_write(_op)

    def list_incidencias(self, mapa_id):
        with self._connect() as conn:
            cur = conn.cursor()
            cur.execute(
                self._sql(
                    """
                    SELECT i.id, i.fecha, i.estado, i.elemento, i.descripcion, i.reporte_path,
                           i.creador_nombre, i.creador_apellido1, i.creador_apellido2, i.creador_movil, i.creador_email,
                           a.nombre, m.nombre, m.serie, m.numero_asignado
                    FROM inc_incidencias i
                    LEFT JOIN inc_areas a ON a.id = i.area_id
                    LEFT JOIN inc_maquinas m ON m.id = i.maquina_id
                    WHERE i.mapa_id=?
                    ORDER BY i.fecha DESC
                    """
                ),
                (mapa_id,),
            )
            return cur.fetchall()

    def update_incidencia_estado(self, inc_id, estado):
        def _op():
            with self._connect() as conn:
                cur = conn.cursor()
                cur.execute(self._sql("UPDATE inc_incidencias SET estado=? WHERE id=?"), (estado, inc_id))
                conn.commit()
        self._run_write(_op)


    def update_incidencia_reporte(self, inc_id, reporte_path):
        def _op():
            with self._connect() as conn:
                cur = conn.cursor()
                cur.execute(self._sql("UPDATE inc_incidencias SET reporte_path=? WHERE id=?"), (reporte_path, inc_id))
                conn.commit()
        self._run_write(_op)

    def update_incidencia(self, inc_id, elemento, descripcion, estado):
        def _op():
            with self._connect() as conn:
                cur = conn.cursor()
                cur.execute(
                    self._sql("UPDATE inc_incidencias SET elemento=?, descripcion=?, estado=? WHERE id=?"),
                    (elemento, descripcion, estado, inc_id),
                )
                conn.commit()
        self._run_write(_op)

    def delete_incidencia(self, inc_id):
        def _op():
            with self._connect() as conn:
                cur = conn.cursor()
                cur.execute(self._sql("DELETE FROM inc_incidencias WHERE id=?"), (inc_id,))
                conn.commit()
        self._run_write(_op)
