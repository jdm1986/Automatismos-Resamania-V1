import os
import sqlite3
from datetime import datetime


class IncidenciasDB:
    def __init__(self, db_path: str):
        self.db_path = db_path
        self.init_db()

    def _connect(self):
        return sqlite3.connect(self.db_path)

    def init_db(self):
        os.makedirs(os.path.dirname(self.db_path), exist_ok=True)
        with self._connect() as conn:
            cur = conn.cursor()
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
            conn.commit()

    def add_map(self, nombre, ruta, ancho, alto):
        with self._connect() as conn:
            cur = conn.cursor()
            cur.execute("SELECT COALESCE(MAX(orden), 0) + 1 FROM inc_mapas")
            orden = cur.fetchone()[0]
            cur.execute(
                "INSERT INTO inc_mapas (nombre, ruta, orden, ancho, alto) VALUES (?, ?, ?, ?, ?)",
                (nombre, ruta, orden, ancho, alto),
            )
            conn.commit()

    def list_maps(self):
        with self._connect() as conn:
            cur = conn.cursor()
            cur.execute("SELECT id, nombre, ruta, orden, ancho, alto FROM inc_mapas ORDER BY orden ASC, id ASC")
            return cur.fetchall()

    def delete_map(self, mapa_id):
        with self._connect() as conn:
            cur = conn.cursor()
            # borrar incidencias
            cur.execute("DELETE FROM inc_incidencias WHERE mapa_id=?", (mapa_id,))
            # borrar maquinas y areas
            cur.execute(
                "DELETE FROM inc_maquinas WHERE area_id IN (SELECT id FROM inc_areas WHERE mapa_id=?)",
                (mapa_id,),
            )
            cur.execute("DELETE FROM inc_areas WHERE mapa_id=?", (mapa_id,))
            # borrar mapa
            cur.execute("DELETE FROM inc_mapas WHERE id=?", (mapa_id,))
            conn.commit()

    def add_area(self, mapa_id, nombre, x1, y1, x2, y2, color):
        with self._connect() as conn:
            cur = conn.cursor()
            cur.execute(
                "INSERT INTO inc_areas (mapa_id, nombre, x1, y1, x2, y2, color) VALUES (?, ?, ?, ?, ?, ?, ?)",
                (mapa_id, nombre, x1, y1, x2, y2, color),
            )
            conn.commit()

    def list_areas(self, mapa_id):
        with self._connect() as conn:
            cur = conn.cursor()
            cur.execute(
                "SELECT id, nombre, x1, y1, x2, y2, color FROM inc_areas WHERE mapa_id=?",
                (mapa_id,),
            )
            return cur.fetchall()

    def update_area(self, area_id, nombre, x1, y1, x2, y2):
        with self._connect() as conn:
            cur = conn.cursor()
            cur.execute(
                "UPDATE inc_areas SET nombre=?, x1=?, y1=?, x2=?, y2=? WHERE id=?",
                (nombre, x1, y1, x2, y2, area_id),
            )
            conn.commit()

    def add_machine(self, area_id, nombre, serie, numero_asignado, x1, y1, x2, y2, color):
        with self._connect() as conn:
            cur = conn.cursor()
            cur.execute(
                """
                INSERT INTO inc_maquinas (area_id, nombre, serie, numero_asignado, x1, y1, x2, y2, color)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (area_id, nombre, serie, numero_asignado, x1, y1, x2, y2, color),
            )
            conn.commit()

    def list_machines(self, mapa_id):
        with self._connect() as conn:
            cur = conn.cursor()
            cur.execute(
                """
                SELECT m.id, m.area_id, m.nombre, m.serie, m.numero_asignado, m.x1, m.y1, m.x2, m.y2, m.color,
                       a.nombre
                FROM inc_maquinas m
                JOIN inc_areas a ON a.id = m.area_id
                WHERE a.mapa_id=?
                """,
                (mapa_id,),
            )
            return cur.fetchall()

    def update_machine(self, machine_id, nombre, serie, numero_asignado, x1, y1, x2, y2):
        with self._connect() as conn:
            cur = conn.cursor()
            cur.execute(
                """
                UPDATE inc_maquinas
                SET nombre=?, serie=?, numero_asignado=?, x1=?, y1=?, x2=?, y2=?
                WHERE id=?
                """,
                (nombre, serie, numero_asignado, x1, y1, x2, y2, machine_id),
            )
            conn.commit()

    def delete_machine(self, machine_id):
        with self._connect() as conn:
            cur = conn.cursor()
            cur.execute("DELETE FROM inc_incidencias WHERE maquina_id=?", (machine_id,))
            cur.execute("DELETE FROM inc_maquinas WHERE id=?", (machine_id,))
            conn.commit()

    def add_incident(self, mapa_id, area_id, maquina_id, elemento, descripcion, estado="PENDIENTE", reporte_path=""):
        with self._connect() as conn:
            cur = conn.cursor()
            cur.execute(
                """
                INSERT INTO inc_incidencias (mapa_id, area_id, maquina_id, fecha, elemento, descripcion, estado, reporte_path)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    mapa_id,
                    area_id,
                    maquina_id,
                    datetime.now().strftime("%Y-%m-%d %H:%M"),
                    elemento,
                    descripcion,
                    estado,
                    reporte_path,
                ),
            )
            conn.commit()

    def list_incidencias(self, mapa_id):
        with self._connect() as conn:
            cur = conn.cursor()
            cur.execute(
                """
                SELECT i.id, i.fecha, i.estado, i.elemento, i.descripcion, i.reporte_path,
                       a.nombre, m.nombre, m.serie, m.numero_asignado
                FROM inc_incidencias i
                LEFT JOIN inc_areas a ON a.id = i.area_id
                LEFT JOIN inc_maquinas m ON m.id = i.maquina_id
                WHERE i.mapa_id=?
                ORDER BY i.fecha DESC
                """,
                (mapa_id,),
            )
            return cur.fetchall()

    def update_incidencia_estado(self, inc_id, estado):
        with self._connect() as conn:
            cur = conn.cursor()
            cur.execute("UPDATE inc_incidencias SET estado=? WHERE id=?", (estado, inc_id))
            conn.commit()

    def update_incidencia(self, inc_id, elemento, descripcion, estado):
        with self._connect() as conn:
            cur = conn.cursor()
            cur.execute(
                "UPDATE inc_incidencias SET elemento=?, descripcion=?, estado=? WHERE id=?",
                (elemento, descripcion, estado, inc_id),
            )
            conn.commit()

    def delete_incidencia(self, inc_id):
        with self._connect() as conn:
            cur = conn.cursor()
            cur.execute("DELETE FROM inc_incidencias WHERE id=?", (inc_id,))
            conn.commit()
