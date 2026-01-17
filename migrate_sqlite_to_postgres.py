#!/usr/bin/env python
import argparse
import json
import os
import sqlite3
import sys

try:
    import psycopg
except Exception:
    psycopg = None

from logic.impagos import ImpagosDB
from logic.incidencias import IncidenciasDB


def load_config(path):
    if not os.path.exists(path):
        return {}
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def get_db_config(cfg):
    db = cfg.get("db", {}) if isinstance(cfg, dict) else {}
    return {
        "host": str(db.get("host", "")).strip(),
        "port": str(db.get("port", "5432")).strip(),
        "name": str(db.get("name", "resamania")).strip(),
        "user": str(db.get("user", "resamania")).strip(),
        "password": str(db.get("password", "")),
    }


def get_data_dir(cfg, config_path, override=None):
    if override:
        return os.path.normpath(override)
    data_dir = cfg.get("data_dir") if isinstance(cfg, dict) else None
    if data_dir:
        return os.path.normpath(data_dir)
    carpeta = cfg.get("carpeta_datos") if isinstance(cfg, dict) else None
    if carpeta:
        return os.path.normpath(os.path.join(carpeta, "data"))
    return os.path.normpath(os.path.join(os.path.dirname(config_path), "data"))


def table_exists(conn, name):
    cur = conn.cursor()
    cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name=?", (name,))
    return cur.fetchone() is not None


def fetch_rows(conn, table, columns):
    if not table_exists(conn, table):
        return []
    cur = conn.cursor()
    cols = ", ".join(columns)
    cur.execute(f"SELECT {cols} FROM {table}")
    return cur.fetchall()


def insert_rows(conn, table, columns, rows):
    if not rows:
        return 0
    cols = ", ".join(columns)
    placeholders = ", ".join(["%s"] * len(columns))
    on_conflict = ""
    if "id" in columns:
        on_conflict = " ON CONFLICT (id) DO NOTHING"
    elif table == "impagos_meta":
        on_conflict = " ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value"
    sql = f"INSERT INTO {table} ({cols}) VALUES ({placeholders}){on_conflict}"
    with conn.cursor() as cur:
        cur.executemany(sql, rows)
    return len(rows)


def set_sequence(conn, table):
    with conn.cursor() as cur:
        cur.execute(
            f"SELECT setval(pg_get_serial_sequence(%s, 'id'), COALESCE(MAX(id), 1), true) FROM {table}",
            (table,),
        )


def main():
    parser = argparse.ArgumentParser(description="Migrar datos de SQLite a PostgreSQL")
    parser.add_argument("--config", default="config.json", help="Ruta a config.json")
    parser.add_argument("--data-dir", help="Ruta a carpeta data (si no esta en config)")
    parser.add_argument("--reset", action="store_true", help="Vaciar tablas en PostgreSQL antes de migrar")
    args = parser.parse_args()

    if psycopg is None:
        print("ERROR: psycopg no esta instalado. Ejecuta: pip install \"psycopg[binary]\"")
        return 2

    cfg = load_config(args.config)
    db_cfg = get_db_config(cfg)
    if not db_cfg.get("host"):
        print("ERROR: No hay DB configurada en config.json (campo db.host).")
        return 2

    data_dir = get_data_dir(cfg, args.config, args.data_dir)
    impagos_sqlite = os.path.join(data_dir, "impagos.db")
    incidencias_sqlite = os.path.join(data_dir, "incidencias.db")

    if not os.path.exists(impagos_sqlite) and not os.path.exists(incidencias_sqlite):
        print(f"ERROR: No se encontraron SQLite en {data_dir}")
        return 2

    # Asegurar esquema en Postgres
    ImpagosDB(impagos_sqlite, db_config=db_cfg)
    IncidenciasDB(incidencias_sqlite, db_config=db_cfg)

    with psycopg.connect(
        host=db_cfg["host"],
        port=db_cfg["port"],
        dbname=db_cfg["name"],
        user=db_cfg["user"],
        password=db_cfg["password"],
    ) as conn_pg:
        conn_pg.autocommit = False
        with conn_pg.cursor() as cur:
            if args.reset:
                cur.execute(
                    "TRUNCATE TABLE inc_incidencias, inc_maquinas, inc_areas, inc_mapas RESTART IDENTITY CASCADE"
                )
                cur.execute(
                    "TRUNCATE TABLE impagos_gestion, impagos_eventos, impagos_clientes, impagos_meta RESTART IDENTITY CASCADE"
                )

        # Migrar impagos
        if os.path.exists(impagos_sqlite):
            conn_sqlite = sqlite3.connect(impagos_sqlite)
            try:
                tables = [
                    ("impagos_clientes", ["id", "numero_cliente", "nombre", "apellidos", "email", "movil"]),
                    ("impagos_eventos", ["id", "cliente_id", "fecha_export", "incidentes"]),
                    ("impagos_gestion", ["id", "cliente_id", "fecha", "accion", "plantilla", "staff", "notas"]),
                    ("impagos_meta", ["key", "value"]),
                ]
                for table, cols in tables:
                    rows = fetch_rows(conn_sqlite, table, cols)
                    inserted = insert_rows(conn_pg, table, cols, rows)
                    print(f"{table}: {inserted} filas")
                for table, cols in tables:
                    if "id" in cols:
                        set_sequence(conn_pg, table)
            finally:
                conn_sqlite.close()

        # Migrar incidencias
        if os.path.exists(incidencias_sqlite):
            conn_sqlite = sqlite3.connect(incidencias_sqlite)
            try:
                tables = [
                    ("inc_mapas", ["id", "nombre", "ruta", "orden", "ancho", "alto"]),
                    ("inc_areas", ["id", "mapa_id", "nombre", "x1", "y1", "x2", "y2", "color"]),
                    ("inc_maquinas", ["id", "area_id", "nombre", "serie", "numero_asignado", "x1", "y1", "x2", "y2", "color"]),
                    ("inc_incidencias", [
                        "id", "mapa_id", "area_id", "maquina_id", "fecha",
                        "creador_nombre", "creador_apellido1", "creador_apellido2",
                        "creador_movil", "creador_email",
                        "elemento", "descripcion", "estado", "reporte_path"
                    ]),
                ]
                for table, cols in tables:
                    rows = fetch_rows(conn_sqlite, table, cols)
                    inserted = insert_rows(conn_pg, table, cols, rows)
                    print(f"{table}: {inserted} filas")
                for table, cols in tables:
                    if "id" in cols:
                        set_sequence(conn_pg, table)
            finally:
                conn_sqlite.close()

        conn_pg.commit()

    print("Migracion completada.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
