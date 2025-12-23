import sqlite3
from datetime import datetime
import unicodedata


def _norm(text: str) -> str:
    raw = unicodedata.normalize("NFD", str(text or "")).upper().strip()
    return "".join(ch for ch in raw if unicodedata.category(ch) != "Mn")


def _find_col(df, keywords):
    for col in df.columns:
        norm = _norm(col)
        if all(k in norm for k in keywords):
            return col
    return None


def normalize_impagos_df(df, resumen_map=None):
    """
    Devuelve una lista de dicts con columnas normalizadas:
    numero_cliente, nombre, apellidos, email, movil, incidentes
    """
    col_num = _find_col(df, ["NUMERO", "CLIENTE"])
    col_nombre = _find_col(df, ["NOMBRE"])  # puede coincidir con "Nombre de ventas", se filtra luego
    col_apellidos = _find_col(df, ["APELLIDOS"])
    col_email = _find_col(df, ["EMAIL"]) or _find_col(df, ["CORREO"])
    col_movil = _find_col(df, ["MOVIL"]) or _find_col(df, ["TELEFONO"])
    col_inc = _find_col(df, ["NUMERO", "INCIDENTE"]) or _find_col(df, ["INCIDENTE"])

    # Si hay varias columnas "Nombre", priorizar "Nombre" simple y no "Nombre de ventas"
    if col_nombre and "VENTA" in _norm(col_nombre):
        for c in df.columns:
            if _norm(c) == "NOMBRE":
                col_nombre = c
                break

    rows = []
    for _, row in df.iterrows():
        numero = str(row.get(col_num, "")).strip() if col_num else ""
        if not numero:
            continue
        nombre = str(row.get(col_nombre, "")).strip() if col_nombre else ""
        apellidos = str(row.get(col_apellidos, "")).strip() if col_apellidos else ""
        email = str(row.get(col_email, "")).strip() if col_email else ""
        movil = str(row.get(col_movil, "")).strip() if col_movil else ""
        if resumen_map and numero in resumen_map:
            ref = resumen_map[numero]
            if not nombre:
                nombre = ref.get("nombre", "")
            if not apellidos:
                apellidos = ref.get("apellidos", "")
            if not email:
                email = ref.get("email", "")
            if not movil:
                movil = ref.get("movil", "")
        try:
            incidentes = int(row.get(col_inc, 1)) if col_inc else 1
        except Exception:
            incidentes = 1
        rows.append({
            "numero_cliente": numero,
            "nombre": nombre,
            "apellidos": apellidos,
            "email": email,
            "movil": movil,
            "incidentes": incidentes,
        })
    return rows


class ImpagosDB:
    def __init__(self, db_path: str):
        self.db_path = db_path
        self.init_db()

    def _connect(self):
        return sqlite3.connect(self.db_path)

    def init_db(self):
        with self._connect() as conn:
            cur = conn.cursor()
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS impagos_clientes (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    numero_cliente TEXT UNIQUE,
                    nombre TEXT,
                    apellidos TEXT,
                    email TEXT,
                    movil TEXT
                )
                """
            )
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS impagos_eventos (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    cliente_id INTEGER,
                    fecha_export TEXT,
                    incidentes INTEGER,
                    UNIQUE(cliente_id, fecha_export),
                    FOREIGN KEY(cliente_id) REFERENCES impagos_clientes(id)
                )
                """
            )
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS impagos_gestion (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    cliente_id INTEGER,
                    fecha TEXT,
                    accion TEXT,
                    plantilla TEXT,
                    staff TEXT,
                    notas TEXT,
                    FOREIGN KEY(cliente_id) REFERENCES impagos_clientes(id)
                )
                """
            )
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS impagos_meta (
                    key TEXT PRIMARY KEY,
                    value TEXT
                )
                """
            )
            conn.commit()

    def set_last_export(self, fecha_export: str):
        with self._connect() as conn:
            cur = conn.cursor()
            cur.execute(
                "INSERT INTO impagos_meta(key, value) VALUES (?, ?) "
                "ON CONFLICT(key) DO UPDATE SET value=excluded.value",
                ("last_export", fecha_export),
            )
            conn.commit()

    def get_last_export(self):
        with self._connect() as conn:
            cur = conn.cursor()
            cur.execute("SELECT value FROM impagos_meta WHERE key='last_export'")
            row = cur.fetchone()
            return row[0] if row else None

    def get_prev_export(self, before_date):
        with self._connect() as conn:
            cur = conn.cursor()
            cur.execute(
                "SELECT MAX(fecha_export) FROM impagos_eventos WHERE fecha_export < ?",
                (before_date,),
            )
            row = cur.fetchone()
            return row[0] if row and row[0] else None

    def upsert_cliente(self, numero_cliente, nombre, apellidos, email, movil):
        with self._connect() as conn:
            cur = conn.cursor()
            cur.execute(
                """
                INSERT INTO impagos_clientes (numero_cliente, nombre, apellidos, email, movil)
                VALUES (?, ?, ?, ?, ?)
                ON CONFLICT(numero_cliente) DO UPDATE SET
                    nombre=excluded.nombre,
                    apellidos=excluded.apellidos,
                    email=excluded.email,
                    movil=excluded.movil
                """,
                (numero_cliente, nombre, apellidos, email, movil),
            )
            conn.commit()

    def get_cliente_id(self, numero_cliente):
        with self._connect() as conn:
            cur = conn.cursor()
            cur.execute("SELECT id FROM impagos_clientes WHERE numero_cliente=?", (numero_cliente,))
            row = cur.fetchone()
            return row[0] if row else None

    def add_evento(self, cliente_id, fecha_export, incidentes):
        with self._connect() as conn:
            cur = conn.cursor()
            cur.execute(
                """
                INSERT INTO impagos_eventos (cliente_id, fecha_export, incidentes)
                VALUES (?, ?, ?)
                ON CONFLICT(cliente_id, fecha_export) DO UPDATE SET
                    incidentes=excluded.incidentes
                """,
                (cliente_id, fecha_export, incidentes),
            )
            conn.commit()

    def add_gestion(self, cliente_id, accion, plantilla="", staff="", notas=""):
        with self._connect() as conn:
            cur = conn.cursor()
            cur.execute(
                """
                INSERT INTO impagos_gestion (cliente_id, fecha, accion, plantilla, staff, notas)
                VALUES (?, ?, ?, ?, ?, ?)
                """,
                (cliente_id, datetime.now().strftime("%Y-%m-%d %H:%M"), accion, plantilla, staff, notas),
            )
            conn.commit()

    def sync_from_df(self, df, resumen_map=None):
        fecha_export = datetime.now().date().isoformat()
        prev_export = self.get_prev_export(fecha_export)
        rows = normalize_impagos_df(df, resumen_map=resumen_map)
        for r in rows:
            self.upsert_cliente(
                r["numero_cliente"], r["nombre"], r["apellidos"], r["email"], r["movil"]
            )
            cliente_id = self.get_cliente_id(r["numero_cliente"])
            if cliente_id:
                self.add_evento(cliente_id, fecha_export, r["incidentes"])
        if prev_export and prev_export != fecha_export:
            self._marcar_resueltos(prev_export, fecha_export)
        self.set_last_export(fecha_export)
        return fecha_export, len(rows)

    def _marcar_resueltos(self, prev_export, current_export):
        """
        Marca como resueltos (accion resuelto_auto) a los clientes que
        estaban en el export anterior y ya no aparecen en el actual,
        solo si han recibido algún email alguna vez.
        """
        with self._connect() as conn:
            cur = conn.cursor()
            cur.execute(
                """
                WITH prev AS (
                  SELECT DISTINCT cliente_id FROM impagos_eventos WHERE fecha_export = ?
                ),
                curr AS (
                  SELECT DISTINCT cliente_id FROM impagos_eventos WHERE fecha_export = ?
                ),
                faltan AS (
                  SELECT p.cliente_id FROM prev p
                  LEFT JOIN curr c ON c.cliente_id = p.cliente_id
                  WHERE c.cliente_id IS NULL
                )
                SELECT f.cliente_id
                FROM faltan f
                WHERE EXISTS (
                  SELECT 1 FROM impagos_gestion g
                  WHERE g.cliente_id = f.cliente_id AND g.accion = 'email'
                )
                """,
                (prev_export, current_export),
            )
            clientes = [row[0] for row in cur.fetchall()]
            for cliente_id in clientes:
                cur.execute(
                    """
                    INSERT INTO impagos_gestion (cliente_id, fecha, accion, plantilla, staff, notas)
                    SELECT ?, ?, 'resuelto_auto', '', '', ''
                    WHERE NOT EXISTS (
                      SELECT 1 FROM impagos_gestion
                      WHERE cliente_id = ? AND accion = 'resuelto_auto' AND fecha = ?
                    )
                    """,
                    (cliente_id, f"{current_export} 00:00", cliente_id, f"{current_export} 00:00"),
                )
            conn.commit()

    def _base_current_query(self, fecha_export):
        return (
            "WITH email_hist AS ("
            "  SELECT cliente_id, GROUP_CONCAT(substr(fecha, 1, 10), ', ') AS email_hist "
            "  FROM impagos_gestion WHERE accion='email' GROUP BY cliente_id"
            "), cycle AS ("
            "  SELECT cliente_id, MAX(fecha) AS cycle_start "
            "  FROM impagos_gestion WHERE accion='resuelto_auto' GROUP BY cliente_id"
            "), last_email_cycle AS ("
            "  SELECT g.cliente_id, g.fecha AS last_email, g.plantilla AS last_plantilla "
            "  FROM impagos_gestion g "
            "  LEFT JOIN cycle c ON c.cliente_id = g.cliente_id "
            "  WHERE g.accion='email' "
            "    AND (c.cycle_start IS NULL OR g.fecha > c.cycle_start) "
            "    AND g.fecha = ("
            "      SELECT MAX(g2.fecha) FROM impagos_gestion g2 "
            "      WHERE g2.cliente_id = g.cliente_id AND g2.accion='email' "
            "        AND (c.cycle_start IS NULL OR g2.fecha > c.cycle_start)"
            "    )"
            "), prev_app AS ("
            "  SELECT cliente_id, MAX(fecha_export) AS prev_fecha "
            "  FROM impagos_eventos WHERE fecha_export < ? GROUP BY cliente_id"
            ") "
            "SELECT c.numero_cliente, c.nombre, c.apellidos, c.email, c.movil, "
            "e.incidentes, e.fecha_export, "
            "CASE WHEN lec.last_email IS NOT NULL THEN 1 ELSE 0 END AS email_enviado, "
            "lec.last_email AS fecha_envio, "
            "eh.email_hist AS email_hist, "
            "CASE WHEN cyc.cycle_start IS NOT NULL "
            "AND date(e.fecha_export) > date(cyc.cycle_start) "
            "THEN 1 ELSE 0 END AS reincidente "
            "FROM impagos_eventos e "
            "JOIN impagos_clientes c ON c.id = e.cliente_id "
            "LEFT JOIN cycle cyc ON cyc.cliente_id = c.id "
            "LEFT JOIN last_email_cycle lec ON lec.cliente_id = c.id "
            "LEFT JOIN email_hist eh ON eh.cliente_id = c.id "
            "LEFT JOIN prev_app pa ON pa.cliente_id = c.id "
            "WHERE e.fecha_export = ?"
        )

    def fetch_view(self, view, fecha_export):
        if not fecha_export:
            return []
        with self._connect() as conn:
            cur = conn.cursor()
            if view == "actuales":
                cur.execute(self._base_current_query(fecha_export), (fecha_export, fecha_export))
            elif view == "reincidentes":
                cur.execute(
                    self._base_current_query(fecha_export)
                    + " AND pa.prev_fecha IS NOT NULL AND (julianday(e.fecha_export) - julianday(pa.prev_fecha)) >= 2",
                    (fecha_export, fecha_export),
                )
            elif view == "incidentes1":
                cur.execute(
                    self._base_current_query(fecha_export)
                    + " AND e.incidentes = 1 AND lec.last_email IS NULL",
                    (fecha_export, fecha_export),
                )
            elif view == "incidentes2":
                cur.execute(
                    self._base_current_query(fecha_export)
                    + " AND e.incidentes >= 2 AND (lec.last_email IS NULL OR lec.last_plantilla != '2inc')",
                    (fecha_export, fecha_export),
                )
            elif view == "resueltos":
                cur.execute(
                    """
                    WITH email_hist AS (
                      SELECT cliente_id, GROUP_CONCAT(substr(fecha, 1, 10), ', ') AS email_hist
                      FROM impagos_gestion WHERE accion='email' GROUP BY cliente_id
                    ), cycle AS (
                      SELECT cliente_id, MAX(fecha) AS cycle_start
                      FROM impagos_gestion WHERE accion='resuelto_auto' GROUP BY cliente_id
                    ), last_email_cycle AS (
                      SELECT g.cliente_id, g.fecha AS last_email, g.plantilla AS last_plantilla
                      FROM impagos_gestion g
                      LEFT JOIN cycle c ON c.cliente_id = g.cliente_id
                      WHERE g.accion='email'
                        AND (c.cycle_start IS NULL OR g.fecha > c.cycle_start)
                        AND g.fecha = (
                          SELECT MAX(g2.fecha) FROM impagos_gestion g2
                          WHERE g2.cliente_id = g.cliente_id AND g2.accion='email'
                            AND (c.cycle_start IS NULL OR g2.fecha > c.cycle_start)
                        )
                    ), resuelto_hoy AS (
                      SELECT cliente_id FROM impagos_gestion
                      WHERE accion='resuelto_auto' AND fecha = ? || ' 00:00'
                    ), resuelto_email_hoy AS (
                      SELECT cliente_id FROM impagos_gestion
                      WHERE accion='resuelto_email' AND date(fecha) = ?
                    )
                    SELECT c.numero_cliente, c.nombre, c.apellidos, c.email, c.movil,
                           e.incidentes, e.fecha_export,
                           0 AS email_enviado,
                           '' AS fecha_envio,
                           eh.email_hist AS email_hist,
                           0 AS reincidente
                    FROM impagos_eventos e
                    JOIN impagos_clientes c ON c.id = e.cliente_id
                    LEFT JOIN email_hist eh ON eh.cliente_id = c.id
                    WHERE e.fecha_export = (
                        SELECT MAX(e2.fecha_export) FROM impagos_eventos e2 WHERE e2.cliente_id = c.id
                    )
                    AND c.id IN (SELECT cliente_id FROM resuelto_hoy)
                    AND c.id NOT IN (SELECT cliente_id FROM resuelto_email_hoy)
                    AND c.id NOT IN (SELECT cliente_id FROM impagos_eventos WHERE fecha_export = ?)
                    """,
                    (fecha_export, fecha_export, fecha_export),
                )
            else:
                cur.execute(self._base_current_query(fecha_export), (fecha_export, fecha_export))
            return cur.fetchall()
