import json
import uuid

try:
    import psycopg
    from psycopg.types.json import Json
except Exception:
    psycopg = None
    Json = None


class AppStateStore:
    def __init__(self, db_config):
        self.db_config = db_config or {}
        self.use_postgres = bool(self.db_config.get("host"))
        if self.use_postgres:
            self._init_db()

    def _connect(self):
        if not self.use_postgres:
            return None
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

    def _init_db(self):
        with self._connect() as conn:
            cur = conn.cursor()
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS app_state (
                    key TEXT PRIMARY KEY,
                    value JSONB,
                    updated_at TIMESTAMP DEFAULT NOW()
                )
                """
            )
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS app_blobs (
                    id TEXT PRIMARY KEY,
                    content_type TEXT,
                    data BYTEA,
                    created_at TIMESTAMP DEFAULT NOW()
                )
                """
            )
            conn.commit()

    def get(self, key, default=None):
        if not self.use_postgres:
            return default
        with self._connect() as conn:
            cur = conn.cursor()
            cur.execute("SELECT value FROM app_state WHERE key=%s", (key,))
            row = cur.fetchone()
            if not row:
                return default
            return row[0]

    def set(self, key, value):
        if not self.use_postgres:
            return
        payload = Json(value) if Json is not None else json.dumps(value)
        with self._connect() as conn:
            cur = conn.cursor()
            cur.execute(
                """
                INSERT INTO app_state (key, value, updated_at)
                VALUES (%s, %s, NOW())
                ON CONFLICT (key) DO UPDATE SET
                    value = EXCLUDED.value,
                    updated_at = NOW()
                """,
                (key, payload),
            )
            conn.commit()

    def delete(self, key):
        if not self.use_postgres:
            return
        with self._connect() as conn:
            cur = conn.cursor()
            cur.execute("DELETE FROM app_state WHERE key=%s", (key,))
            conn.commit()

    def put_blob(self, data, content_type="application/octet-stream"):
        if not self.use_postgres:
            return ""
        blob_id = uuid.uuid4().hex
        with self._connect() as conn:
            cur = conn.cursor()
            cur.execute(
                """
                INSERT INTO app_blobs (id, content_type, data)
                VALUES (%s, %s, %s)
                """,
                (blob_id, content_type, data),
            )
            conn.commit()
        return blob_id

    def get_blob(self, blob_id):
        if not self.use_postgres:
            return None, None
        with self._connect() as conn:
            cur = conn.cursor()
            cur.execute("SELECT content_type, data FROM app_blobs WHERE id=%s", (blob_id,))
            row = cur.fetchone()
            if not row:
                return None, None
            return row[0], row[1]

    def delete_blob(self, blob_id):
        if not self.use_postgres:
            return
        with self._connect() as conn:
            cur = conn.cursor()
            cur.execute("DELETE FROM app_blobs WHERE id=%s", (blob_id,))
            conn.commit()
