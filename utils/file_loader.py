import io
import json
import os
import sys
import pandas as pd

from logic.state_store import AppStateStore


def _get_app_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def _get_config_path():
    base = _get_app_dir()
    candidate = os.path.join(base, "config.json")
    if os.path.exists(candidate):
        return candidate
    parent = os.path.dirname(base)
    candidate = os.path.join(parent, "config.json")
    return candidate


CONFIG_PATH = _get_config_path()


def _read_config():
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r") as f:
                return json.load(f)
        except json.JSONDecodeError:
            return {}
    return {}


def _get_db_config():
    data = _read_config()
    db = data.get("db", {}) if isinstance(data, dict) else {}
    return {
        "host": str(db.get("host", "")).strip(),
        "port": str(db.get("port", "5432")).strip(),
        "name": str(db.get("name", "resamania")).strip(),
        "user": str(db.get("user", "resamania")).strip(),
        "password": str(db.get("password", "")),
    }


def _load_from_db(base_name: str):
    db_cfg = _get_db_config()
    if not db_cfg.get("host"):
        return None
    store = AppStateStore(db_cfg)
    meta = store.get(f"export:{base_name}", {})
    if not isinstance(meta, dict):
        return None
    blob_id = str(meta.get("blob", "")).strip()
    filename = str(meta.get("filename", "")).strip()
    if not blob_id:
        return None
    _ctype, data = store.get_blob(blob_id)
    if not data:
        return None
    return filename, data


def load_data_file(folder_path: str, base_name: str) -> pd.DataFrame:
    """
    Load a data file prioritizing CSV (Resamania exports) and falling back to XLSX.
    Args:
        folder_path: Directory where the exports live.
        base_name: Filename without extension, e.g. "RESUMEN CLIENTE".
    Returns:
        pandas.DataFrame with the loaded contents.
    Raises:
        FileNotFoundError if no matching file is found.
    """
    db_payload = _load_from_db(base_name)
    if db_payload:
        filename, data = db_payload
        buf = io.BytesIO(data)
        if filename.lower().endswith(".xlsx"):
            return pd.read_excel(buf)
        try:
            return pd.read_csv(buf, sep=None, engine="python", encoding="utf-8-sig")
        except UnicodeDecodeError:
            try:
                buf.seek(0)
                return pd.read_csv(buf, sep=None, engine="python", encoding="latin-1")
            except Exception:
                buf.seek(0)
                return pd.read_csv(buf, sep=";", encoding="latin-1")
        except Exception:
            buf.seek(0)
            return pd.read_csv(buf, sep=";", encoding="utf-8-sig")

    candidates = [
        (f"{base_name}.csv", "csv"),
        (f"{base_name}.xlsx", "excel"),
    ]

    for filename, kind in candidates:
        full_path = os.path.join(folder_path, filename)
        if not os.path.exists(full_path):
            continue

        if kind == "csv":
            try:
                # sep=None with engine="python" infers delimiter (Resamania exports often use ';').
                return pd.read_csv(full_path, sep=None, engine="python", encoding="utf-8-sig")
            except UnicodeDecodeError:
                # Fallback a Latin-1 si el CSV viene en ANSI/Windows-1252.
                try:
                    return pd.read_csv(full_path, sep=None, engine="python", encoding="latin-1")
                except Exception:
                    return pd.read_csv(full_path, sep=";", encoding="latin-1")
            except Exception:
                # Fallback a separador ';' en UTF-8 si la inferencia falla por delimitador.
                return pd.read_csv(full_path, sep=";", encoding="utf-8-sig")

        return pd.read_excel(full_path)

    raise FileNotFoundError(
        f"No se encontro ninguno de estos archivos en {folder_path}: "
        f"{', '.join(name for name, _ in candidates)}"
    )
