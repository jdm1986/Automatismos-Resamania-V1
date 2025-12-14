import json
import unicodedata
from datetime import datetime, timedelta

import pandas as pd

from utils.file_loader import load_data_file


def _normalize(text: str) -> str:
    """Uppercase, trim and strip accents to compare column names reliably."""
    normalized = unicodedata.normalize("NFD", str(text or "")).upper().strip()
    return "".join(ch for ch in normalized if unicodedata.category(ch) != "Mn")


def _parse_fecha(value):
    """Parse dates similar a parseExcelDate del script original."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None

    if isinstance(value, (pd.Timestamp, datetime)):
        return value.to_pydatetime() if hasattr(value, "to_pydatetime") else value

    if isinstance(value, (int, float)):
        excel_epoch = datetime(1899, 12, 30)
        try:
            return excel_epoch + timedelta(days=float(value))
        except Exception:
            return None

    if isinstance(value, str):
        text = value.strip()
        if not text:
            return None

        fecha_part = text.split(" ")[0]
        sep = "/" if "/" in fecha_part else "-" if "-" in fecha_part else None
        if sep:
            parts = fecha_part.split(sep)
            if len(parts) == 3:
                try:
                    d, m, y = map(int, parts)
                    return datetime(y, m, d)
                except Exception:
                    pass

        try:
            dt = pd.to_datetime(text, dayfirst=True, errors="coerce")
            if pd.notna(dt):
                return dt.to_pydatetime()
        except Exception:
            return None

    return None


def obtener_cumpleanos_hoy() -> pd.DataFrame:
    """
    Devuelve clientes cuyo cumple es hoy (día y mes) y estado Cliente.
    Lee RESUMEN CLIENTE de la carpeta configurada.
    """
    with open("config.json", "r") as f:
        config = json.load(f)
    carpeta_datos = config.get("carpeta_datos", "")

    try:
        df = load_data_file(carpeta_datos, "RESUMEN CLIENTE")
    except FileNotFoundError:
        return pd.DataFrame(columns=["Numero de cliente", "Nombre", "Apellidos", "Fecha de nacimiento", "Email", "Estado"])

    colmap = {_normalize(col): col for col in df.columns}
    col_fecha_nac = None
    # cubrir "Fecha de nacimiento" y variantes
    for key, col in colmap.items():
        if key.startswith("FECHA DE NAC"):
            col_fecha_nac = col
            break
    col_estado = colmap.get("ESTADO")
    col_cliente = colmap.get("NUMERO DE CLIENTE") or colmap.get("NカMERO DE CLIENTE")
    col_nombre = colmap.get("NOMBRE")
    col_apellidos = colmap.get("APELLIDOS")
    col_email = colmap.get("EMAIL") or colmap.get("CORREO")

    if not col_fecha_nac:
        return pd.DataFrame(columns=["Numero de cliente", "Nombre", "Apellidos", "Fecha de nacimiento", "Email", "Estado"])

    hoy = datetime.now()
    hoy_dia = hoy.day
    hoy_mes = hoy.month

    fechas = df[col_fecha_nac].apply(_parse_fecha)
    mask_fecha = fechas.apply(lambda f: f is not None and f.day == hoy_dia and f.month == hoy_mes)

    if col_estado:
        mask_estado = df[col_estado].astype(str).str.strip().str.lower() == "cliente"
    else:
        mask_estado = True

    filtrado = df[mask_fecha & mask_estado].copy()

    # Ordenar columnas clave primero
    preferred = [c for c in [col_cliente, col_nombre, col_apellidos, col_email, col_fecha_nac, col_estado] if c in filtrado.columns]
    otras = [c for c in filtrado.columns if c not in preferred]
    filtrado = filtrado[preferred + otras]

    return filtrado.reset_index(drop=True)
