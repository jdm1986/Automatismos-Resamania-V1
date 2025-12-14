import json
from datetime import datetime, timedelta
import unicodedata

import pandas as pd

from utils.file_loader import load_data_file


def _normalize(text: str) -> str:
    """Uppercase, trim and strip accents to compare column names and values reliably."""
    normalized = unicodedata.normalize("NFD", str(text or "")).upper().strip()
    return "".join(ch for ch in normalized if unicodedata.category(ch) != "Mn")


def _parse_fecha(value):
    """Imita parseExcelDate del script: admite Date/num de serie/strings."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None

    # pandas Timestamp o datetime
    if isinstance(value, (pd.Timestamp, datetime)):
        return value.to_pydatetime() if hasattr(value, "to_pydatetime") else value

    # Numero de serie Excel
    if isinstance(value, (int, float)):
        excel_epoch = datetime(1899, 12, 30)
        try:
            return excel_epoch + timedelta(days=float(value))
        except Exception:
            return None

    # Texto
    if isinstance(value, str):
        text = value.strip()
        if not text:
            return None

        # Separar parte de fecha si viene con hora
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


def _rango_martes_a_lunes():
    """
    Calcula el rango por defecto: desde el martes pasado (00:00) hasta el lunes actual (23:59).
    """
    hoy = datetime.now()
    today = datetime(hoy.year, hoy.month, hoy.day)
    day = today.weekday()  # 0 lunes ... 6 domingo

    monday_current = today - timedelta(days=day)  # lunes de la semana actual
    fecha_inicio = monday_current - timedelta(days=6)  # martes anterior
    fecha_inicio = fecha_inicio.replace(hour=0, minute=0, second=0, microsecond=0)

    fecha_fin = monday_current.replace(hour=23, minute=59, second=59, microsecond=999000)

    return fecha_inicio, fecha_fin


def obtener_avanza_fit() -> pd.DataFrame:
    """
    Filtra clientes de 'RESUMEN CLIENTE' cuya 'Fecha de creación' esté en el rango
    martes (semana pasada) a lunes (semana actual) y con Estado = Cliente.
    """
    with open("config.json", "r") as f:
        config = json.load(f)
    carpeta_datos = config.get("carpeta_datos", "")

    try:
        df = load_data_file(carpeta_datos, "RESUMEN CLIENTE")
    except FileNotFoundError:
        return pd.DataFrame(columns=["Numero de cliente", "Nombre", "Apellidos", "Fecha de creación", "Estado"])

    # Mapear columnas normalizadas
    colmap = {_normalize(col): col for col in df.columns}
    col_fecha = colmap.get("FECHA DE CREACION")
    col_estado = colmap.get("ESTADO")
    col_cliente = colmap.get("NUMERO DE CLIENTE") or colmap.get("NカMERO DE CLIENTE")
    col_nombre = colmap.get("NOMBRE")
    col_apellidos = colmap.get("APELLIDOS")

    if not col_fecha:
        return pd.DataFrame(columns=["Numero de cliente", "Nombre", "Apellidos", "Fecha de creación", "Estado"])

    fecha_inicio, fecha_fin = _rango_martes_a_lunes()

    # Parsear fechas y filtrar
    fechas = df[col_fecha].apply(_parse_fecha)
    mask_fecha = fechas.apply(lambda f: f is not None and fecha_inicio <= f <= fecha_fin)

    if col_estado:
        mask_estado = df[col_estado].astype(str).str.strip().str.lower() == "cliente"
    else:
        mask_estado = True

    filtrado = df[mask_fecha & mask_estado].copy()

    # Reordenar columnas para mostrar primero las clave
    preferred = [c for c in [col_cliente, col_nombre, col_apellidos, col_fecha, col_estado] if c in filtrado.columns]
    otras = [c for c in filtrado.columns if c not in preferred]
    filtrado = filtrado[preferred + otras]

    return filtrado.reset_index(drop=True)
