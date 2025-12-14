import json
import os
import unicodedata

import pandas as pd


def _normalize(value) -> str:
    """Trim + uppercase + remove accents to compare text robustly."""
    text = unicodedata.normalize("NFD", str(value or "")).upper().strip()
    return "".join(ch for ch in text if unicodedata.category(ch) != "Mn")


def _find_facturas_file(carpeta_datos: str) -> str:
    """Return the first existing path for FACTURAS Y VALES with supported extensions."""
    for ext in (".csv", ".xlsx"):
        candidate = os.path.join(carpeta_datos, f"FACTURAS Y VALES{ext}")
        if os.path.exists(candidate):
            return candidate
    raise FileNotFoundError("No se encontro FACTURAS Y VALES.csv/xlsx en la carpeta configurada.")


def _leer_facturas_sin_cabecera(ruta: str) -> pd.DataFrame:
    """
    Carga el archivo sin asumir fila de cabeceras para poder detectarla (mismas opciones que load_data_file).
    """
    if ruta.lower().endswith(".xlsx"):
        return pd.read_excel(ruta, header=None)

    # CSV
    try:
        return pd.read_csv(ruta, sep=None, engine="python", header=None, encoding="utf-8-sig")
    except UnicodeDecodeError:
        try:
            return pd.read_csv(ruta, sep=None, engine="python", header=None, encoding="latin-1")
        except Exception:
            return pd.read_csv(ruta, sep=";", header=None, encoding="latin-1")
    except Exception:
        return pd.read_csv(ruta, sep=";", header=None, encoding="utf-8-sig")


def _detectar_cabeceras(df_raw: pd.DataFrame):
    """
    Busca la fila donde viven 'NOMBRE DEL PRODUCTO' y 'NUMERO DE CLIENTE' (normalizado).
    Devuelve (headers, data_df) o (None, None) si no se encuentra.
    """
    row_count = len(df_raw)
    max_scan = min(200, row_count)

    for idx in range(max_scan):
        row_norm = [_normalize(v) for v in df_raw.iloc[idx].tolist()]
        if "NOMBRE DEL PRODUCTO" in row_norm and "NUMERO DE CLIENTE" in row_norm:
            headers = [str(v).strip() for v in df_raw.iloc[idx].tolist()]
            data = df_raw.iloc[idx + 1 :].copy()
            data.columns = headers
            # Quita filas completamente vacias que suelen venir despues de la tabla.
            data = data.dropna(how="all")
            return headers, data
    return None, None


def obtener_socios_ultimate():
    """
    Replica la logica del script de Excel: detecta cabeceras aunque no esten en la primera fila,
    normaliza texto y filtra filas que contengan 'ULTIMATE' en el nombre de producto, dejando
    un unico registro por cliente.
    """
    with open("config.json", "r") as f:
        config = json.load(f)
    carpeta_datos = config.get("carpeta_datos", "")

    try:
        ruta_facturas = _find_facturas_file(carpeta_datos)
        df_raw = _leer_facturas_sin_cabecera(ruta_facturas)
    except FileNotFoundError:
        return pd.DataFrame(columns=["Numero de cliente", "Nombre", "Apellidos"])

    headers, df = _detectar_cabeceras(df_raw)
    if headers is None:
        return pd.DataFrame(columns=["Numero de cliente", "Nombre", "Apellidos"])

    columnas_norm = {_normalize(col): col for col in df.columns}
    col_producto = columnas_norm.get("NOMBRE DEL PRODUCTO")
    col_cliente = columnas_norm.get("NUMERO DE CLIENTE")

    if not col_producto or not col_cliente:
        return pd.DataFrame(columns=["Numero de cliente", "Nombre", "Apellidos"])

    df[col_cliente] = df[col_cliente].fillna("").astype(str).str.strip()
    productos_norm = df[col_producto].apply(_normalize)

    df_filtrado = df[
        (df[col_cliente] != "")
        & productos_norm.str.contains("ULTIMATE", na=False)
    ].copy()

    df_filtrado = df_filtrado.drop_duplicates(subset=[col_cliente])
    return df_filtrado.reset_index(drop=True)
