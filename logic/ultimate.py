import pandas as pd
import json
from utils.file_loader import load_data_file

def obtener_socios_ultimate():
    # Cargar ruta desde config.json
    with open("config.json", "r") as f:
        config = json.load(f)
    carpeta_datos = config.get("carpeta_datos", "")

    try:
        df = load_data_file(carpeta_datos, "FACTURAS Y VALES")
    except FileNotFoundError:
        return pd.DataFrame(columns=["número de cliente", "Nombre", "Apellidos"])

    if "Número de cliente" not in df.columns or "Nombre del producto" not in df.columns:
        return pd.DataFrame(columns=["número de cliente", "Nombre", "Apellidos"])

    # Filtrar socios con producto Ultimate
    df_filtrado = df[
        (df["Nombre del producto"] == "PACK ULTIMATE") |
        (df["Nombre del producto"].astype(str).str.contains("ULTIMATE - renovación", case=False, na=False))
        ]

    # Limpiar y renombrar columnas
    df_filtrado = df_filtrado[["Número de cliente", "Nombre", "Apellidos"]].drop_duplicates()
    df_filtrado = df_filtrado.rename(columns={"Número de cliente": "número de cliente"})

    return df_filtrado.reset_index(drop=True)
