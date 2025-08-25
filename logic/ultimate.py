import os
import pandas as pd
import json

def obtener_socios_ultimate():
    # Cargar ruta desde config.json
    with open("config.json", "r") as f:
        config = json.load(f)
    carpeta_datos = config.get("carpeta_datos", "")

    # Ruta completa del archivo
    ruta_archivo = os.path.join(carpeta_datos, "FACTURAS Y VALES.xlsx")

    if not os.path.exists(ruta_archivo):
        return pd.DataFrame(columns=["número de cliente", "Nombre", "Apellidos"])

    df = pd.read_excel(ruta_archivo)

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
