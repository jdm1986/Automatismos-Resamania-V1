import pandas as pd
from datetime import datetime, timedelta

def procesar_accesos_dobles(resumen_df, accesos_df):
    # Filtrar solo "Cliente" desde RESUMEN
    resumen_df = resumen_df[resumen_df["Estado"].str.lower() == "cliente"]

    # Filtrar accesos de los últimos 7 días
    accesos_df["Fecha corta de acceso"] = pd.to_datetime(accesos_df["Fecha corta de acceso"], errors="coerce")
    una_semana = datetime.today() - timedelta(days=7)
    accesos_df = accesos_df[accesos_df["Fecha corta de acceso"] >= una_semana]

    # Solo entradas válidas
    entradas = accesos_df[accesos_df["Punto de acceso del Pasaje"].isin(["Entrada_trípode_1", "Entrada_trípode_2"])]

    # Agrupar por cliente + día y contar accesos
    conteo = entradas.groupby(["Número de cliente", entradas["Fecha corta de acceso"].dt.date]).size().reset_index(name="conteo")
    dobles = conteo[conteo["conteo"] >= 2]["Número de cliente"].unique()

    # Cruce con RESUMEN
    resultados = resumen_df[resumen_df["Número de cliente"].isin(dobles)]

    return resultados[["Nombre", "Apellidos", "Número de cliente", "Correo electrónico", "Móvil"]]

def procesar_accesos_descuadrados(resumen_df, accesos_df):
    resumen_df = resumen_df[resumen_df["Estado"].str.lower() == "cliente"]

    accesos_df["Fecha corta de acceso"] = pd.to_datetime(accesos_df["Fecha corta de acceso"], errors="coerce")
    hace_7_dias = datetime.today() - timedelta(days=7)
    accesos_df = accesos_df[accesos_df["Fecha corta de acceso"] >= hace_7_dias]

    entradas = accesos_df[accesos_df["Punto de acceso del Pasaje"].isin(["Entrada_trípode_1", "Entrada_trípode_2"])]
    salidas = accesos_df[accesos_df["Punto de acceso del Pasaje"].isin(["Salida_trípode_1", "Salida_trípode_2", "Pmr_salida_1"])]

    entradas_count = entradas.groupby("Número de cliente").size()
    salidas_count = salidas.groupby("Número de cliente").size()

    descuadrados = []

    for cliente_id in entradas_count.index:
        entradas_n = entradas_count.get(cliente_id, 0)
        salidas_n = salidas_count.get(cliente_id, 0)
        if entradas_n != salidas_n:
            descuadrados.append(cliente_id)

    resultado = resumen_df[resumen_df["Número de cliente"].isin(descuadrados)]
    return resultado[["Nombre", "Apellidos", "Número de cliente", "Correo electrónico", "Móvil"]]



def procesar_salidas_pmr_no_autorizadas(resumen_df, accesos_df):
    hoy = datetime.today().date()
    hace_una_semana = hoy - timedelta(days=7)

    # Convertir y filtrar por fecha
    accesos_df["Fecha corta de acceso"] = pd.to_datetime(accesos_df["Fecha corta de acceso"], errors="coerce")
    accesos_rango = accesos_df[
        (accesos_df["Fecha corta de acceso"].dt.date >= hace_una_semana) &
        (accesos_df["Fecha corta de acceso"].dt.date <= hoy)
        ]

    # Filtrar solo salidas por PMR
    salidas_pmr = accesos_rango[accesos_rango["Punto de acceso del Pasaje"] == "Pmr_salida_1"]

    # Obtener clientes únicos que han usado esa salida
    clientes_pmr = salidas_pmr["Número de cliente"].unique()

    # Extraer información desde el resumen
    resultado = resumen_df[resumen_df["Número de cliente"].isin(clientes_pmr)]

    # Columnas clave para mostrar
    return resultado[["Nombre", "Apellidos", "Número de cliente", "Correo electrónico", "Móvil"]]

def procesar_morosos_activos(incidencias_df, accesos_df=None):
    # Convertimos nombres de columnas a mayúsculas por compatibilidad
    df = incidencias_df.copy()
    df.columns = [col.strip().lower() for col in df.columns]

    # Aseguramos que las columnas clave están presentes
    if 'estado del incidente' not in df.columns or 'número de cliente' not in df.columns:
        raise ValueError("Falta la columna 'Estado del incidente' o 'Número de cliente'")

    # Filtramos morosos: solo los que están en estado "Abierto"
    morosos_df = df[df['estado del incidente'].str.upper() == 'ABIERTO']

    # Eliminamos duplicados por número de cliente
    morosos_df = morosos_df.drop_duplicates(subset='número de cliente')

    # Devolvemos el DataFrame de morosos activos
    return morosos_df

def procesar_morosos_accediendo(incidencias_df, accesos_df):
    from datetime import datetime, timedelta
    import pandas as pd

    # Normalizar columnas
    incidencias_df.columns = [c.lower().strip() for c in incidencias_df.columns]
    accesos_df.columns = [c.lower().strip() for c in accesos_df.columns]

    # Filtrar solo morosos activos
    morosos = incidencias_df[incidencias_df['estado del incidente'].str.upper() == 'ABIERTO']
    clientes_morosos = morosos['número de cliente'].unique()

    # Convertir fecha y filtrar últimos 7 días
    accesos_df['fecha de acceso'] = pd.to_datetime(accesos_df['fecha de acceso'], errors='coerce')
    hace_7_dias = datetime.today() - timedelta(days=7)
    accesos_recientes = accesos_df[accesos_df['fecha de acceso'] >= hace_7_dias]

    # Filtrar accesos de morosos
    accesos_morosos = accesos_recientes[accesos_recientes['número de cliente'].isin(clientes_morosos)]

    # Tomar último acceso por cliente
    ultimos = accesos_morosos.sort_values("fecha de acceso").groupby("número de cliente").tail(1)

    # Unir con datos básicos (por si queremos nombre, etc.)
    resumen = accesos_morosos.drop_duplicates(subset="número de cliente")[["número de cliente", "nombre", "apellidos"]]
    resultado = pd.merge(resumen, ultimos, on="número de cliente", suffixes=("_x", "_y"))

    # Formatear la hora legible
    resultado["fecha de acceso"] = resultado["fecha de acceso"].dt.strftime("%d/%m/%Y %H:%M")

    # Devolver columnas organizadas
    return resultado[["nombre_x", "apellidos_x", "número de cliente", "fecha de acceso", "punto de acceso del pasaje"]].rename(
        columns={
            "nombre_x": "Nombre",
            "apellidos_x": "Apellidos",
            "fecha de acceso": "Último acceso",
            "punto de acceso del pasaje": "Acceso por"
        }
    )




