import pandas as pd
from datetime import datetime, timedelta

def _find_column(df, keywords):
    """
    Busca la primera columna cuyo nombre (minúsculas, strip) contenga todos los keywords.
    """
    for col in df.columns:
        normal = col.lower().strip()
        if all(k in normal for k in keywords):
            return col
    return None

def procesar_accesos_dobles(resumen_df, accesos_df):
    # Filtrar solo "Cliente" desde RESUMEN
    resumen_df = resumen_df[resumen_df["Estado"].str.lower() == "cliente"]

    # Filtrar accesos de los últimos 7 días
    accesos_df["Fecha corta de acceso"] = pd.to_datetime(
        accesos_df["Fecha corta de acceso"], errors="coerce", dayfirst=True
    )
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

def procesar_accesos_dobles_ayer(resumen_df, accesos_df):
    """
    Accesos dobles SOLO del día anterior usando Entrada_trípode_1/2.
    """
    resumen_df = resumen_df[resumen_df["Estado"].str.lower() == "cliente"]

    col_cliente = _find_column(accesos_df, ["número", "cliente"]) or _find_column(accesos_df, ["numero", "cliente"])
    col_fecha = _find_column(accesos_df, ["fecha", "acceso"])
    col_punto = _find_column(accesos_df, ["punto", "acceso"])
    if not all([col_cliente, col_fecha, col_punto]):
        return pd.DataFrame(columns=["Nombre", "Apellidos", "Número de cliente", "Correo electrónico", "Móvil"])

    accesos_df = accesos_df.copy()
    accesos_df[col_fecha] = pd.to_datetime(accesos_df[col_fecha], errors="coerce", dayfirst=True)
    ayer = datetime.today().date() - timedelta(days=1)
    accesos_ayer = accesos_df[accesos_df[col_fecha].dt.date == ayer]

    entradas = accesos_ayer[accesos_ayer[col_punto].isin(["Entrada_trípode_1", "Entrada_trípode_2"])]
    entradas = entradas.dropna(subset=[col_fecha, col_cliente])
    entradas[col_cliente] = entradas[col_cliente].astype(str)
    stats = entradas.groupby(col_cliente)[col_fecha].agg(["min", "max", "size"])
    stats["delta"] = stats["max"] - stats["min"]
    dobles = stats[(stats["size"] >= 2) & (stats["delta"] >= timedelta(hours=2))].index.astype(str)

    # Columnas preferidas desde RESUMEN
    col_nombre = _find_column(resumen_df, ["nombre"])
    col_apellidos = _find_column(resumen_df, ["apell"])
    col_email = _find_column(resumen_df, ["correo"]) or _find_column(resumen_df, ["email"])
    col_movil = _find_column(resumen_df, ["movil"]) or _find_column(resumen_df, ["tel"])
    col_res_cliente = _find_column(resumen_df, ["número", "cliente"]) or _find_column(resumen_df, ["numero", "cliente"])

    # Columnas fallback desde ACCESOS
    col_acc_nombre = _find_column(accesos_df, ["nombre"])
    col_acc_apellidos = _find_column(accesos_df, ["apell"])
    col_acc_email = _find_column(accesos_df, ["correo"]) or _find_column(accesos_df, ["email"])
    col_acc_movil = _find_column(accesos_df, ["movil"]) or _find_column(accesos_df, ["tel"])

    if col_res_cliente:
        resultado = resumen_df[resumen_df[col_res_cliente].astype(str).isin(dobles)].copy()
    else:
        resultado = pd.DataFrame(columns=[col_nombre, col_apellidos, col_res_cliente, col_email, col_movil])

    # Fallback con ACCESOS para completar datos si faltan
    if col_cliente in accesos_df.columns:
        acc_yer = accesos_ayer.copy()
        acc_yer[col_cliente] = acc_yer[col_cliente].astype(str).str.strip()
        base = acc_yer[acc_yer[col_cliente].astype(str).isin(dobles)]
        base = base.drop_duplicates(subset=[col_cliente])
        base = base.rename(columns={
            col_cliente: "Número de cliente",
            col_acc_nombre: "Nombre",
            col_acc_apellidos: "Apellidos",
            col_acc_email: "Correo electrónico",
            col_acc_movil: "Móvil",
        })
        base = base[["Número de cliente", "Nombre", "Apellidos", "Correo electrónico", "Móvil"]]
        if col_res_cliente and not resultado.empty:
            resultado = resultado.rename(columns={
                col_res_cliente: "Número de cliente",
                col_nombre: "Nombre",
                col_apellidos: "Apellidos",
                col_email: "Correo electrónico",
                col_movil: "Móvil",
            })
            resultado = resultado[["Número de cliente", "Nombre", "Apellidos", "Correo electrónico", "Móvil"]]
            resultado = resultado.merge(base, on="Número de cliente", how="left", suffixes=("", "_acc"))
            for field in ["Nombre", "Apellidos", "Correo electrónico", "Móvil"]:
                acc_field = f"{field}_acc"
                if acc_field in resultado.columns:
                    resultado[field] = resultado[field].fillna(resultado[acc_field])
            resultado = resultado[["Nombre", "Apellidos", "Número de cliente", "Correo electrónico", "Móvil"]]
        else:
            resultado = base.rename(columns={"Número de cliente": "Número de cliente"})
            resultado = resultado[["Nombre", "Apellidos", "Número de cliente", "Correo electrónico", "Móvil"]]

    if resultado.empty:
        return pd.DataFrame(columns=["Nombre", "Apellidos", "Número de cliente", "Correo electrónico", "Móvil"])
    return resultado.fillna("")

def procesar_accesos_descuadrados(resumen_df, accesos_df):
    resumen_df = resumen_df[resumen_df["Estado"].str.lower() == "cliente"]

    accesos_df["Fecha corta de acceso"] = pd.to_datetime(
        accesos_df["Fecha corta de acceso"], errors="coerce", dayfirst=True
    )
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
    accesos_df["Fecha corta de acceso"] = pd.to_datetime(
        accesos_df["Fecha corta de acceso"], errors="coerce", dayfirst=True
    )
    accesos_rango = accesos_df[
        (accesos_df["Fecha corta de acceso"].dt.date >= hace_una_semana) &
        (accesos_df["Fecha corta de acceso"].dt.date <= hoy)
        ]

    # Filtrar solo salidas por PMR
    salidas_pmr = accesos_rango[accesos_rango["Punto de acceso del Pasaje"] == "Pmr_salida_1"]

    if salidas_pmr.empty:
        return resumen_df.head(0)[["Nombre", "Apellidos", "N??mero de cliente", "Correo electr??nico", "M??vil"]]

    # ??ltimo acceso PMR por cliente
    ultimo = (
        salidas_pmr.groupby("N??mero de cliente")["Fecha corta de acceso"]
        .max()
        .reset_index()
    )
    ultimo["Ultimo acceso PMR"] = ultimo["Fecha corta de acceso"].dt.strftime("%d/%m/%Y %H:%M")

    # Extraer informaci??n desde el resumen
    resultado = resumen_df[resumen_df["N??mero de cliente"].isin(ultimo["N??mero de cliente"])].copy()
    resultado = resultado.merge(
        ultimo[["N??mero de cliente", "Ultimo acceso PMR"]],
        on="N??mero de cliente",
        how="left",
    )

    # Columnas clave para mostrar
    cols = ["Nombre", "Apellidos", "N??mero de cliente", "Correo electr??nico", "M??vil", "Ultimo acceso PMR"]
    return resultado[cols]

    # Convertir y filtrar por fecha
    accesos_df["Fecha corta de acceso"] = pd.to_datetime(
        accesos_df["Fecha corta de acceso"], errors="coerce", dayfirst=True
    )
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
    # Copias para no mutar originales
    accesos_df = accesos_df.copy()
    incidencias_df = incidencias_df.copy()

    col_cliente_impagos = _find_column(incidencias_df, ["número", "cliente"]) or _find_column(incidencias_df, ["numero", "cliente"])
    if not col_cliente_impagos:
        raise ValueError("Falta la columna 'Número de cliente' en IMPAGOS.csv")

    col_cliente_accesos = _find_column(accesos_df, ["número", "cliente"]) or _find_column(accesos_df, ["numero", "cliente"])
    col_fecha = _find_column(accesos_df, ["fecha", "acceso"])
    col_paso = _find_column(accesos_df, ["punto", "acceso"])
    col_nombre = _find_column(accesos_df, ["nombre"])
    col_apellidos = _find_column(accesos_df, ["apell"])
    col_bloqueado = None
    for col in accesos_df.columns:
        if "bloque" in col.lower():
            col_bloqueado = col
            break

    if not all([col_cliente_accesos, col_fecha, col_paso, col_nombre, col_apellidos]):
        raise ValueError("Faltan columnas requeridas en ACCESOS.csv (cliente/fecha/punto/nombre/apellidos).")
    if not col_bloqueado:
        raise ValueError("Falta la columna de 'acceso bloqueado' en ACCESOS.csv.")

    # Lista de morosos (deduplicada)
    clientes_morosos = incidencias_df[col_cliente_impagos].dropna().unique()

    # Parseo de fecha y filtro últimos 7 días
    accesos_df[col_fecha] = pd.to_datetime(accesos_df[col_fecha], errors='coerce', dayfirst=True)
    hace_7_dias = datetime.today() - timedelta(days=7)
    accesos_recientes = accesos_df[accesos_df[col_fecha] >= hace_7_dias]

    # Filtrar accesos de morosos y con bloqueo activo (valor 1)
    accesos_filtrados = accesos_recientes[
        (accesos_recientes[col_cliente_accesos].isin(clientes_morosos)) &
        (accesos_recientes[col_bloqueado] == 1)
    ]

    if accesos_filtrados.empty:
        return pd.DataFrame(columns=["Nombre", "Apellidos", "Número de cliente", "Último acceso", "Acceso por", "Bloqueado"])

    # Tomar último acceso por cliente
    ultimos = accesos_filtrados.sort_values(col_fecha).groupby(col_cliente_accesos).tail(1)

    # Armar resultado
    resultado = pd.DataFrame({
        "Número de cliente": ultimos[col_cliente_accesos],
        "Nombre": ultimos[col_nombre],
        "Apellidos": ultimos[col_apellidos],
        "Último acceso": ultimos[col_fecha].dt.strftime("%d/%m/%Y %H:%M"),
        "Acceso por": ultimos[col_paso],
        "Bloqueado": ultimos[col_bloqueado],
    })

    return resultado.reset_index(drop=True)

