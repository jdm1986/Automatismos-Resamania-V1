import pandas as pd
from datetime import datetime, timedelta
from collections import Counter

def calcular_dias_alta(fecha_alta):
    hoy = datetime.today().date()
    return (hoy - fecha_alta.date()).days

def calcular_franja_horaria(accesos_cliente):
    """ Devuelve la franja horaria más frecuente de acceso """
    if accesos_cliente.empty:
        return "Sin datos"

    horas = pd.to_datetime(accesos_cliente["Fecha de acceso"], dayfirst=True, errors="coerce").dt.hour
    franjas = []

    for h in horas:
        if 6 <= h < 8: franjas.append("06:00 - 08:00")
        elif 8 <= h < 10: franjas.append("08:00 - 10:00")
        elif 10 <= h < 12: franjas.append("10:00 - 12:00")
        elif 12 <= h < 14: franjas.append("12:00 - 14:00")
        elif 14 <= h < 16: franjas.append("14:00 - 16:00")
        elif 16 <= h < 18: franjas.append("16:00 - 18:00")
        elif 18 <= h < 20: franjas.append("18:00 - 20:00")
        elif 20 <= h < 22: franjas.append("20:00 - 22:00")
        elif 22 <= h <= 23: franjas.append("22:00 - 00:00")
        else: franjas.append("00:00 - 06:00")

    if franjas:
        return Counter(franjas).most_common(1)[0][0]
    return "Sin datos"

def procesar_wizville(resumen_df, accesos_df):
    resumen_df = resumen_df.copy()
    accesos_df = accesos_df.copy()

    # Filtrar solo clientes activos
    clientes = resumen_df[resumen_df["Estado"].str.lower() == "cliente"]

    # Parsear fecha de alta
    clientes.loc[:, "Inicio del abono"] = pd.to_datetime(clientes["Inicio del abono"], errors="coerce", dayfirst=True)
    clientes = clientes.dropna(subset=["Inicio del abono"])  # Elimina filas sin fecha válida
    clientes.loc[:, "Días desde alta"] = clientes["Inicio del abono"].apply(calcular_dias_alta)


# Filtrar solo quienes cumplen 16 o 180 días
    seleccion = clientes[clientes["Días desde alta"].isin([16, 180])]

    resultados = []

    for _, fila in seleccion.iterrows():
        numero = fila["Número de cliente"]
        accesos_cliente = accesos_df[accesos_df["Número de cliente"] == numero]

        franja = calcular_franja_horaria(accesos_cliente)

        resultados.append({
            "Nombre": fila["Nombre"],
            "Apellidos": fila["Apellidos"],
            "Número de cliente": numero,
            "Correo electrónico": fila["Correo electrónico"],
            "Móvil": fila["Móvil"],
            "Días desde alta": fila["Días desde alta"],
            "Franja probable": franja
        })

    return pd.DataFrame(resultados)
