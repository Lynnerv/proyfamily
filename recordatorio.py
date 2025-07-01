import pandas as pd
from datetime import datetime, timedelta
from tkinter import messagebox, Tk
import os
import pytz

archivo = "cumpleaños.xlsx"

if not os.path.exists(archivo):
    print("No hay registros aún.")
    exit()

df = pd.read_excel(archivo)
df = df.dropna(subset=["Fecha de Nacimiento"])

# Zona horaria GMT-5 (America/Lima)
zona = pytz.timezone('America/Lima')
ahora_con_zona = datetime.now(zona)
hoy = ahora_con_zona.date()

# Listas para alertas
alerta_7_dias = []
alerta_hoy = []

# Lista para tabla resumen
tabla_resumen = []

for _, fila in df.iterrows():
    fecha_nac = fila["Fecha de Nacimiento"]
    if isinstance(fecha_nac, str):
        fecha_nac = datetime.strptime(fecha_nac, "%Y-%m-%d").date()

    nombre_completo = f"{fila['Nombre']} {fila['Apellido Paterno']} {fila['Apellido Materno']}"

    # Fecha de cumpleaños este año
    cumple_este_anio = fecha_nac.replace(year=hoy.year)

    # Si ya pasó este año, considerar cumpleaños del próximo año
    if cumple_este_anio < hoy:
        cumple_este_anio = cumple_este_anio.replace(year=hoy.year + 1)

    dias_restantes = (cumple_este_anio - hoy).days

    # Edad que cumplirá en esa fecha
    edad = cumple_este_anio.year - fecha_nac.year

    # Añadimos a tabla resumen
    tabla_resumen.append({
        "Nombre": nombre_completo,
        "Edad a cumplir": edad,
        "Días restantes": dias_restantes
    })

    # Generar alertas si faltan 7 días o 0 días
    if dias_restantes == 7:
        alerta_7_dias.append(f"{nombre_completo} cumplirá {edad} años en 7 días.")
    elif dias_restantes == 0:
        alerta_hoy.append(f"¡Hoy es el cumpleaños de {nombre_completo}! 🎉 Cumple {edad} años.")

# Mostrar tabla resumen en consola
print("\n📋 Tabla de cumpleaños:")
print("{:<35} {:<15} {:<15}".format("Nombre", "Edad a cumplir", "Días restantes"))
print("-" * 65)
for persona in tabla_resumen:
    print(f"{persona['Nombre']:<35} {persona['Edad a cumplir']:<15} {persona['Días restantes']:<15}")

# Mostrar alertas popup
root = Tk()
root.withdraw()

if alerta_7_dias:
    messagebox.showinfo("⏰ Cumpleaños en 7 días", "\n".join(alerta_7_dias))

if alerta_hoy:
    messagebox.showinfo("🎉 Cumpleaños Hoy", "\n".join(alerta_hoy))

if not alerta_7_dias and not alerta_hoy:
    print("\nNo hay alertas de cumpleaños hoy ni en 7 días.")
