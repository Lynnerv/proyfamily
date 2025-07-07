# ===================== 📦 IMPORTACIÓN DE MÓDULOS =====================
import pandas as pd                         # Para manejar datos tipo Excel
from datetime import datetime, timedelta    # Para manipular fechas
from tkinter import messagebox, Tk          # Para mostrar notificaciones emergentes
import os                                   # Para verificar si el archivo existe
import pytz                                 # Para manejar zonas horarias

# ===================== 📁 CARGAR ARCHIVO EXCEL =====================
archivo = "cumpleaños.xlsx"

# Si el archivo no existe, se detiene el programa
if not os.path.exists(archivo):
    print("No hay registros aún.")
    exit()

# Leer el archivo Excel
df = pd.read_excel(archivo)

# Asegurar que haya fechas válidas
df = df.dropna(subset=["Fecha de Nacimiento"])

# ===================== 🌎 ESTABLECER ZONA HORARIA =====================
zona = pytz.timezone('America/Lima')  # GMT-5
ahora_con_zona = datetime.now(zona)  # Fecha y hora actual con zona horaria
hoy = ahora_con_zona.date()          # Solo la fecha, sin hora

# ===================== 📋 LISTAS PARA GUARDAR RESULTADOS =====================
alerta_7_dias = []    # Personas que cumplen años en 7 días
alerta_hoy = []       # Personas que cumplen hoy
tabla_resumen = []    # Datos generales para mostrar en consola

# ===================== 🔁 RECORRER TODOS LOS REGISTROS =====================
for _, fila in df.iterrows():
    fecha_nac = fila["Fecha de Nacimiento"]

    # Convertir fecha si está como texto
    if isinstance(fecha_nac, str):
        fecha_nac = datetime.strptime(fecha_nac, "%Y-%m-%d").date()

    # Construir nombre completo
    nombre_completo = f"{fila['Nombre']} {fila['Apellido Paterno']} {fila['Apellido Materno']}"

    # Calcular la fecha del cumpleaños en el año actual
    cumple_este_anio = fecha_nac.replace(year=hoy.year)

    # Si ya pasó este año, considerar el próximo año
    if cumple_este_anio < hoy:
        cumple_este_anio = cumple_este_anio.replace(year=hoy.year + 1)

    # Cuántos días faltan para el cumpleaños
    dias_restantes = (cumple_este_anio - hoy).days

    # Edad que cumplirá
    edad = cumple_este_anio.year - fecha_nac.year

    # Agregar a la tabla de resumen
    tabla_resumen.append({
        "Nombre": nombre_completo,
        "Edad a cumplir": edad,
        "Días restantes": dias_restantes
    })

    # Verificar si se debe generar una alerta
    if dias_restantes == 7:
        alerta_7_dias.append(f"{nombre_completo} cumplirá {edad} años en 7 días.")
    elif dias_restantes == 0:
        alerta_hoy.append(f"¡Hoy es el cumpleaños de {nombre_completo}! 🎉 Cumple {edad} años.")

# ===================== 🖨️ MOSTRAR TABLA EN CONSOLA =====================
print("\n📋 Tabla de cumpleaños:")
print("{:<35} {:<15} {:<15}".format("Nombre", "Edad a cumplir", "Días restantes"))
print("-" * 65)
for persona in tabla_resumen:
    print(f"{persona['Nombre']:<35} {persona['Edad a cumplir']:<15} {persona['Días restantes']:<15}")

# ===================== 🔔 MOSTRAR ALERTAS EMERGENTES =====================
root = Tk()
root.withdraw()  # Oculta la ventana principal de Tkinter

# Mostrar notificaciones si aplica
if alerta_7_dias:
    messagebox.showinfo("⏰ Cumpleaños en 7 días", "\n".join(alerta_7_dias))

if alerta_hoy:
    messagebox.showinfo("🎉 Cumpleaños Hoy", "\n".join(alerta_hoy))

# Mensaje en consola si no hay cumpleaños relevantes
if not alerta_7_dias and not alerta_hoy:
    print("\nNo hay alertas de cumpleaños hoy ni en 7 días.")
