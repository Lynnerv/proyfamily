import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
from datetime import datetime
import pytz
import os

# === 📂 Verificar archivo Excel ===
archivo_excel = "cumpleaños.xlsx"
if not os.path.exists(archivo_excel):
    messagebox.showerror("Error", "El archivo de cumpleaños no existe.")
    exit()

# === 🕒 Fecha actual con zona horaria Perú ===
zona = pytz.timezone("America/Lima")
hoy = datetime.now(zona).date()

# === 📚 Cargar datos y calcular edades y días restantes ===
df = pd.read_excel(archivo_excel)
df = df.dropna(subset=["Fecha de Nacimiento"])

datos_tabla = []

for _, fila in df.iterrows():
    fecha_nac = fila["Fecha de Nacimiento"]
    if isinstance(fecha_nac, str):
        fecha_nac = datetime.strptime(fecha_nac, "%Y-%m-%d").date()

    cumple_este_anio = fecha_nac.replace(year=hoy.year)
    if cumple_este_anio < hoy:
        cumple_este_anio = cumple_este_anio.replace(year=hoy.year + 1)

    dias_restantes = (cumple_este_anio - hoy).days
    edad = cumple_este_anio.year - fecha_nac.year
    nombre_completo = f"{fila['Nombre']} {fila['Apellido Paterno']} {fila['Apellido Materno']}"

    datos_tabla.append((nombre_completo, edad, dias_restantes))

# === 🎨 Crear ventana y aplicar estilo oscuro ===
ventana = tk.Tk()
ventana.title("🎂 Cumpleaños Registrados")
ventana.geometry("720x420")
ventana.configure(bg="#1f1f1f")

# 🎨 Colores y fuentes
fondo_color = "#1f1f1f"
texto_color = "#f5f5f5"
boton_color = "#d4af37"
hover_color = "#f0c846"
fuente_general = ("Segoe UI", 12)
fuente_heading = ("Segoe UI", 13, "bold")

# === Estilos de tabla ===
style = ttk.Style()
style.theme_use("clam")
style.configure("Treeview",
    background="#2c2c2c",
    fieldbackground="#2c2c2c",
    foreground=texto_color,
    font=fuente_general,
    rowheight=28
)
style.configure("Treeview.Heading",
    background=boton_color,
    foreground="black",
    font=fuente_heading
)
style.map("Treeview",
    background=[("selected", "#444444")],
    foreground=[("selected", "#ffffff")]
)

# === Tabla Treeview ===
tree = ttk.Treeview(ventana, columns=("Nombre", "Edad", "Días"), show="headings")
tree.heading("Nombre", text="👤 Nombre")
tree.heading("Edad", text="🎂 Edad que cumplirá")
tree.heading("Días", text="📅 Días restantes")

tree.column("Nombre", width=350)
tree.column("Edad", width=150, anchor="center")
tree.column("Días", width=170, anchor="center")

for persona in datos_tabla:
    tree.insert("", tk.END, values=persona)

tree.pack(padx=20, pady=20, fill="both", expand=True)

# === Botón Cerrar ===
style.configure("TButton", font=fuente_general, background=boton_color, foreground="black", borderwidth=0)
style.map("TButton", background=[("active", hover_color)])
ttk.Button(ventana, text="Cerrar", command=ventana.destroy).pack(pady=10)

ventana.mainloop()
