import tkinter as tk
from tkinter import ttk
import subprocess
import os

# ===== 🎨 ESTILOS Y COLORES COMPARTIDOS =====
fondo_color = "#1f1f1f"
boton_color = "#d4af37"
hover_color = "#f0c846"
texto_color = "#f5f5f5"
fuente_general = ("Segoe UI", 12, "bold")

# ===== 🌐 FUNCIONES PARA EJECUTAR LOS MÓDULOS =====
def abrir_registro():
    subprocess.Popen(["python", "registro_gui.py"], shell=True)

def abrir_tabla():
    subprocess.Popen(["python", "ver_tabla_gui.py"], shell=True)

def abrir_recordatorio():
    subprocess.Popen(["python", "recordatorio.py"], shell=True)

# ===== 🖼️ INTERFAZ PRINCIPAL =====
ventana = tk.Tk()
ventana.title("🎉 Menú de Cumpleaños")
ventana.geometry("400x400")
ventana.configure(bg=fondo_color)

# ===== 🧭 ESTILO =====
style = ttk.Style()
style.theme_use("clam")
style.configure("TButton", font=fuente_general, background=boton_color, foreground="black", borderwidth=0)
style.map("TButton", background=[("active", hover_color)])

# ===== 🎯 TÍTULO =====
titulo = tk.Label(ventana, text="🎂 Sistema de Cumpleaños", bg=fondo_color, fg=texto_color, font=("Segoe UI", 15, "bold"))
titulo.pack(pady=30)

# ===== 🧭 BOTONES DEL MENÚ =====
ttk.Button(ventana, text="📝 Registrar Cumpleaños", command=abrir_registro).pack(pady=10, ipadx=10, ipady=5)
ttk.Button(ventana, text="📊 Ver Tabla Cumpleaños", command=abrir_tabla).pack(pady=10, ipadx=10, ipady=5)
ttk.Button(ventana, text="🔔 Ejecutar Recordatorio", command=abrir_recordatorio).pack(pady=10, ipadx=10, ipady=5)
ttk.Button(ventana, text="❌ Salir", command=ventana.destroy).pack(pady=30, ipadx=10, ipady=5)

ventana.mainloop()
