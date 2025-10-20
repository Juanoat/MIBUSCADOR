import pandas as pd
import glob
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk

def buscar_en_archivos():
    # Seleccionar carpeta
    carpeta = filedialog.askdirectory(title="Selecciona la carpeta con los archivos Excel")
    if not carpeta:
        return
    
    # Pedir qué buscar
    buscar = entrada_busqueda.get()
    if not buscar:
        messagebox.showwarning("Advertencia", "Debes ingresar algo para buscar")
        return
    
    resultados = []
    archivos = glob.glob(f"{carpeta}/**/*.xlsx", recursive=True)
    
    if not archivos:
        messagebox.showinfo("Sin archivos", "No se encontraron archivos Excel en esa carpeta")
        return
    
    progreso['maximum'] = len(archivos)
    
    for i, archivo in enumerate(archivos):
        try:
            label_estado.config(text=f"Procesando: {Path(archivo).name}")
            ventana.update()
            
            df = pd.read_excel(archivo, sheet_name=None)
            for hoja, datos in df.items():
                mascara = datos.apply(
                    lambda row: row.astype(str).str.contains(buscar, case=False).any(), 
                    axis=1
                )
                if mascara.any():
                    for fila in mascara[mascara].index.tolist():
                        resultados.append({
                            'archivo': Path(archivo).name,
                            'ruta_completa': archivo,
                            'hoja': hoja,
                            'fila': fila + 2
                        })
            progreso['value'] = i + 1
            ventana.update()
        except Exception as e:
            print(f"Error en {archivo}: {e}")
    
    # Guardar resultados
    if resultados:
        salida = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile="resultados.xlsx"
        )
        if salida:
            pd.DataFrame(resultados).to_excel(salida, index=False)
            messagebox.showinfo("Éxito", f"Se encontraron {len(resultados)} coincidencias!\n\nArchivo guardado en:\n{salida}")
    else:
        messagebox.showinfo("Sin resultados", "No se encontraron coincidencias")
    
    label_estado.config(text="Listo!")
    progreso['value'] = 0

# Crear ventana
ventana = tk.Tk()
ventana.title("Buscador en Archivos Excel")
ventana.geometry("500x280")
ventana.resizable(False, False)

# Título
tk.Label(ventana, text="🔍 BUSCADOR DE ARCHIVOS EXCEL", 
         font=("Arial", 14, "bold"), fg="#2c3e50").pack(pady=15)

# Campo de búsqueda
tk.Label(ventana, text="¿Qué querés buscar?", font=("Arial", 10)).pack(pady=5)
entrada_busqueda = tk.Entry(ventana, width=40, font=("Arial", 11))
entrada_busqueda.pack(pady=5)

# Botón buscar
tk.Button(ventana, text="🔎 BUSCAR EN ARCHIVOS", command=buscar_en_archivos, 
          bg="#27ae60", fg="white", font=("Arial", 12, "bold"), 
          padx=20, pady=10, cursor="hand2").pack(pady=15)

# Estado
label_estado = tk.Label(ventana, text="", font=("Arial", 9), fg="#7f8c8d")
label_estado.pack(pady=5)

# Barra de progreso
progreso = ttk.Progressbar(ventana, length=400, mode='determinate')
progreso.pack(pady=10)

# Info
tk.Label(ventana, text="Creado por el Mago", font=("Arial", 8), fg="#95a5a6").pack(side="bottom", pady=5)

ventana.mainloop()