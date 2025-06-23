# Proyecto: Excel Cleaner & Logger
# Descripción: Limpia archivos Excel/CSV y genera logs separados para el cliente y para uso interno.

import os
import time
import pandas as pd
from datetime import datetime
from tkinter import *
from tkinter import filedialog, messagebox

# Variables globales
selected_files = []
output_folder = ""

# ---------- FUNCIONES PRINCIPALES ----------

def clean_excel_file(file_path, output_folder, client_log, internal_log):
    start_time = time.time()
    filename = os.path.basename(file_path)

    try:
        if file_path.endswith(".csv"):
            df = pd.read_csv(file_path, dtype=str)
        elif file_path.endswith(".xlsx"):
            df = pd.read_excel(file_path, dtype=str)
        else:
            raise ValueError("Unsupported file format")

        original_shape = df.shape
        original_cells = df.size

        # Eliminar filas y columnas vacías
        df.dropna(how='all', inplace=True)
        df.dropna(axis=1, how='all', inplace=True)

        # Eliminar duplicados
        df.drop_duplicates(inplace=True)

        # Limpiar nombres de columnas
        df.columns = [col.strip().lower().replace(' ', '_').replace('$', '').replace('-', '_') for col in df.columns]

        # Limpiar espacios en celdas de texto
        for col in df.select_dtypes(include='object').columns:
            df[col] = df[col].astype(str).str.strip()

        cleaned_shape = df.shape
        cleaned_cells = df.size

        # Guardar archivo limpio
        output_path = os.path.join(output_folder, filename)
        if file_path.endswith(".csv"):
            df.to_csv(output_path, index=False)
        else:
            df.to_excel(output_path, index=False)

        end_time = time.time()
        duration = round(end_time - start_time, 2)

        # Log para el cliente
        client_log.append(f"Archivo: {filename}\nFilas originales: {original_shape[0]}, Finales: {cleaned_shape[0]}\n"
                          f"Columnas originales: {original_shape[1]}, Finales: {cleaned_shape[1]}\nCeldas antes: {original_cells}, Después: {cleaned_cells}\n\n")

        # Log interno para el freelancer
        internal_log.append(f"{datetime.now()} - {filename} procesado en {duration} segundos\n")

        print(f"✔️ {filename} limpiado correctamente.")
    except Exception as e:
        print(f"❌ Error con {filename}: {e}")

# ---------- FUNCIONES DE INTERFAZ ----------

def select_files():
    global selected_files
    selected_files = filedialog.askopenfilenames(filetypes=[("Excel/CSV Files", "*.csv *.xlsx")])
    label_files.config(text=f"{len(selected_files)} archivo(s) seleccionado(s)")

def select_output_folder():
    global output_folder
    output_folder = filedialog.askdirectory()
    label_output.config(text=f"Destino: {output_folder}")

def process_files():
    if not selected_files or not output_folder:
        messagebox.showwarning("Faltan datos", "Selecciona archivos y carpeta de destino.")
        return

    client_log = []
    internal_log = []

    for file in selected_files:
        clean_excel_file(file, output_folder, client_log, internal_log)

    # Guardar logs
    with open(os.path.join(output_folder, "log_cliente.txt"), "w", encoding="utf-8") as f:
        f.writelines(client_log)

    with open(os.path.join(output_folder, "log_freelancer.txt"), "w", encoding="utf-8") as f:
        f.writelines(internal_log)

    messagebox.showinfo("Finalizado", "Todos los archivos han sido limpiados y los logs generados.")

# ---------- INTERFAZ GRÁFICA ----------

root = Tk()
root.title("Excel Cleaner & Logger")
root.geometry("500x300")
root.config(padx=20, pady=20)

Label(root, text="Selecciona archivos Excel o CSV para limpiar", font=("Arial", 12)).pack(pady=10)
Button(root, text="1. Seleccionar archivos", command=select_files).pack(pady=5)
label_files = Label(root, text="Ningún archivo seleccionado", fg="gray")
label_files.pack()

Button(root, text="2. Seleccionar carpeta de destino", command=select_output_folder).pack(pady=5)
label_output = Label(root, text="Carpeta no seleccionada", fg="gray")
label_output.pack()

Button(root, text="3. Procesar y generar logs", command=process_files, bg="green", fg="white").pack(pady=20)

root.mainloop()
