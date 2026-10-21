import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox

# --- Lógica de Renombrado (Misma función, adaptada para la GUI) ---

def renombrar_archivos_desde_excel(archivo_excel, nombre_hoja, log_text):
    """Lógica principal para renombrar archivos y actualizar el estado."""
    log_text.set("Iniciando proceso...")

    # 1. Leer los datos
    try:
        # Lee el archivo, asumiendo datos a partir de la Fila 2
        df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja, 
                           names=['ruta', 'nombre_antiguo', 'nombre_nuevo', 'estado'],
                           header=None, skiprows=1)
    except Exception as e:
        log_text.set(f"ERROR: No se pudo leer el archivo. Asegúrese de que no esté abierto. Detalle: {e}")
        return

    if df.empty:
        log_text.set("Advertencia: La hoja de cálculo está vacía o no se encontraron datos.")
        return

    if 'estado' not in df.columns:
        df['estado'] = ""
    
    # 2. Iterar y Renombrar
    archivos_renombrados = 0
    
    for index, row in df.iterrows():
        # Limpiar y convertir a string
        ruta_base = str(row['ruta']).strip()
        nombre_antiguo = str(row['nombre_antiguo']).strip()
        nombre_nuevo = str(row['nombre_nuevo']).strip()
        
        # Construir rutas
        ruta_completa_antigua = os.path.join(ruta_base, nombre_antiguo)
        ruta_completa_nueva = os.path.join(ruta_base, nombre_nuevo)
        
        if ruta_base and nombre_antiguo and nombre_nuevo:
            try:
                if os.path.exists(ruta_completa_antigua):
                    os.rename(ruta_completa_antigua, ruta_completa_nueva)
                    df.loc[index, 'estado'] = "Renombrado OK"
                    archivos_renombrados += 1
                else:
                    df.loc[index, 'estado'] = "No encontrado"
                    
            except Exception as e:
                df.loc[index, 'estado'] = f"Error: {e}"
        else:
             df.loc[index, 'estado'] = "Datos incompletos"

    # 3. Guardar los resultados
    try:
        with pd.ExcelWriter(archivo_excel, mode='a', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name=nombre_hoja, index=False, header=False, startrow=1)

        log_text.set(f"PROCESO TERMINADO. Renombrados: {archivos_renombrados}. Archivo Excel actualizado.")
        messagebox.showinfo("Éxito", f"¡Proceso terminado!\nSe renombraron {archivos_renombrados} archivos y se actualizó el Excel.")
    except Exception as e:
        log_text.set(f"ERROR crítico al guardar el Excel. Asegúrese de que el archivo esté CERRADO. Detalle: {e}")
        messagebox.showerror("Error de Guardado", "ERROR: Asegúrese de que el archivo de Excel esté CERRADO.")

# --- Funciones de la Interfaz Gráfica (Tkinter) ---

def seleccionar_archivo(entry_widget):
    """Abre un diálogo para seleccionar el archivo Excel."""
    filepath = filedialog.askopenfilename(
        defaultextension=".xlsx",
        filetypes=[("Archivos de Excel", "*.xlsx"), ("Archivos de Excel antiguos", "*.xls")]
    )
    if filepath:
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, filepath)

def ejecutar_proceso(entry_ruta, entry_hoja, log_text):
    """Valida los datos de entrada y ejecuta la función principal."""
    ruta_excel = entry_ruta.get()
    nombre_hoja = entry_hoja.get()

    if not os.path.exists(ruta_excel):
        messagebox.showerror("Error de Validación", "La ruta del archivo de Excel no es válida.")
        return
    
    if not nombre_hoja:
        messagebox.showerror("Error de Validación", "Debe ingresar el nombre de la hoja de cálculo.")
        return

    # Llamada a la función de renombrado
    renombrar_archivos_desde_excel(ruta_excel, nombre_hoja, log_text)

def crear_gui():
    """Crea y muestra la ventana principal."""
    root = tk.Tk()
    root.title("Renombrador de Archivos por Lote")
    root.geometry("600x300")
    root.resizable(False, False)

    # Variables de control para los campos de texto
    ruta_excel_var = tk.StringVar()
    nombre_hoja_var = tk.StringVar(value="Hoja1") # Valor predeterminado
    log_var = tk.StringVar(value="Esperando ejecución...")

    # --- Marco principal ---
    main_frame = tk.Frame(root, padx=10, pady=10)
    main_frame.pack(fill='both', expand=True)

    # --- Fila 1: Ruta del archivo Excel ---
    tk.Label(main_frame, text="Ruta del Archivo Excel:", anchor="w").grid(row=0, column=0, padx=5, pady=5, sticky="w")
    entry_ruta = tk.Entry(main_frame, textvariable=ruta_excel_var, width=50)
    entry_ruta.grid(row=0, column=1, padx=5, pady=5)
    btn_buscar = tk.Button(main_frame, text="Buscar...", command=lambda: seleccionar_archivo(entry_ruta))
    btn_buscar.grid(row=0, column=2, padx=5, pady=5)

    # --- Fila 2: Nombre de la Hoja ---
    tk.Label(main_frame, text="Nombre de la Hoja:", anchor="w").grid(row=1, column=0, padx=5, pady=5, sticky="w")
    entry_hoja = tk.Entry(main_frame, textvariable=nombre_hoja_var, width=50)
    entry_hoja.grid(row=1, column=1, padx=5, pady=5)
    
    # --- Separador ---
    tk.Frame(main_frame, height=1, bg='gray').grid(row=2, column=0, columnspan=3, sticky="ew", pady=10)

    # --- Fila 3: Botón de Ejecutar ---
    btn_ejecutar = tk.Button(main_frame, text="  🚀  EJECUTAR RENOMBRADO  ", 
                             command=lambda: ejecutar_proceso(entry_ruta, entry_hoja, log_var), 
                             bg='lightblue', font=('Arial', 10, 'bold'))
    btn_ejecutar.grid(row=3, column=0, columnspan=3, pady=10)

    # --- Fila 4: Log de Estado ---
    tk.Label(main_frame, text="Estado:", anchor="w").grid(row=4, column=0, padx=5, pady=5, sticky="w")
    tk.Label(main_frame, textvariable=log_var, fg='blue', wraplength=450, justify='left').grid(row=4, column=1, columnspan=2, padx=5, pady=5, sticky="w")
    
    root.mainloop()

# --- Inicio del Programa ---
if __name__ == "__main__":
    crear_gui()