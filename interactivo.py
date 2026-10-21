import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox

# --- LIBRERÍA ESPECÍFICA DE WINDOWS PARA OUTLOOK ---
try:
    import win32com.client as win32
except ImportError:
    messagebox.showerror("Error de Librería", 
                         "La librería 'pywin32' (win32com.client) no está instalada o no es Windows. "
                         "Instale con: pip install pywin32")
    exit()
# ---------------------------------------------------

def enviar_correos_desde_excel(config, log_text):
    """Lógica principal: lee Excel y envía/muestra correos vía Outlook COM."""
    
    # 1. Obtener Configuración de la GUI
    try:
        ruta_excel = config['ruta_excel'].get()
        nombre_hoja = config['nombre_hoja'].get()
        ruta_adjuntos = config['ruta_adjuntos'].get()
        extension = config['extension'].get().strip().lower()
        col_correo = config['col_correo'].get().strip()
        col_nombre = config['col_nombre'].get().strip()
        col_asunto = config['col_asunto'].get().strip()
        col_mensaje = config['col_mensaje'].get().strip()
        modo_envio = config['modo_envio'].get() # 1=Enviar, 0=Mostrar
    except Exception as e:
        log_text.set(f"Error de Configuración: {e}")
        return

    # Validar ruta de Excel
    if not os.path.exists(ruta_excel):
        log_text.set("ERROR: Ruta de Excel no válida.")
        return

    # Validar ruta de Adjuntos
    if not os.path.exists(ruta_adjuntos):
        log_text.set("ERROR: Ruta de Adjuntos no válida.")
        return
        
    # Asegurar que la ruta termine con "\" para construcción (os.path.join es mejor, pero aquí lo forzamos)
    if not ruta_adjuntos.endswith(os.sep):
        ruta_adjuntos += os.sep

    log_text.set("Iniciando proceso: Leyendo Excel...")

    # 2. Leer los datos
    try:
        df = pd.read_excel(ruta_excel, sheet_name=nombre_hoja)
    except Exception as e:
        log_text.set(f"ERROR al leer Excel: Asegúrese de que no esté abierto. Detalle: {e}")
        return

    # Validar que todas las columnas existan
    required_cols = [col_correo, col_nombre, col_asunto, col_mensaje]
    if not all(col in df.columns for col in required_cols):
        log_text.set("ERROR: Una o más columnas (Correo, Nombre, Asunto, Mensaje) no se encontraron en el Excel.")
        return

    # 3. Inicializar Outlook (similar al CreateObject de VBA)
    try:
        outlook = win32.Dispatch('Outlook.Application')
        # objNamespace = outlook.GetNamespace('MAPI') # Usualmente no es necesario para enviar
        log_text.set("Conexión con Outlook exitosa. Procesando correos...")
    except Exception as e:
        log_text.set(f"ERROR: No se pudo conectar con Outlook. Asegúrese de que esté instalado. Detalle: {e}")
        return

    # 4. Bucle de Envío
    correos_procesados = 0
    df['Estado_Envio'] = ""

    for index, row in df.iterrows():
        try:
            destinatario = str(row[col_correo]).strip()
            nombre_archivo = str(row[col_nombre]).strip()
            asunto = str(row[col_asunto]).strip()
            mensaje = str(row[col_mensaje])
            
            # 5. Construir Ruta de Adjunto
            adjunto_path = os.path.join(ruta_adjuntos, f"{nombre_archivo}.{extension}")
            
            if pd.isna(destinatario) or destinatario == "":
                 df.loc[index, 'Estado_Envio'] = "ERROR: Destinatario vacío"
                 continue

            if not os.path.exists(adjunto_path):
                df.loc[index, 'Estado_Envio'] = f"ERROR: Adjunto no encontrado: {os.path.basename(adjunto_path)}"
                continue

            # 6. Crear el Correo (CreateItem(0) en VBA)
            mail = outlook.CreateItem(0)
            
            # Usar HTMLBody para un formateo más limpio (reemplazando saltos de línea de texto)
            html_body = f"""<html><body>
                            {mensaje.replace(os.linesep, '<br>').replace('\n', '<br>')}
                            </body></html>"""

            mail.To = destinatario
            mail.Subject = asunto
            mail.HTMLBody = html_body 
            mail.Attachments.Add(adjunto_path)

            if modo_envio == 1:
                mail.Send()
                df.loc[index, 'Estado_Envio'] = "ENVIADO OK"
            else:
                mail.Display(False) # False para no ser modal
                df.loc[index, 'Estado_Envio'] = "GENERADO para Revisión"
                
            correos_procesados += 1

        except Exception as e:
            df.loc[index, 'Estado_Envio'] = f"ERROR de Outlook/Sistema: {e}"

    # 7. Guardar los resultados en el Excel
    try:
        with pd.ExcelWriter(ruta_excel, mode='a', if_sheet_exists='replace', engine='openpyxl') as writer:
            # Asegúrate de no incluir el índice de Pandas al guardar
            df.to_excel(writer, sheet_name=nombre_hoja, index=False) 

        log_text.set(f"PROCESO TERMINADO. Procesados: {len(df)}. Enviados/Generados: {correos_procesados}. Excel actualizado.")
        messagebox.showinfo("Éxito", f"¡Proceso terminado!\n{correos_procesados} correos procesados.\nRevise la nueva columna 'Estado_Envio' en su Excel.")
    except Exception as e:
        log_text.set(f"ERROR crítico al guardar el Excel. Ciérrelo. Detalle: {e}")
        messagebox.showerror("Error de Guardado", "ERROR: Asegúrese de que el archivo de Excel esté CERRADO y vuelva a ejecutar.")

# --- Funciones de la Interfaz Gráfica (Tkinter) ---

def seleccionar_archivo(entry_widget):
    """Abre diálogo para seleccionar el archivo Excel."""
    filepath = filedialog.askopenfilename(
        defaultextension=".xlsx",
        filetypes=[("Archivos de Excel", "*.xlsx")]
    )
    if filepath:
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, filepath)

def seleccionar_carpeta(entry_widget):
    """Abre diálogo para seleccionar la carpeta de adjuntos."""
    folderpath = filedialog.askdirectory()
    if folderpath:
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, folderpath)

def crear_gui():
    """Crea y muestra la ventana principal."""
    root = tk.Tk()
    root.title("📧 Envío Masivo por Lote (Python/Outlook)")
    root.geometry("750x600")

    # Variables de control
    config = {
        'ruta_excel': tk.StringVar(),
        'nombre_hoja': tk.StringVar(value="Sheet1"),
        'ruta_adjuntos': tk.StringVar(),
        'extension': tk.StringVar(value="pdf"),
        'col_correo': tk.StringVar(),
        'col_nombre': tk.StringVar(),
        'col_asunto': tk.StringVar(),
        'col_mensaje': tk.StringVar(),
        'modo_envio': tk.IntVar(value=0) # 0=Display, 1=Send
    }
    log_var = tk.StringVar(value="Esperando configuración e inicio...")

    # --- Marco principal ---
    main_frame = tk.Frame(root, padx=10, pady=10)
    main_frame.pack(fill='both', expand=True)

    # --- Título ---
    tk.Label(main_frame, text="Configuración del Envío Masivo", font=('Arial', 14, 'bold')).grid(row=0, columnspan=3, pady=10)
    
    # --- SECCIÓN DE ARCHIVOS Y RUTAS ---
    row_counter = 1
    tk.Label(main_frame, text="1. Archivo de Datos y Adjuntos", font=('Arial', 10, 'bold')).grid(row=row_counter, column=0, columnspan=3, pady=5, sticky="w")
    
    # Ruta Excel
    tk.Label(main_frame, text="Ruta Excel:").grid(row=row_counter + 1, column=0, padx=5, pady=2, sticky="w")
    entry_ruta = tk.Entry(main_frame, textvariable=config['ruta_excel'], width=60)
    entry_ruta.grid(row=row_counter + 1, column=1, padx=5, pady=2)
    tk.Button(main_frame, text="Buscar...", command=lambda: seleccionar_archivo(entry_ruta)).grid(row=row_counter + 1, column=2, padx=5, pady=2)

    # Hoja y Extensión
    tk.Label(main_frame, text="Hoja:").grid(row=row_counter + 2, column=0, padx=5, pady=2, sticky="w")
    tk.Entry(main_frame, textvariable=config['nombre_hoja'], width=10).grid(row=row_counter + 2, column=1, padx=5, pady=2, sticky="w")
    
    tk.Label(main_frame, text="Extensión Adjunto:").grid(row=row_counter + 2, column=1, padx=5, pady=2, sticky="e")
    tk.Entry(main_frame, textvariable=config['extension'], width=10).grid(row=row_counter + 2, column=2, padx=5, pady=2, sticky="w")

    # Carpeta Adjuntos
    tk.Label(main_frame, text="Carpeta Adjuntos:").grid(row=row_counter + 3, column=0, padx=5, pady=2, sticky="w")
    entry_folder = tk.Entry(main_frame, textvariable=config['ruta_adjuntos'], width=60)
    entry_folder.grid(row=row_counter + 3, column=1, padx=5, pady=2)
    tk.Button(main_frame, text="Buscar...", command=lambda: seleccionar_carpeta(entry_folder)).grid(row=row_counter + 3, column=2, padx=5, pady=2)

    # --- SECCIÓN DE COLUMNAS ---
    row_counter += 4
    tk.Frame(main_frame, height=1, bg='gray').grid(row=row_counter, column=0, columnspan=3, sticky="ew", pady=5)
    row_counter += 1
    tk.Label(main_frame, text="2. Columnas en el Excel (Ingrese el nombre de la columna)", font=('Arial', 10, 'bold')).grid(row=row_counter, column=0, columnspan=3, pady=5, sticky="w")

    tk.Label(main_frame, text="Correo Destinatario:").grid(row=row_counter + 1, column=0, padx=5, pady=2, sticky="w")
    tk.Entry(main_frame, textvariable=config['col_correo'], width=20).grid(row=row_counter + 1, column=1, padx=5, pady=2, sticky="w")
    
    tk.Label(main_frame, text="Nombre del Archivo:").grid(row=row_counter + 2, column=0, padx=5, pady=2, sticky="w")
    tk.Entry(main_frame, textvariable=config['col_nombre'], width=20).grid(row=row_counter + 2, column=1, padx=5, pady=2, sticky="w")
    
    tk.Label(main_frame, text="Asunto del Correo:").grid(row=row_counter + 1, column=1, padx=5, pady=2, sticky="e")
    tk.Entry(main_frame, textvariable=config['col_asunto'], width=20).grid(row=row_counter + 1, column=2, padx=5, pady=2, sticky="w")

    tk.Label(main_frame, text="Cuerpo del Mensaje:").grid(row=row_counter + 2, column=1, padx=5, pady=2, sticky="e")
    tk.Entry(main_frame, textvariable=config['col_mensaje'], width=20).grid(row=row_counter + 2, column=2, padx=5, pady=2, sticky="w")

    # --- SECCIÓN DE MODO DE ENVÍO ---
    row_counter += 3
    tk.Frame(main_frame, height=1, bg='gray').grid(row=row_counter, column=0, columnspan=3, sticky="ew", pady=5)
    row_counter += 1
    tk.Label(main_frame, text="3. Modo de Ejecución", font=('Arial', 10, 'bold')).grid(row=row_counter, column=0, columnspan=3, pady=5, sticky="w")

    tk.Radiobutton(main_frame, text="Solo GENERAR (Mostrar en Outlook, no enviar)", variable=config['modo_envio'], value=0).grid(row=row_counter + 1, column=0, columnspan=3, sticky="w")
    tk.Radiobutton(main_frame, text="ENVIAR Directamente (Usar con precaución)", variable=config['modo_envio'], value=1).grid(row=row_counter + 2, column=0, columnspan=3, sticky="w")

    # --- BOTÓN DE EJECUTAR ---
    row_counter += 3
    tk.Button(main_frame, text=" 📤 INICIAR ENVÍO MASIVO", 
              command=lambda: enviar_correos_desde_excel(config, log_var), 
              bg='#4CAF50', fg='white', font=('Arial', 12, 'bold')).grid(row=row_counter, column=0, columnspan=3, pady=15)

    # --- LOG DE ESTADO ---
    row_counter += 1
    tk.Label(main_frame, text="Estado:", anchor="w", font=('Arial', 10, 'bold')).grid(row=row_counter, column=0, padx=5, pady=5, sticky="w")
    tk.Label(main_frame, textvariable=log_var, fg='blue', wraplength=600, justify='left').grid(row=row_counter, column=1, columnspan=2, padx=5, pady=5, sticky="w")
    
    root.mainloop()

# --- Inicio del Programa ---
if __name__ == "__main__":
    crear_gui()