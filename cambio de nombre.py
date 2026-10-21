import pandas as pd
import os

def renombrar_archivos_python(archivo_excel, nombre_hoja='Hoja1'):
    # 1. Leer los datos de la hoja de Excel
    try:
        # pd.read_excel lee la tabla completa.
        # Nombra las columnas para facilitar el acceso
        df = pd.read_excel(archivo_excel, sheet_name=nombre_hoja, 
                           names=['ruta', 'nombre_antiguo', 'nombre_nuevo', 'estado'],
                           header=None, skiprows=1) # Asume que la primera fila es encabezado
    except FileNotFoundError:
        print(f"ERROR: No se encontró el archivo de Excel: {archivo_excel}")
        return

    # Inicializa la columna 'estado'
    df['estado'] = ""
    
    # 2. Iterar sobre las filas de la tabla (simulando el bucle Do While de VBA)
    for index, row in df.iterrows():
        # Construye las rutas completas
        ruta_completa_antigua = os.path.join(row['ruta'], row['nombre_antiguo'])
        ruta_completa_nueva = os.path.join(row['ruta'], row['nombre_nuevo'])
        
        # os.path.join maneja automáticamente la barra inclinada ("\")
        # que fue un punto de fallo en el código VBA.
        
        try:
            # 3. Comprobar si el archivo existe
            if os.path.exists(ruta_completa_antigua):
                
                # 4. Renombrar el archivo (similar al comando 'Name' de VBA)
                os.rename(ruta_completa_antigua, ruta_completa_nueva)
                df.loc[index, 'estado'] = "✅ Renombrado"
            else:
                df.loc[index, 'estado'] = "❌ No encontrado"
                
        except Exception as e:
            # Captura errores como 'archivo en uso' o 'permiso denegado'
            df.loc[index, 'estado'] = f"⚠️ Error: {e}"

    # 5. Guardar los resultados en el archivo de Excel (actualizando la Columna D)
    with pd.ExcelWriter(archivo_excel, mode='a', if_sheet_exists='replace') as writer:
        df.to_excel(writer, sheet_name=nombre_hoja, index=False, header=False, startrow=1)

    print("Proceso terminado. Se actualizó la columna de estado en el archivo de Excel.")

# --- Ejecución ---
nombre_del_archivo_de_datos = 'datos_renombrado.xlsx'
renombrar_archivos_python(nombre_del_archivo_de_datos)