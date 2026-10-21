from PyPDF2 import PdfReader, PdfWriter
import os

# Ruta del archivo PDF original
ruta_pdf = r"C:\Users\JUAN DAVID\Desktop\Nueva carpeta\ica_noviembre\ICA_NOVIEMBRE.pdf"

# Carpeta de salida (puedes cambiarla si quieres)
carpeta_salida = r"C:\Users\JUAN DAVID\Desktop\Nueva carpeta\ica_noviembre\ica"

# Crear carpeta de salida si no existe
os.makedirs(carpeta_salida, exist_ok=True)

# Leer el PDF
reader = PdfReader(ruta_pdf)

# Dividir el PDF en páginas individuales
for i, pagina in enumerate(reader.pages, start=1):
    writer = PdfWriter()
    writer.add_page(pagina)

    nombre_salida = f"hoja_{i}.pdf"
    ruta_salida = os.path.join(carpeta_salida, nombre_salida)

    with open(ruta_salida, "wb") as salida:
        writer.write(salida)

    print(f" Página {i} guardada como {nombre_salida}")

print("\n Proceso completado. Todos los archivos están en:")
print(carpeta_salida)