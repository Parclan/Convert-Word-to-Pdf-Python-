import os
import win32com.client

# Ruta al archivo de Word que queremos convertir a PDF
word_file_path = "C:/Users/DELL/Documents/sonia_proyecto_de_vida.doc"

# Ruta donde queremos guardar el archivo PDF resultante
pdf_file_path = "C:/Users/DELL/Documents/sonia.pdf"

# Creamos una instancia de Word
word = win32com.client.Dispatch('Word.Application')

try:
    # Abrimos el archivo de Word
    doc = word.Documents.Open(word_file_path)

    # Guardamos el archivo en formato PDF
    doc.SaveAs(pdf_file_path, FileFormat=17)

    # Cerramos el archivo de Word
    doc.Close()

    print(f"Archivo convertido a PDF: {pdf_file_path}")

except Exception as e:
    print(f"Error al convertir el archivo: {e}")

finally:
    # Cerramos la aplicación de Word
    word.Quit()
