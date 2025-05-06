import openpyxl as op
from datetime import datetime
from docxtpl import DocxTemplate

#Abrimos el archivo de excel y el documento de word que vamos a usar como plantilla 
input_file = "Notas de alumnos.xlsx"
input_latter = DocxTemplate("Carta.docx")

#Activamos el libro y la hoja de excel para trabajar con el
try:
    wordbook_input = op.load_workbook(input_file)
    sheet_input = wordbook_input.active
except FileNotFoundError:
    print(f"ERROR: El archivo {input_file} no se encontro") 


for indexrow, row in enumerate(sheet_input.iter_rows(min_row=2), start=2):
    """Aqui estamos opteniendo los valores de las columnas, Teniendo en cuenta que las primeras filas
    son los encabezados ósea, Nombre, Apellido. etc"""
    fecha = datetime.now().strftime("%d/%m/%Y")
    nombre = row[0].value
    apellido = row[1].value
    nota = row[2].value
    materia = row[3].value
    
    if nombre and apellido:
        print(f'Procesando los datos de {nombre} {apellido}')
    
    # Creamos un diccionario con los valores que vamos a reemplazar en el documento
    # y los valores que vamos a usar para el nombre del archivo
    # La clave del diccionario es el nombre de la variable que se encuentra en el documento
    Valores = {
        'Fecha': fecha,
        'Nombre': nombre,
        'Apellido': apellido,
        'Nota': nota,
        'Materia': materia
    }
    
    input_latter.render(Valores)
    
    # Guardar un archivo .docx para cada alumno
    output_filename = f"{nombre}_{apellido}.docx"
    input_latter.save(output_filename)
    print(f"Archivo generado: {output_filename}")

