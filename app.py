import openpyxl as op
from datetime import datetime
from docxtpl import DocxTemplate

input_file = "Notas de alumnos.xlsx"
input_latter = DocxTemplate("Carta.docx")

#Activamos el libro y la hoja de excel para trabajar con el
try:
    wordbook = op.load_workbook(input_file)
    sheet_input = wordbook.active
except FileNotFoundError:
    print(f"ERROR: El archivo {input_file} no se encontro") 

for indexrow, row in enumerate(sheet_input.iter_rows(min_row=2), start=2):
    """Aqui estamos opteniendo los valores de las columnas, Teniendo en cuenta que las primeras filas
    son los encabezados ósea, Nombre, Apellido. etc"""
    nombre = row[0].value
    apellido = row[1].value
    nota = row[2].value
    materia = row[3].value
    
    if nombre and apellido:
        print(f'Procesando los datos de {nombre} {apellido}')
        
    Valores = {
        'Nombre': nombre,
        'Apellido': apellido,
        'Nota': nota,
        'Materia': materia
    }
    
    input_latter.render(Valores)
    
    