from tkinter import *
from tkinter import filedialog as FileDialog
from io import open
import openpyxl
import docx
from docx.shared import RGBColor
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT


def crear():
   ss = openpyxl.load_workbook(ex.get() + ".xlsx")
   sheet = ss.get_sheet_by_name(hoja.get())
   
   for i in sheet.iter_rows(max_row=0):
      n = len(i)
   #Extraemos datos de contrato desde Excel
   for j in range(6, n):
      nombre = sheet.cell(row = j, column = 1).value
      nacionalidad = sheet.cell(row = j, column = 2).value
      domicilio = sheet.cell(row = j, column = 3).value
      curp = sheet.cell(row = j, column = 4).value
      rfc = sheet.cell(row = j, column = 5).value
      start = sheet.cell(row = j, column = 6).value
      final = sheet.cell(row = j, column = 7).value
      hrs = sheet.cell(row = j, column = 8).value
      importe = sheet.cell(row = j, column = 9).value
      importe_letra = sheet.cell(row = j, column = 10).value
      antiguedad = sheet.cell(row = j, column = 11).value

      #Creamos el docx
      doc = docx.Document()
      paragraph = doc.add_paragraph()
      run = paragraph.add_run("CONTRATO INDIVIDUAL DE TRABAJO POR TIEMPO DETERMINADO PARA PROFESORES")
      font = run.font
      paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
      font.color.rgb = RGBColor(0x00, 0x00, 0x00)
      font.name = 'Tahoma'
      font.size = Pt(11)
      font.bold = True
      
      doc.add_paragraph("CONTRATO INDIVIDUAL DE TRABAJO POR TIEMPO DETERMINADO PARA MAESTROS QUE CELEBRAN POR UNA PARTE EL COLEGIO.")
      doc.save("Contrato " + nombre + ".docx")

#Creamos interfas
root = Tk()
ex = StringVar()
hoja = StringVar()
root.title('Generación de contratos')

Label(root, text="Contratos", fg="darkblue", font=("Arial", 28, "bold")).pack()

#Pedimos datos necesarios de Excel
Label(root, text="Nombre del Excel",fg="black",font=("Arial", 16, "bold")).pack()
Entry(root, justify="center", textvariable=ex).pack()

Label(root, text="Nombre de la hoja", fg="black", font=("Arial", 16, "bold")).pack()
Entry(root, justify="center", textvariable=hoja).pack()

Button(root, text="Aceptar", command=crear).pack()

root.mainloop()