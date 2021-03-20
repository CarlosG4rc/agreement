from tkinter import *
from tkinter import filedialog as FileDialog
from io import open
import openpyxl
import docx

def crear():
   ss = openpyxl.load_workbook(ex.get() + ".xlsx")
   sheet = ss.get_sheet_by_name(hoja.get())
   
   for i in sheet.iter_rows(max_row=0):
      n = len(i)
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

      

root = Tk()
ex = StringVar()
hoja = StringVar()
root.title('Generación de contratos')

Label(root, text="Contratos", fg="darkblue", font=("Arial", 28, "bold")).pack()

Label(root, text="Nombre del Excel",fg="black",font=("Arial", 16, "bold")).pack()
Entry(root, justify="center", textvariable=ex).pack()

Label(root, text="Nombre de la hoja", fg="black", font=("Arial", 16, "bold")).pack()
Entry(root, justify="center", textvariable=hoja).pack()

Button(root, text="Aceptar", command=crear).pack()

root.mainloop()