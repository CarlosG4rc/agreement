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

      paragraph = doc.add_paragraph()
      run = paragraph.add_run("CONTRATO INDIVIDUAL DE TRABAJO POR TIEMPO DETERMINADO PARA MAESTROS QUE CELEBRAN POR UNA PARTE EL COLEGIO ")
      font = run.font
      paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)
      
      run1 = paragraph.add_run("INSTITUTO FRANCISCO POSSENTI, A. C. ")
      run1.bold = TRUE
      font = run1.font
      paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      run = paragraph.add_run("CON DOMICILIO EN AV. TOLUCA No. 621 COL. OLIVAR DE LOS PADRES DEL. ALVARO OBREGON C. P. 01780 REPRESENTADO POR EL C. ")
      font = run.font
      paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      run1 = paragraph.add_run("J. ANTONIO BARRIENTOS RODRIGUEZ")
      run1.bold = TRUE
      font = run1.font
      paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      run = paragraph.add_run("A QUIEN EN LO SUCESIVO SE DENOMINARA EL PATRON, Y POR LA OTRA." + nombre + "DE NACIONALIDAD "+ nacionalidad + "CON DOMICILIO "+ domicilio + ", A QUIEN EN ADELANTE SE DENOMINARA EL TRABAJADOR, DE ACUERDO CON LAS SIGUIENTES:")
      font = run.font
      paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      clau = doc.add_paragraph()
      run2 = clau.add_run("C L A U S U L A S ")
      font = run2.font
      clau.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
      font.color.rgb = RGBColor(0x00, 0x00, 0x00)
      font.name = 'Arial'
      font.size = Pt(11)
      font.bold = True

      paragraph2 = doc.add_paragraph()
      run3 = paragraph2.add_run("PRIMERA.-")
      run3.bold = TRUE
      font = run3.font
      paragraph2.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      run3 = paragraph2.add_run(" El (a) Profesor (a) manifiesta, bajo protesta de decir verdad, que tiene la Clave Única de Registro de Población " + curp + " y el Registro Federal de Contribuyentes "+ rfc + " que tiene  la capacidad, aptitudes, facultades y conocimientos necesarios para desempeñar el trabajo que se le encomienda, así como  la documentación completa y actualizada por la Secretaria de Educación Publica y/o la UNAM, así como  a las disposiciones señaladas por los artículos 42 fracción VII, de la nueva Ley Federal  del Trabajo publicada en el Diario Oficial de la Federación el día 30 de noviembre  del 2012  que se requiere así como está de acuerdo en que el no cumplir con cualquiera de estos requisitos será causa suficiente para que el patrón le rescinda su contrato de trabajo en el momento que tenga conocimiento de la carencia de alguna de esta condiciones, así mismo se compromete a que en caso de que el profesor (a)  cambie de domicilio durante la vigencia del presente contrato notificará por escrito al patrón dentro de los  cinco días siguientes que cambie de domicilio.")
      font = run3.font
      paragraph2.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      paragraph3 = doc.add_paragraph()
      run3 = paragraph3.add_run("SEGUNDA.-")
      run3.bold = TRUE
      font = run3.font
      paragraph2.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      run3 = paragraph3.add_run(" Este contrato por exigencias expresas de la Secretaría de Educación Pública se celebra por tiempo determinado, el cual se precisa en el  Acuerdo ")
      font = run3.font
      paragraph3.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)
      
      run3 = paragraph3.add_run("14/072020")
      run3.bold = TRUE
      font = run3.font
      paragraph2.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      run3 = paragraph3.add_run(" de la Secretaría de Educación Pública publicado en el Diario Oficial del ")
      font = run3.font
      paragraph3.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      run3 = paragraph3.add_run("03 DE AGOSTO DEL 2020,")
      run3.bold = TRUE
      font = run3.font
      paragraph2.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      run3 = paragraph3.add_run(" y sólo podrá modificarse, rescindirse o terminarse en los casos y condiciones especificados en la Ley Federal del Trabajo, o por aquellas autoridades que en su momento cuenten con facultades suficientes para modificar, rescindir o dar por terminado el presente contrato. ")
      font = run3.font
      paragraph3.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)


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