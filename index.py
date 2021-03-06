from tkinter import *
from tkinter import filedialog as FileDialog
from io import open
import openpyxl
import docx
from docx.shared import RGBColor
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


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
      hrs = str(hrs)
      importe = sheet.cell(row = j, column = 9).value
      importe = str(importe)
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
      
      run1 = paragraph.add_run('"INSTITUTO FRANCISCO POSSENTI, A. C." ')
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

      run1 = paragraph.add_run("J. ANTONIO BARRIENTOS RODRIGUEZ ")
      run1.bold = TRUE
      font = run1.font
      paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      run = paragraph.add_run("A QUIEN EN LO SUCESIVO SE DENOMINARA EL PATRON, Y POR LA OTRA. " + nombre + " DE NACIONALIDAD "+ nacionalidad + " CON DOMICILIO "+ domicilio + ", A QUIEN EN ADELANTE SE DENOMINARA EL TRABAJADOR, DE ACUERDO CON LAS SIGUIENTES:")
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
      run3 = paragraph2.add_run("PRIMERA.- ")
      run3.bold = TRUE
      font = run3.font
      paragraph2.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      run3 = paragraph2.add_run("El (a) Profesor (a) manifiesta, bajo protesta de decir verdad, que tiene la Clave ??nica de Registro de Poblaci??n " + curp + " y el Registro Federal de Contribuyentes "+ rfc + " que tiene  la capacidad, aptitudes, facultades y conocimientos necesarios para desempe??ar el trabajo que se le encomienda, as?? como  la documentaci??n completa y actualizada por la Secretaria de Educaci??n Publica y/o la UNAM, as?? como  a las disposiciones se??aladas por los art??culos 42 fracci??n VII, de la nueva Ley Federal  del Trabajo publicada en el Diario Oficial de la Federaci??n el d??a 30 de noviembre  del 2012  que se requiere as?? como est?? de acuerdo en que el no cumplir con cualquiera de estos requisitos ser?? causa suficiente para que el patr??n le rescinda su contrato de trabajo en el momento que tenga conocimiento de la carencia de alguna de esta condiciones, as?? mismo se compromete a que en caso de que el profesor (a)  cambie de domicilio durante la vigencia del presente contrato notificar?? por escrito al patr??n dentro de los  cinco d??as siguientes que cambie de domicilio. ")
      font = run3.font
      paragraph2.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      paragraph3 = doc.add_paragraph()
      run3 = paragraph3.add_run("SEGUNDA.- ")
      run3.bold = TRUE
      font = run3.font
      paragraph3.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      run3 = paragraph3.add_run("Este contrato por exigencias expresas de la Secretar??a de Educaci??n P??blica se celebra por tiempo determinado, el cual se precisa en el  Acuerdo ")
      font = run3.font
      paragraph3.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)
      
      run3 = paragraph3.add_run("14/072020 ")
      run3.bold = TRUE
      font = run3.font
      paragraph2.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      run3 = paragraph3.add_run("de la Secretar??a de Educaci??n P??blica publicado en el Diario Oficial del ")
      font = run3.font
      paragraph3.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      run3 = paragraph3.add_run("03 DE AGOSTO DEL 2020, ")
      run3.bold = TRUE
      font = run3.font
      paragraph2.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      run3 = paragraph3.add_run("y s??lo podr?? modificarse, rescindirse o terminarse en los casos y condiciones especificados en la Ley Federal del Trabajo, o por aquellas autoridades que en su momento cuenten con facultades suficientes para modificar, rescindir o dar por terminado el presente contrato. ")
      font = run3.font
      paragraph3.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      paragraph4 = doc.add_paragraph()
      run3 = paragraph4.add_run("TERCERA.- ")
      run3.bold = TRUE
      font = run3.font
      paragraph4.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      run3 = paragraph4.add_run("El Patr??n y el (a) Profesor (a), convienen expresamente y con fundamento en el Art. 47 Fracci??n I de la Ley Federal del Trabajo, que dentro de los primeros treinta d??as o cuando el patr??n tenga conocimiento de la carencia o incumplimiento de alguna de las condiciones b??sicas requeridas para desempe??ar el trabajo contratado, se podr?? rescindir este Contrato de Trabajo sin responsabilidad para el Patr??n. ")
      font = run3.font
      paragraph4.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      paragraph5 = doc.add_paragraph()
      run3 = paragraph5.add_run("CUARTA.- ")
      run3.bold = TRUE
      font = run3.font
      paragraph5.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      run3 = paragraph5.add_run("El (a) Profesor (a) se obliga a prestar sus servicios personales al INSTITUTO, bajo su direcci??n, dependencia y subordinaci??n, las cuales consistir??n precisamente en: ")
      font = run3.font
      paragraph5.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      paragraph6 = doc.add_paragraph()
      run3 = paragraph6.add_run("Proporcionar personalmente, a los alumnos que se le indiquen o le sean asignados ense??anza eficiente durante el tiempo determinado del ciclo Escolar vigente, y como lo dispone el art??culo 56 Bis de la nueva Ley Federal del Trabajo publicada en el Diario Oficial de la Federaci??n el d??a 30 de noviembre del 2012,  sujet??ndose a los programas y planes de estudio correspondientes que le sean entregados por el INSTITUTO, debidamente autorizados. ")
      font = run3.font
      paragraph6.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      paragraph6 = doc.add_paragraph()
      run3 = paragraph6.add_run("El presente contrato se celebra por un tiempo que ser?? de " + start + " y termina " + final)
      font = run3.font
      paragraph6.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      paragraph7 = doc.add_paragraph()
      run3 = paragraph7.add_run("QUINTA.- ")
      run3.bold = TRUE
      font = run3.font
      paragraph7.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      run3 = paragraph7.add_run("Los servicios contratados se estipulan en forma enunciativa y no limitativa; por tanto, el (a) Profesor (a)  se obliga a desempe??ar  todas las labores anexas o conexas con su obligaci??n principal y las dem??s que le ordene el Patr??n o sus representantes, tales como guardias escolares, cursos de verano, de capacitaci??n de reprogramaci??n de estudios, ex??menes extraordinarios, ex??menes de diagn??stico, noche colonial, ex??menes de admisi??n, ofrenda de d??a de muertos, posada navide??a etc, cuya retribuci??n econ??mica est?? convenida y comprendida en la Cl??usula D??cima Primera del presente contrato.")
      font = run3.font
      paragraph7.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      paragraph8 = doc.add_paragraph()
      run3 = paragraph8.add_run("De la misma manera y solo para el caso de que sea aplicable y derivado de  un caso  fortuito o fuerza mayor las modificaciones al presente contrato referidas en las reformas al art??culo  311 de la Ley Federal del Trabajo capitulo XII Bis en materia de Teletrabajo publicada en el diario Oficial de la federaci??n el d??a 11 de enero del 2021.  ")
      font = run3.font
      paragraph8.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      paragraph9 = doc.add_paragraph()
      run3 = paragraph9.add_run("La desobediencia a las ??rdenes o indicaciones del Patr??n o sus representantes para el cumplimiento del trabajo contratado, ser?? causa de rescisi??n sin responsabilidad para el Patr??n.")
      font = run3.font
      paragraph9.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      paragraphA = doc.add_paragraph()
      run3 = paragraphA.add_run("SEXTA.- ")
      run3.bold = TRUE
      font = run3.font
      paragraphA.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      run3 = paragraphA.add_run("Los servicios objeto de la relaci??n de trabajo deben prestarse en el lugar o en los lugares que designe el Patr??n o sus representantes, quedando convenido que ??ste tendr?? derecho de cambiar el lugar de trabajo del (a) profesor (a) cuando se estime pertinente o necesario, siempre y cuando dicho cambio no se traduzca en una merma de su remuneraci??n para el (la) mismo (a),esto incluye para aquellos casos donde las autoridades competentes establezcan el cierre de la fuente de trabajo derivado de caso fortuito o fuerza mayor as?? como las posibles modificaciones al presente contrato seg??n las reformas al art??culo 311 de la ley Federal del Trabajo capitulo XII Bis en materia de teletrabajo publicada en el diario oficial de la federaci??n el d??a 11 de enero del 2021. ")
      font = run3.font
      paragraphA.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      paragraphB = doc.add_paragraph()
      run3 = paragraphB.add_run("SEPTIMA.- ")
      run3.bold = TRUE
      font = run3.font
      paragraphB.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      run3 = paragraphB.add_run("La duraci??n de la jornada de trabajo ser?? de: ")
      font = run3.font
      paragraphA.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      run3 = paragraphB.add_run(hrs)
      run3.bold = TRUE
      font = run3.font
      paragraphB.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      run3 = paragraphB.add_run(" horas, ")
      run3.bold = TRUE
      font = run3.font
      paragraphB.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      run3 = paragraphB.add_run("seg??n horario anexo. El (la) Profesor (a) est?? de acuerdo en que deber?? asistir los d??as que sean necesarios para las distintas actividades que se precisan en la cl??usula quinta del presente contrato. El pago correspondiente a esta jornada de trabajo, est?? ya integrado en el sueldo convenido que se indica  en la cl??usula D??cima del presente instrumento.")
      font = run3.font
      paragraphB.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      paragraphC = doc.add_paragraph()
      run3 = paragraphC.add_run("OCTAVA.- ")
      run3.bold = TRUE
      font = run3.font
      paragraphC.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      run3 = paragraphC.add_run("El (a) Profesor (a) se obliga a desempe??ar sus labores con la intensidad, cuidado y esmero apropiados en la forma, tiempo y lugar a que se refiere este Contrato y el Reglamento Interior de Trabajo. El incumplimiento de esta disposici??n se considera falta de probidad del (a) Profesor (a) y, de ocurrir, se sancionar?? con la rescisi??n  del Contrato sin responsabilidad para el Patr??n.")
      font = run3.font
      paragraphC.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      paragraphD = doc.add_paragraph()
      run3 = paragraphD.add_run("NOVENA.- ")
      run3.bold = TRUE
      font = run3.font
      paragraphD.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      run3 = paragraphD.add_run("El (a) Profesor (a) est?? obligado a checar en el reloj o firmar las listas de asistencia a la entrada y salida de sus labores, el incumplimiento de este requisito se considerar?? como una desobediencia para todos los efectos legales a que haya lugar,  el retiro del Profesor de su lugar de Trabajo durante su jornada de labores, sin autorizaci??n, ser?? considerado como una desobediencia y por lo tanto abandono de trabajo.")
      font = run3.font
      paragraphD.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      paragraphE = doc.add_paragraph()
      run3 = paragraphE.add_run("El registrar entradas o salidas por otra persona, ser?? causa de rescisi??n del presente contrato sin responsabilidad para el Patr??n.")
      font = run3.font
      paragraphE.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      paragraphF = doc.add_paragraph()
      run3 = paragraphF.add_run("DECIMA.- ")
      run3.bold = TRUE
      font = run3.font
      paragraphF.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      run3 = paragraphF.add_run("Se conviene como salario nominal, que el Patr??n deber?? pagar por los trabajos personales que reciba del(a)Profesor(a),la cantidad de: $ "+ importe + " ( " + importe_letra + " 00/100 M. N.) Dicho salario mensual ser?? cubierto por mitad al (a) profesor (a) despu??s de sumar las prestaciones y restar los impuestos correspondientes, cada d??a quince y ??ltimo de cada mes mediante dep??sito bancario  tal y como lo dispone el art??culo 101 de la nueva Ley Federal  del Trabajo publicada  en el Diario Oficial de la Federaci??n  el d??a 30 de noviembre del 2012, en las oficinas de la Escuela, en efectivo o en cheque, o mediante alg??n otro sistema de pago que las partes estimen adecuado y seguro.") 
      font = run3.font
      paragraphF.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      paragraphG = doc.add_paragraph()
      run3 = paragraphG.add_run("Adem??s del salario nominal mencionado y de todas las prestaciones que establece la Ley Federal del Trabajo como m??nimas EL INSTITUTO otorgar?? las siguientes prestaciones.") 
      font = run3.font
      paragraphG.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      paragraphH = doc.add_paragraph()
      run3 = paragraphH.add_run("10% del salario nominal por concepto de premios por asistencia siempre y cuando el profesor no falte y  asista   a  lo convocado por el Instituto. Estos premios se entregar??n en efectivo en los mismos plazos y condiciones que el salario nominal y seg??n el Reglamento Interno de Trabajo.") 
      font = run3.font
      paragraphH.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      paragraphI = doc.add_paragraph()
      run3 = paragraphI.add_run("10% del salario nominal por concepto de premio de puntualidad. Esta prestaci??n la tendr?? el profesor cuando llegue al Instituto 5 minutos antes de la primera hora-clase seg??n su horario.  Tambi??n se entregar?? en efectivo en los mismos plazos y condiciones que el salario nominal y seg??n el Reglamento Interno de Trabajo.") 
      font = run3.font
      paragraphI.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      paragraphJ = doc.add_paragraph()
      run3 = paragraphJ.add_run("12% del salario nominal por concepto de vales de despensa como concepto de previsi??n social seg??n el Plan de Previsi??n  del Instituto Francisco, Possenti,A.C. los cuales se entregar??n el d??a 28 de cada mes, en monedero electr??nico.") 
      font = run3.font
      paragraphJ.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      paragraphK = doc.add_paragraph()
      run3 = paragraphK.add_run("13% del salario nominal por concepto de aportaci??n a un fondo de ahorro, que junto con un porcentaje igual que se retendr?? al trabajador cada quincena, se depositar?? en una cuenta bancaria y se retirar?? al final del ciclo escolar") 
      font = run3.font
      paragraphK.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      paragraphL = doc.add_paragraph()
      run3 = paragraphL.add_run("Los pr??stamos que se otorguen a los trabajadores, ser??n de acuerdo a los lineamientos que regulan el Fondo de Ahorro.") 
      font = run3.font
      paragraphL.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      paragraphM = doc.add_paragraph()
      run3 = paragraphM.add_run("En las cantidades anteriores queda comprendido el pago de los s??ptimos d??as, los d??as de descanso obligatorio, las vacaciones, los d??cimos sextos d??as del mes, as?? como los conceptos se??alados en la Cl??usula Quinta del presente contrato, as?? como las labores conexas o complementarias que desempe??a de acuerdo a su labor principal tal y como lo dispone el art??culo 56 Bis de la nueva Ley Federal del Trabajo publicada en el Diario Oficial de la Federaci??n el d??a 30 de noviembre del 2012.") 
      font = run3.font
      paragraphM.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      paragraphN = doc.add_paragraph()
      run3 = paragraphN.add_run("El (la) Profesor (a) est?? de acuerdo en que el patr??n le efect??e los descuentos de cuotas al Seguro Social que le correspondan.") 
      font = run3.font
      paragraphN.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      paragraphO = doc.add_paragraph()
      run3 = paragraphO.add_run("DECIMA PRIMERA.- ")
      run3.bold = TRUE
      font = run3.font
      paragraphO.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      run3 = paragraphO.add_run("El (a) Profesor (a) asistir?? seg??n el horario establecido en la cl??usula s??ptima  convini??ndose que si trabaja el domingo  tendr?? derecho a que se le pague una prima de un 25 % sobre su salario tabulado, quedando obligado a asistir a cursos de capacitaci??n y desarrollo los d??as establecidos por la direcci??n t??cnica, la remuneraci??n por estos ??ltimos conceptos est?? incluida e integrado en la Cl??usula Decima del presente Contrato.") 
      font = run3.font
      paragraphO.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      paragraphP = doc.add_paragraph()
      run3 = paragraphP.add_run("DECIMA SEGUNDA.- ")
      run3.bold = TRUE
      font = run3.font
      paragraphP.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      run3 = paragraphP.add_run("Son d??as de descanso obligatorio de acuerdo con el Articulo 74 de la Ley Federal del Trabajo, el 1?? de Enero, 5 de Febrero, 21 de Marzo, 1?? de Mayo, 16 de Septiembre, 20 de Noviembre, 25 de Diciembre y 1?? de Diciembre de cada seis a??os, cuando corresponda a la transmisi??n del Poder Ejecutivo Federal, quedando prohibido que el (a) Profesor (a)  labore esos d??as, salvo permiso previo y por escrito del Patr??n. El pago de estos d??as queda cubierto en la cantidad convenida como salario que aparece en la Cl??usula D??cima de este contrato. ") 
      font = run3.font
      paragraphP.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)
      
      paragraphQ = doc.add_paragraph()
      run3 = paragraphQ.add_run("Las partes aceptan los cambios a estas fechas para aprovechar los ???Puentes??? que determine la Ley Laboral.") 
      font = run3.font
      paragraphQ.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      paragraphR = doc.add_paragraph()
      run3 = paragraphR.add_run("DECIMA TERCERA.- ")
      run3.bold = TRUE
      font = run3.font
      paragraphR.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      run3 = paragraphR.add_run("El (a) Profesor (a) se compromete a sujetarse a los cursos de capacitaci??n y adiestramiento a que se refieren los Art??culos 153 A al 153 X de la Ley Federal del Trabajo.") 
      font = run3.font
      paragraphR.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      paragraphS = doc.add_paragraph()
      run3 = paragraphS.add_run("DECIMA CUARTA.- ")
      run3.bold = TRUE
      font = run3.font
      paragraphS.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      run3 = paragraphS.add_run("El (a) Profesor (a) gozar?? EXCLUSIVAMENTE de las VACACIONES Y DE LA PRIMA VACACIONAL que le correspondan con base en los Art??culos 76 y 80 de la Ley Federal del Trabajo de acuerdo con su antig??edad en la escuela. Estas vacaciones ser??n disfrutadas exclusivamente durante los periodos de la ??ltima quincena de Diciembre de cada a??o y el periodo de Semana Santa tal como lo se??ala y ordena el Calendario Oficial de la SEP vigente.") 
      font = run3.font
      paragraphS.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      paragraphT = doc.add_paragraph()
      run3 = paragraphT.add_run("El Profesor (a) se da por enterado (a) y de acuerdo en que los periodos de Julio y Agosto corresponden a RECESO DE CLASES para los alumnos, seg??n dictamen de la Secretar??a de Educaci??n P??blica y/o UNAM, y que por lo tanto, estos periodos en ning??n caso corresponde a vacaciones para el personal docente y del Colegio en general.") 
      font = run3.font
      paragraphT.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      paragraphU = doc.add_paragraph()
      run3 = paragraphU.add_run("DECIMA QUINTA.- ")
      run3.bold = TRUE
      font = run3.font
      paragraphU.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      run3 = paragraphU.add_run("Todos los estudios, planes, programas, datos y en general cualquier documentaci??n e informaci??n que el Profesor reciba o que elabore en el desempe??o de sus servicios o por el encargo espec??fico de la Escuela, ser??n propiedad de ??sta ??ltima, en cuya virtud se obliga a devolverlos en el momento de ser requeridos para ello o al terminar este Contrato, o sea el " + final) 
      font = run3.font
      paragraphU.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      paragraphV = doc.add_paragraph()
      run3 = paragraphV.add_run("DECIMA SEXTA.- ")
      run3.bold = TRUE
      font = run3.font
      paragraphV.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      run3 = paragraphV.add_run("El (a) Profesor (a) est?? de acuerdo en no divulgar con ninguna persona o en otra Instituci??n los datos, documentos, conocimientos e informes que haya obtenido con motivo de la prestaci??n de sus servicios en la Escuela, ya que tienen el car??cter de confidenciales, en caso de hacerlo ser?? considerada su conducta como falta de probidad, adem??s de las consecuencias legales que de esto origine.") 
      font = run3.font
      paragraphV.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      paragraphW = doc.add_paragraph()
      run3 = paragraphW.add_run("DECIMA SEPTIMA.- ")
      run3.bold = TRUE
      font = run3.font
      paragraphW.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      run3 = paragraphW.add_run("Los contratantes declaran que conocen el Reglamento Interior de Trabajo del Colegio, al cual se sujetar??n en todas sus Cl??usulas, por haberlo firmado protestando su estricto y legal cumplimiento.") 
      font = run3.font
      paragraphW.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      paragraphX = doc.add_paragraph()
      run3 = paragraphX.add_run("DECIMA OCTAVA.- ")
      run3.bold = TRUE
      font = run3.font
      paragraphX.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      run3 = paragraphX.add_run("El patr??n reconoce al trabajador una antig??edad a partir del " + antiguedad + ". En todo lo no previsto, en el presente Contrato, se estar?? a las disposiciones de la Ley Federal del Trabajo vigente.") 
      font = run3.font
      paragraphX.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      paragraphY = doc.add_paragraph()
      run3 = paragraphY.add_run("DECIMA NOVENA.- ")
      run3.bold = TRUE
      font = run3.font
      paragraphY.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      run3 = paragraphY.add_run("El Instituto Francisco Possenti, A. C., le informa que sus datos personales, incluyendo los sensibles se utilizar??n para identificar, informar, operar, gestionar y dem??s acciones que sean necesarias para la prestaci??n de servicios laborales subordinados en el Instituto. El derecho de acceso, rectificaci??n, cancelaci??n oposici??n, limitaci??n o la revocaci??n de uso de sus datos personales, que para tal fin nos haya otorgado, a trav??s de los procedimientos que hemos implementado, podr?? solicitarse por escrito en dependencias gubernamentales como, la SEP, el SAT, IMSS, INFONAVIT, Secretar??a del Trabajo etc., le informamos que si usted no manifiesta su oposici??n para que sus datos personales sean utilizados por el Instituto, significa que ha le??do, entendido y aceptado los t??rminos antes expuestos otorgando su consentimiento para ello.") 
      font = run3.font
      paragraphY.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      paragraphZ = doc.add_paragraph()
      run3 = paragraphZ.add_run("VIG??CIMA.- ")
      run3.bold = TRUE
      font = run3.font
      paragraphZ.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      run3 = paragraphZ.add_run("Ambas partes acuerdan que en caso de suspensi??n de las actividades escolares por causas de fuerza mayor o caso fortuito se estar?? a las disposiciones que se??alen las autoridades competentes.") 
      font = run3.font
      paragraphZ.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      paragraphA1 = doc.add_paragraph()
      run3 = paragraphA1.add_run("VIG??CIMA PRIMERA.- ")
      run3.bold = TRUE
      font = run3.font
      paragraphA1.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      run3 = paragraphA1.add_run("Ambas partes manifiestan que tienen pleno conocimiento  y cumplimiento de los alcances de la NOM-035-STPS-2018, factores de riesgo psicosocial en el  trabajo, identificaci??n, an??lisis y prevenci??n, publicada en el Diario  Oficial de la Federaci??n el 23 de octubre del 2018, en t??rminos de los dispuesto por los art??culos 40, fracciones I y XI, de la Ley Org??nica  de la Administraci??n P??blica Federal; 512, 523, fracci??n  I, 524 y 527, ??ltimo p??rrafo,  de  la Ley Federal  del Trabajo; 1??.,3??., fracci??n  XI, 38, fracci??n  II, 40, fracci??n VII, 41,47,  fracci??n IV,51, primer p??rrafo, 62,68, y 87 de la Ley Federal sobre  Metrolog??a   y Normalizaci??n ;  28 del Reglamento de la  ley federal sobre Metrolog??a  y Normalizaci??n; 5??., fracci??n III,7, fracciones I,II, III,IV,V, VII,IX,XI y XII, 8, fracciones I, III, V, VIII, X y XI,10, 32, fracci??n XI, 43,44, fracci??n VIII, y 55, del Reglamento  Federal de Seguridad y Salud en el Trabajo, y 5, fracci??n III, y 24 del Reglamento Interior de la Secretar??a  del Trabajo y Previsi??n Social.") 
      font = run3.font
      paragraphA1.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      paragraphA2 = doc.add_paragraph()
      run3 = paragraphA2.add_run("VIG??CIMA SEGUNDA.- ")
      run3.bold = TRUE
      font = run3.font
      paragraphA2.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      run3 = paragraphA2.add_run("Ambas partes est??n de acuerdo que en caso de que hubiera una contingencia sanitaria conforme a los art??culos 42 Bis, 429 fracci??n l y IV de la Ley Federal del Trabajo o declaraci??n de emergencia sanitaria derivada de cualquier tipo de virus o bacteria, o situaci??n derivada de cualquier tipo de ???caso fortuito??? o de ???fuerza mayor??? y como consecuencia de ello se tuviera que suspender la relaci??n de trabajo o cerrar por tiempo determinado o indeterminado la fuente de trabajo, incluyendo la suspensi??n de la actividad escolar por tiempo definido o indefinido, Patr??n y profesor(a) acuerdan que durante el periodo de dicha suspensi??n el salario quincenal integrado se ajustar?? de conformidad con las posibilidades econ??micas del Patr??n, lo anterior con el objetivo de que los recursos econ??micos  (ingresos) de dicho patr??n pueden ser repartidos de forma equitativa entre todos los trabajadores, incluyendo el hacer frente al pago de las obligaciones fiscales, gastos fijos y administrativos de la escuela y de seguridad social de cada Trabajador. A su vez y derivado de caso fortuito o fuerza mayor cuando las autoridades competentes nos obliguen a modificar en todo  o en parte el presente contrato se tomar?? en cuenta lo establecido en el decreto por el que se reforma el art??culo 311 y se adiciona el capitulo XII Bis de la Ley Federal de Trabajo en materia de Teletrabajo publicado en el Diario Oficial de la Federaci??n el d??a 11 de enero 2021 para aquellos casos o situaciones que apliquen en el presente contrato y relaci??n de trabajo.") 
      font = run3.font
      paragraphA2.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      paragraphA3 = doc.add_paragraph()
      run3 = paragraphA3.add_run("Le??do que fue por ambas partes este documento y sabedoras de las obligaciones que contraen, lo firman de conformidad por duplicado en el lugar y fecha se??alados a continuaci??n.") 
      font = run3.font
      paragraphA3.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
      font.name = 'Arial'
      font.size = Pt(11)

      paragraphA2 = doc.add_paragraph()
      run3 = paragraphA2.add_run("CIUDAD DE MEXICO, A " + start)
      run3.bold = TRUE
      font = run3.font
      paragraphA2.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
      font.name = 'Arial'
      font.size = Pt(11)

      table = doc.add_table(rows = 3, cols = 3)

      cell1 = table.cell(2,0)
      cell1.text = '      EL PATR??N                        C. J. ANTONIO BARRIENTOS R. Representante Legal'
      run = cell1.paragraphs[0].runs[0]
      run.font.bold = True
      run.font.name = "Arial"
      run.font.size = Pt(9)
      cell1.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

      tc = cell1._tc
      tcPr = tc.get_or_add_tcPr()
      tcBorders = OxmlElement('w:tcBorders')
      top = OxmlElement('w:top')
      top.set(qn('w:val'), 'single')

      tcBorders.append(top)
      tcPr.append(tcBorders)
      
      cell3 = table.cell(2,2)
      cell3.text = 'EL (A) PROFESOR (A) ' + nombre
      run = cell3.paragraphs[0].runs[0]
      run.font.bold = True
      run.font.name = "Arial"
      run.font.size = Pt(9)
      cell3.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
      
      tc = cell3._tc
      tcPr = tc.get_or_add_tcPr()
      tcBorders = OxmlElement('w:tcBorders')
      top = OxmlElement('w:top')
      top.set(qn('w:val'), 'single')


      tcBorders.append(top)
      tcPr.append(tcBorders)

      doc.save("Contrato " + nombre + ".docx")
      

#Creamos interfas
root = Tk()
ex = StringVar()
hoja = StringVar()
root.title('Generaci??n de contratos')

Label(root, text="Contratos", fg="darkblue", font=("Arial", 28, "bold")).pack()

#Pedimos datos necesarios de Excel
Label(root, text="Nombre del Excel",fg="black",font=("Arial", 16, "bold")).pack()
Entry(root, justify="center", textvariable=ex).pack()

Label(root, text="Nombre de la hoja", fg="black", font=("Arial", 16, "bold")).pack()
Entry(root, justify="center", textvariable=hoja).pack()

Button(root, text="Aceptar", command=crear).pack()

root.mainloop()