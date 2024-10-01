from openpyxl import Workbook
from openpyxl.styles import Protection
from openpyxl.worksheet.datavalidation import DataValidation
import xlwings as xw

def inicializar_hoja(ws):
    # Escribir Alumno1, Alumno2, ..., Alumno9 en las celdas A2 hasta A11 y protegerlas
    for i in range(2, 12):
        cell = ws[f'A{i}']
        cell.value = f'Alumno{i-1}'
        cell.protection = Protection(locked=True)  # Proteger celdas A2:A11

    # Escribir Pregunta1, Pregunta2, ..., Pregunta8 en las celdas B1 hasta F1 y protegerlas
    for j in range(2, 9):
        cell = ws.cell(row=1, column=j)
        cell.value = f'Pregunta{j-1}'
        cell.protection = Protection(locked=True)  # Proteger celdas B1:F1

    # Crear lógica para preguntas que acepten enteros y decimales, incluyendo negativos (P1)
    dv1 = DataValidation(type="decimal",
                        operator="between",
                        formula1=None,
                        formula2=None,
                        showErrorMessage=True,
                        errorTitle="Entrada inválida",
                        error="Solo se permiten números enteros o decimales con separador ','")

    for row in ws['B2:B11']:  # Solo para Pregunta 1 (Columna B)
        for cell in row:
            dv1.add(cell)

    # Crear lista desplegable para preguntas abiertas simples (P2)
    dv2 = DataValidation(type="list",
                        formula1='"0,1,2"',
                        showErrorMessage=True,
                        errorTitle="Valor Inválido",
                        error="El valor debe ser uno de los seleccionados en la lista.")

    for row in ws['C2:C11']:  # Solo para Pregunta 2 (Columna C)
        for cell in row:
            dv2.add(cell)

    # Crear lógica para preguntas que acepten fracciones (P3)
    dv3 = DataValidation(type="custom",
                        formula1='=AND(ISNUMBER(VALUE(LEFT(D2,FIND("/",D2)-1))),ISNUMBER(VALUE(MID(D2,FIND("/",D2)+1,LEN(D2)-FIND("/",D2)))),COUNTIF(D2,"*/?*")=1)',
                        showErrorMessage=True,
                        error="Solo se permiten fracciones en formato numerador/denominador, ej: 1/2",
                        errorTitle="Entrada inválida")

    for row in ws['D2:D11']:  # Solo para Pregunta 3 (Columna D)
        for cell in row:
            dv3.add(cell)
            cell.number_format = '@'  # Formato de texto

    # Crear lista desplegable para preguntas de selección única (P4)
    dv4 = DataValidation(type="list",
                        formula1='"A, B, C, D, N, N/C"',
                        showErrorMessage=True,
                        errorTitle="Valor Inválido",
                        error="El valor debe ser uno de los seleccionados en la lista.")

    for row in ws['E2:E11']:  # Solo para Pregunta 4 (Columna E)
        for cell in row:
            dv4.add(cell)

    # Crear lógica para preguntas que acepten pares ordenados (P6)
    dv6 = DataValidation(type="custom",
                        formula1='=AND(ISNUMBER(VALUE(LEFT(G2,FIND(";",G2)-1))),ISNUMBER(VALUE(MID(G2,FIND(";",G2)+1,LEN(G2)-FIND(";",G2)))),COUNTIF(G2,"*;?*")=1)',
                        showErrorMessage=True,
                        error="Solo se permiten valores en formato de par ordenado, ej: X;Y",
                        errorTitle="Entrada inválida")
    
    for row in ws['G2:G11']:
        for cell in row:
            dv6.add(cell)
            cell.number_format = '@'  # Formato de texto

    # Agregar validaciones a la hoja
    ws.add_data_validation(dv1)
    ws.add_data_validation(dv2)
    ws.add_data_validation(dv3)
    ws.add_data_validation(dv4)
    ws.add_data_validation(dv6)

    # Desbloquear solo las celdas B2:H11 (Preguntas 1 a 6)
    for row in ws_preguntas['B2:H11']:
        for cell in row:
            cell.protection = Protection(locked=False)

    # Bloquear las demás celdas por defecto
    ws.protection.sheet = False 

# Crear un archivo Excel
wb = Workbook()

# Inicializar la hoja principal
ws_preguntas = wb.active
ws_preguntas.title = 'Preguntas'
inicializar_hoja(ws_preguntas)

ws_preguntas.column_dimensions['F'].width = 40  # Ajustar ancho de columna F (Pregunta 5)
ws_preguntas.column_dimensions['H'].width = 52.56  # Ajustar ancho de columna F (Pregunta 5)

# Crear e inicializar una nueva hoja
ws_respuestas = wb.create_sheet(title='Respuestas')
inicializar_hoja(ws_respuestas)
ws_respuestas.protection.sheet = True


# Referenciar los datos de B2:E11 en la hoja 'Respuestas', pero dejando vacía si la celda original está vacía
for i in range(2, 12):
    for j in range(2, 6):
        cell_respuesta = ws_respuestas.cell(row=i, column=j)
        cell_pregunta = ws_preguntas.cell(row=i, column=j)
        cell_respuesta.value = f'=IF(Preguntas!{cell_pregunta.coordinate}="","",Preguntas!{cell_pregunta.coordinate})'

# Referenciar los datos de la Pregunta 6 (columna G) en la hoja 'Respuestas'
for i in range(2, 12):
    cell_respuesta = ws_respuestas.cell(row=i, column=7)
    cell_pregunta = ws_preguntas.cell(row=i, column=7)
    cell_respuesta.value = f'=IF(Preguntas!{cell_pregunta.coordinate}="","", "("&Preguntas!{cell_pregunta.coordinate}&")")'

# Referenciar los datos de K1:K10 en la hoja 'Datos' a F2:F11 en la hoja 'Respuestas'
for i in range(2, 12):
    cell_respuesta = ws_respuestas.cell(row=i, column=6)  
    cell_dato = ws_preguntas.cell(row=i-1, column=11)  
    cell_respuesta.value = f"=Datos!K{i-1}"

# Referenciar los valores de T2:T11 en la hoja 'Datos' a H2:H11 en la hoja 'Respuestas'
for row in range(2, 12):
    ws_respuestas[f'H{row}'] = f'=Datos!T{row}'

# Crear la hoja de Datos
ws_datos = wb.create_sheet(title='Datos')
# ws_datos.sheet_state = 'veryHidden'
for i in range(2, 12):
        cell = ws_datos[f'M{i}']
        cell.value = f'Alumno{i-1}'

ws_datos['N1'] = 'Pregunta7'
ws_datos['Q1'] = 'VELOCIDAD'
ws_datos['R1'] = 'PRECISION'
ws_datos['S1'] = 'EXPRESION'


# Se define rango de filas para insertar fórmula para valores de los radio buttons
# Para posteriormente concatenar los valores y que queden como una sola expresión

for row in range(2, 12):
    ws_datos[f'Q{row}'] = f'=IF(N{row}=1, "NIVEL D", IF(N{row}=2, "NIVEL C", IF(N{row}=3, "NIVEL B", IF(N{row}=4, "NIVEL A", ""))))'

for row in range(2, 12): 
    ws_datos[f'R{row}'] = f'=IF(O{row}=1, "NIVEL B", IF(O{row}=2, "NIVEL A", ""))'


for row in range(2, 12):
    ws_datos[f'S{row}'] = f'=IF(P{row}=1, "NIVEL D", IF(P{row}=2, "NIVEL B", IF(P{row}=3, "NIVEL A", "")))'

for row in range(2, 12):
    ws_datos[f'T{row}'] = f'=Q{row} & ";" & R{row} & ";" & S{row}'

# Guardar el archivo Excel
file_path = r'C:\Users\joaquin.rodriguezm\Desktop\cfgPrueba.xlsx'
wb.save(file_path)

# Usar xlwings para agregar el código VBA que genera los CheckBoxes
app = xw.App(visible=False) 
wb = app.books.open(file_path)

# Código VBA para checkboxes y radiobuttons
vba_code = """
Sub AddCheckBoxes()
    Dim ws As Worksheet
    Dim wsLinks As Worksheet
    Set ws = ThisWorkbook.Sheets("Preguntas")
    Set wsLinks = ThisWorkbook.Sheets("Datos")
    
    Dim i As Integer
    Dim j As Integer
    Dim leftPos As Double
    Dim topPos As Double
    Dim checkBoxWidth As Double
    Dim labels As Variant
    checkBoxWidth = 45  ' Aumentar el ancho de espaciado
    
    ' Etiquetas para los checkboxes
    labels = Array("OP1", "OP2", "OP3", "OP4", "OP5")
    
    ' Crear los checkboxes en el rango F2:F11 
    For i = 2 To 11
        topPos = ws.Cells(i, 6).Top
        leftPos = ws.Cells(i, 6).Left
        
        ' Crear CheckBoxes para cada celda en el rango
        For j = LBound(labels) To UBound(labels)
            With ws.CheckBoxes.Add(leftPos + j * checkBoxWidth, topPos, 20, 15) ' Ajustar el espaciado horizontal
                .Caption = labels(j)  ' Texto al lado del checkbox
                .Value = xlOff
                
                ' Establecer la celda de enlace en la hoja Datos
                .LinkedCell = wsLinks.Cells(i - 1, j + 1).Address(External:=True)
            End With
        Next j
    Next i
    
    ' Convertir los valores VERDADERO/FALSO en 1/0 en las celdas F1:J10
    For i = 1 To 10
        For j = 1 To 5
            wsLinks.Cells(i, j + 5).Formula = "=IF(" & wsLinks.Cells(i, j).Address & ",1,0)"
        Next j
        
        ' Concatenar los valores en la columna K
        wsLinks.Cells(i, 11).Formula = "=" & wsLinks.Cells(i, 6).Address & "&" & wsLinks.Cells(i, 7).Address & "&" & wsLinks.Cells(i, 8).Address & "&" & wsLinks.Cells(i, 9).Address & "&" & wsLinks.Cells(i, 10).Address
    Next i
End Sub


Sub CrearRadioButtonsEnRango()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Dim wsDatos As Worksheet
    Set wsDatos = ThisWorkbook.Sheets("Datos")

    ' Definir el rango H2:H11
    Dim celda As Range, linkedCell As Range
    For Each celda In ws.Range("H2:H11")

        ' Determinar la celda vinculada en la hoja 'Datos'
        Set linkedCell = wsDatos.Cells(celda.Row - 1 + 1, 14) ' N2 corresponde a la fila 2 y columna 14 (N)

        ' Colocar texto en la celda actual
        celda.Value = "VELOCIDAD" & Chr(10) & "" & Chr(10) & "PRECISION" & Chr(10) & "" & Chr(10) & "EXPRESION" & Chr(10) & ""

        ' Ajustes de posicionamiento
        Dim top_offset As Integer
        Dim left_offset As Integer
        top_offset = celda.Top + -11
        left_offset = celda.Left + 55

        ' Ajustes finos de las posiciones verticales para alineación con palabras
        Dim posicionesY(1 To 3) As Integer
        posicionesY(1) = top_offset + 13    ' Ajuste para VELOCIDAD
        posicionesY(2) = top_offset + 43    ' Ajuste para PRECISION
        posicionesY(3) = top_offset + 70    ' Ajuste para EXPRESION

        ' Crear un GroupBox para VELOCIDAD
        Dim groupBoxV As Object
        Set groupBoxV = ws.GroupBoxes.Add(left_offset - 10, posicionesY(1) - 10, 250, 50)
        groupBoxV.Caption = "VELOCIDAD"
        groupBoxV.Visible = False ' Hacer el GroupBox invisible

        ' Crear los radio buttons para VELOCIDAD
        Dim j As Integer
        For j = 1 To 4
            Dim optButtonV As Object
            Set optButtonV = ws.OptionButtons.Add(left_offset + ((j - 1) * 60), posicionesY(1), 60, 15)
            Select Case j
                Case 1: optButtonV.Caption = "NIVEL D": optButtonV.LinkedCell = linkedCell.Offset(0, 0).Address(External:=True)
                Case 2: optButtonV.Caption = "NIVEL C": optButtonV.LinkedCell = linkedCell.Offset(0, 0).Address(External:=True)
                Case 3: optButtonV.Caption = "NIVEL B": optButtonV.LinkedCell = linkedCell.Offset(0, 0).Address(External:=True)
                Case 4: optButtonV.Caption = "NIVEL A": optButtonV.LinkedCell = linkedCell.Offset(0, 0).Address(External:=True)
            End Select
        Next j

        ' Crear un GroupBox para PRECISIÓN
        Dim groupBoxP As Object
        Set groupBoxP = ws.GroupBoxes.Add(left_offset - 10, posicionesY(2) - 10, 150, 50)
        groupBoxP.Caption = "PRECISIÓN"
        groupBoxP.Visible = False ' Hacer el GroupBox invisible

        ' Crear los radio buttons para PRECISION
        For j = 1 To 2
            Dim optButtonP As Object
            Set optButtonP = ws.OptionButtons.Add(left_offset + ((j - 1) * 60), posicionesY(2), 60, 15)
            Select Case j
                Case 1: optButtonP.Caption = "NIVEL B": optButtonP.LinkedCell = linkedCell.Offset(0, 1).Address(External:=True)
                Case 2: optButtonP.Caption = "NIVEL A": optButtonP.LinkedCell = linkedCell.Offset(0, 1).Address(External:=True)
            End Select
        Next j

        ' Crear un GroupBox para EXPRESIÓN
        Dim groupBoxE As Object
        Set groupBoxE = ws.GroupBoxes.Add(left_offset - 10, posicionesY(3) - 10, 250, 50)
        groupBoxE.Caption = "EXPRESIÓN"
        groupBoxE.Visible = False ' Hacer el GroupBox invisible

        ' Crear los radio buttons para EXPRESION
        For j = 1 To 3
            Dim optButtonE As Object
            Set optButtonE = ws.OptionButtons.Add(left_offset + ((j - 1) * 60), posicionesY(3), 60, 15)
            Select Case j
                Case 1: optButtonE.Caption = "NIVEL D": optButtonE.LinkedCell = linkedCell.Offset(0, 2).Address(External:=True)
                Case 2: optButtonE.Caption = "NIVEL B": optButtonE.LinkedCell = linkedCell.Offset(0, 2).Address(External:=True)
                Case 3: optButtonE.Caption = "NIVEL A": optButtonE.LinkedCell = linkedCell.Offset(0, 2).Address(External:=True)
            End Select
        Next j

    Next celda
End Sub
"""

# Agregar el código VBA al módulo
vba_module = wb.api.VBProject.VBComponents.Add(1) 
vba_module.CodeModule.AddFromString(vba_code)

# Ejecutar la macro para agregar los checkboxes y establecer los vínculos
wb.api.Application.Run("AddCheckBoxes")
wb.api.Application.Run("CrearRadioButtonsEnRango")

# Guardar y cerrar el archivo
wb.save()
wb.close()
app.quit()