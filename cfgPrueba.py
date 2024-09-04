from openpyxl import Workbook
from openpyxl.styles import Protection
from openpyxl.worksheet.datavalidation import DataValidation

# Crear un archivo Excel
wb = Workbook()
ws = wb.active

ws.protection.sheet = True
ws.protection.password = 'admin1234'

# Escribir Alumno1, Alumno2, ..., Alumno9 en las celdas A2 hasta A11 y protegerlas
for i in range(2, 12):
    cell = ws[f'A{i}']
    cell.value = f'Alumno{i-1}'

# Escribir Pregunta1, Pregunta2, ..., Pregunta4 en las celdas B1 hasta E1 y protegerlas
for j in range(2, 6):
    cell = ws.cell(row=1, column=j)
    cell.value = f'Pregunta{j-1}'

# Crear lista desplegable para preguntas abiertas simples (P2)
dv2 = DataValidation(type="list",
                    formula1='"0,1,2"',
                    showErrorMessage=True,
                    errorTitle="Valor Inválido",
                    error="El valor debe ser uno de los seleccionados en la lista.")

for row in ws['C2:C11']:
    for cell in row:
        dv2.add(cell)


# Crear lista desplegable para preguntas de selección múltiple (P4)
dv4 = DataValidation(type="list",
                    formula1='"A, B, C, D, N, N/C"',
                    showErrorMessage=True,
                    errorTitle="Valor Inválido",
                    error="El valor debe ser uno de los seleccionados en la lista.")

for row in ws['E2:E11']:
    for cell in row:
        dv4.add(cell)

ws.add_data_validation(dv2)
ws.add_data_validation(dv4)


for row in ws['B2:E11']:
    for cell in row:
        cell.protection = Protection(locked=False)

wb.save(r'C:\Users\joaquin.rodriguezm\Desktop\cfgPrueba.xlsx')