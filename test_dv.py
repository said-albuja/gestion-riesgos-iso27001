import openpyxl

# Restore original
import shutil
# Need original file, let's get it from git HEAD
import subprocess
subprocess.run(['git', 'restore', 'GGESTION DE RIESGOS - ACTIVOS DE INF-2026.xlsx'])

wb = openpyxl.load_workbook('GGESTION DE RIESGOS - ACTIVOS DE INF-2026.xlsx', data_only=False)
ws = wb['INVENTARIO']
print(f"Data validations before: {len(ws.data_validations.dataValidation)}")

wb.save('test_dv_output.xlsx')

wb2 = openpyxl.load_workbook('test_dv_output.xlsx', data_only=False)
ws2 = wb2['INVENTARIO']
print(f"Data validations after: {len(ws2.data_validations.dataValidation)}")
