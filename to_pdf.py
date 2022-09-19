import win32com.client
import os
from pywintypes import com_error
from app import settings

WB_PATH = os.path.join(settings.OUTPUT_DIR, '1Q LISTADO DE ALUMNOS.xlsx')

PATH_TO_PDF = os.path.join(settings.OUTPUT_DIR, '1Q LISTADO DE ALUMNOS.pdf')

excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False

try:
  print("Abriendo planilla...")
  # Open the workbook
  wb = excel.Workbooks.Open(WB_PATH)
  wb.Worksheets('Listado').Select()
  wb.ActiveSheet.ExportAsFixedFormat(0, PATH_TO_PDF)

except com_error as e:
  print('failed to open and convert the workbook')
  
else:
  print('Realizado')

finally:
  wb.Close()
  excel.Quit()