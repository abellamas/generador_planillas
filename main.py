from app import settings
from app.listado import Listado, ExcelDataframe

def execute(comision_id):
  planilla = Listado(
    comision = comision_id,
    template = settings.TEMPLATES + '\\template.xlsx',
    output_dir = settings.OUTPUT_DIR,
    database = settings.REGISTRO,
    registro_historico='2022 1C',
    sheetname = 'Listado'
  )
  
  print(f"Cargando datos comision {comision_id}...")
  
  planilla.load_data()
  # planilla.print_df()
  
  print("Realizando escritura de planilla en Excel...")
  
  
  planilla.write(
    cols_dataframe=['comision','dni','apellido','nombre','email','telefono','observaciones'],
    wr_cols_excel = ['B','C','D','F','H','I','J'],
    start_row=6
  )
  
  
  df_info_comision = ExcelDataframe(settings.SEDES_COMISIONES,'COMISIONES').filter_data('comision',planilla.comision).to_dict('records')[0]
  
  header_listado = {
    'E1': df_info_comision['comision'],
    'E2': df_info_comision['direccion'],
    'E3': df_info_comision['turno'],
    'F1': df_info_comision['sede'],
    'H3': df_info_comision['orientacion'],
    'J1':'=COUNTIF(C6:C36,"<>")', # alumnos formula excel
    'J2': df_info_comision['referente'],
    'J3': df_info_comision['telefono'],
    'L1': df_info_comision['dias'],
    'L2': df_info_comision['horarios'],
    'L3': df_info_comision['cuatrimestre']
  }
  
  planilla.headers(structure=header_listado)

  # hoja calificaciones

  planilla.set_sheetname('Calificaciones')
  planilla.load_data()
  planilla.write(
    cols_dataframe = ['apellido','nombre','dni','fnac'],
    wr_cols_excel = ['B','C','E','G'],
    start_row=16
  )
  header_calificaciones = {
    'C8': df_info_comision['comision'],
    'C9': df_info_comision['sede'],
    'B10': df_info_comision['cuatrimestre'],
    'F9': df_info_comision['direccion'],
    'F10': df_info_comision['cuatrimestre']
  }
  planilla.headers(structure=header_calificaciones)
  planilla.save()
  
  print(f"Planilla de la comisión {planilla.comision} creada con exito!")
  
  
  
  
def main():
  
  comision_input = input('Ingrese la comisión (ENTER para obtener todas): ')
  
  try:
    
    if comision_input != '':
      execute(comision_id=comision_input)
    
    else:
      comisiones = list(ExcelDataframe(settings.SEDES_COMISIONES,'COMISIONES').filter_data('comision', 'All'))

      for comision in comisiones:        
        execute(comision_id=comision)

  except Exception as e:
    print("No se a podido crear la planilla")
    print(e)
    
  
  
  
  

  

if __name__ == '__main__':
  main()