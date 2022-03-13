import pandas as pd 
import openpyxl

''''''''''''''''''''''''
''''''''''''''''''''''''
''''''''''''''''''''''''
  
class ExcelDataframe:
  def __init__(self,excel,sheet):
    self.excel = excel # archivo
    self.sheet = sheet # hoja
    self.dataframe = pd.read_excel(self.excel, self.sheet)
  
  def get_dataframe(self):
    return self.dataframe
  
  def filter_data(self,key,value):
    if value == 'All':
      self.dataframe = self.dataframe[key]
      return self.dataframe
    else:
      self.dataframe = self.dataframe[self.dataframe[key] == value]
      self.dataframe.reset_index(drop=True, inplace=True)
      self.dataframe.index += 1
      return self.dataframe

  def get_column(self,column):
    return self.dataframe[column]
  
''''''''''''''''''''''''
''''''''''''''''''''''''
''''''''''''''''''''''''


class Listado:
  
  def __init__(self,comision,template, output_dir, database,registro_historico,sheetname):
    self.comision = comision
    self.template = template
    self.output_dir = output_dir
    self.database = database
    self.registro_historico = registro_historico
    self.__df = ExcelDataframe(self.database,self.registro_historico) # contiene todos los datos
    self.__wb = openpyxl.load_workbook(self.template) # abre el template para copiarlo
    self.sheetname = sheetname
    self.__sheet = self.__wb[self.sheetname] # debe ubicar la hoja donde escribir
    self.__df_filtered = ''
  
  def set_sheetname(self,new_sheetname):
    self.sheetname = new_sheetname
    self.__sheet = self.__wb[self.sheetname]
    
  def load_data(self):
    dfs_store = []
    comisiones = self.comision.split('/')

    
    if comisiones.__len__()> 1:
      df_alumnos = self.__df.get_dataframe()
      for comision in comisiones:
        result_filter = df_alumnos[df_alumnos['comision'] == comision]
        dfs_store.append(result_filter) # filtra por cada comision y lo guarda en array
      self.__df_filtered = pd.concat(dfs_store) # se concatenan los dataframes del array
      self.__df_filtered.reset_index(drop=True, inplace=True) # se resetea el indice
      self.__df_filtered.index += 1 # se le suma una unidad al indice
    else:
      self.__df_filtered = self.__df.filter_data('comision',self.comision) # filtra por solo una comision
      
  def print_df(self):
    print(self.__df_filtered)
    
  def write(self, cols_dataframe=[], wr_cols_excel=[],start_row=1):
    alumnos = self.__df_filtered[cols_dataframe].to_dict('records')
    keys_df = list(alumnos[0].keys())

    for m in wr_cols_excel:
      n=start_row
      for alumno in alumnos:
        value = alumno[keys_df[wr_cols_excel.index(m)]] 
        if value != None:
            self.__sheet[m+str(n)] = value
        n+=1
      
  def save(self):
    comision_name = self.comision.replace('/',' ')
    self.__wb.save(self.output_dir + '/' + comision_name + ' LISTADO DE ALUMNOS' + '.xlsx')
  
  def headers(self,structure):
    cells = list(structure.keys())
    for cell in cells:
      self.__sheet[cell] = structure[cell]
        
        
        
  # def load_calificador(self):
    
        
  