import os
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent.parent

#Folders
DATABASES = os.path.join(BASE_DIR, 'databases')
OUTPUT_DIR = os.path.join(BASE_DIR, 'output')
TEMPLATES = os.path.join(BASE_DIR, 'templates')

#Files
REGISTRO = os.path.join(DATABASES, 'REGISTRO DE ALUMNOS.xlsx')
SEDES_COMISIONES = os.path.join(DATABASES, 'SEDES Y COMISIONES 2022 1C.xlsx')


