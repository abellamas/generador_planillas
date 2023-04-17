import os
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent.parent

#Folders
DATABASES = os.path.join(BASE_DIR, 'databases')
OUTPUT_DIR = os.path.join(BASE_DIR, 'output')
TEMPLATES = os.path.join(BASE_DIR, 'templates')


#Files
DB = os.path.join(DATABASES, '2023 1C FINES.xlsx')


