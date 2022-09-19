# GENERADOR DE PLANILLAS
Instalación de Entorno virtual:

```console
python -m virtualenv .vnev
```
Activación del entorno virtual en windows:

```console
.venv/Scripts/activate
```

Instalación de dependencias:

```console
pip install -r requeriments.txt
```

Ejecucion del script:
```console
python main.py
```

# IDEA


1. tener un panel para ingresar que comision debe obtenerse el listado, si es default entonces selecciona todas las comisiones

## Para una sola comisión:

2. Enviar ese dato a dataframe.py y obtener un dataframe con todos los alumnos de dicha comisión

3. Buscar los datos de la comisin en sedes_comisiones.xlsx y retornar un dict

3. Exportar ese dataframe a un excel con startrow=4
4. Realizar el diseño con openpyxl
