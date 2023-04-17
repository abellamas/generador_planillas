from app import settings
from app.listado import Listado, ExcelDataframe


def execute(comision_id):
    
    planilla = Listado(
        comision=comision_id,
        template=settings.TEMPLATES + '\\template.xlsx',
        output_dir=settings.OUTPUT_DIR,
        database=settings.DB,
        registro_historico='ALUMNOS',
        sheetname='Listado'
    )

    print(f"Cargando datos comision {comision_id}...")

    planilla.load_data()

    print("Realizando escritura de planilla en Excel...")

    # planilla de alumnos

    planilla.write(
        cols_dataframe=['comision', 'dni', 'apellido', 'nombre', 'email', 'telefono', 'observaciones',
                        'dni_doc', 'p_nac', 'certificado'],
        wr_cols_excel=['B', 'C', 'D', 'F', 'H', 'I', 'J', 'M', 'N', 'O'],
        start_row=8
    )

    df_info_comision = ExcelDataframe(settings.DB, 'COMISIONES').filter_data('comision', planilla.comision).to_dict(
        'records')[0]
    
    planilla.tutor = df_info_comision['tutor']
    # planilla.comision_name = planilla.comision.replace('/', ' ')
    
    header_listado = {
        'E2': df_info_comision['comision'],
        'E3': df_info_comision['sede'],
        'E4': df_info_comision['direccion'],
        'E5': df_info_comision['turno'],

        'H2': df_info_comision['orientacion'],
        'H3': df_info_comision['cuatrimestre'],
        'H4': df_info_comision['dias'],
        'H5': df_info_comision['horarios'],

        'K2': '=COUNTIF(C8:C37,"<>")',  # cantidad de alumnos formula excel
        'K3': df_info_comision['referente'],
        'K4': df_info_comision['telefono'],
    }

    planilla.headers(structure=header_listado)

    # planilla de calificaciones

    planilla.set_sheetname('Calificaciones')
    planilla.load_data()
    planilla.write(
        cols_dataframe=['apellido', 'nombre', 'dni'],
        wr_cols_excel=['B', 'C', 'D'],
        start_row=15
    )

    header_calificaciones = {
        'C8': df_info_comision['comision'],
        'C9': df_info_comision['sede'],
        'B10': df_info_comision['cuatrimestre'][0:2],#'1°'
        'F9': df_info_comision['direccion'],
        'F10': df_info_comision['cuatrimestre'][-2:] #'1C'
    }

    planilla.headers(structure=header_calificaciones)

    # Planilla de Asistencia

    planilla.set_sheetname('Asistencia')
    planilla.load_data()
    planilla.write(
        cols_dataframe=['apellido', 'nombre', 'dni'],
        wr_cols_excel=['B', 'C', 'D'],
        start_row=8
    )

    header_asistencia = {
        'C2': df_info_comision['comision'],
        'C3': df_info_comision['sede'],
        'K2': df_info_comision['cuatrimestre'],
        'K3': df_info_comision['direccion'],
    }

    planilla.headers(structure=header_asistencia)

    planilla.set_sheetname('Listado')
    planilla.save()
    planilla.convert_to_pdf(sheetnames=['Listado', 'Asistencia', 'Calificaciones'])

    print(f"Planilla de la comisión {planilla.comision} creada con exito!")


def main():
    comision_input = input('Ingrese la comisión (ENTER para obtener todas): ')

    try:

        if comision_input != '':
            execute(comision_id=comision_input)

        else:
            comisiones = list(ExcelDataframe(settings.DB, 'COMISIONES').filter_data('comision', 'All'))

            for comision in comisiones:
                execute(comision_id=comision)

    except Exception as e:
        print("No se a podido crear la planilla")
        print(e)


if __name__ == '__main__':
    main()
