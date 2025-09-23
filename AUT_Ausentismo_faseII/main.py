import os
from datetime import datetime
from src.Modules.correo_sin_respuesta import enviar_correo_sin_respuesta
from src.Fuji.get_data import get_datos_id
from src.Emails.descargaCorreo import DescargaCorreo
from src.Modules.procesos import Cruce_datos
from src.Modules.consolidado_mensual import reporte_consolidado_mensual
from src.SQL.insertar_consolidado_mensual_sql import insert_consolidado_mensual_sql
from src.Modules.recordatorio import filtrar_zonas_ausentismos
from src.Emails.enviar_correo_zonas import enviar_correo_zonas
from src.Emails.crear_correos import crearCorreos
from src.Modules.eliminar_archivos_antiguos import eliminar_archivos_antiguos
from src.Modules.cruce_do import cruce_do
from src.Emails.enviar_correo_notificacion import enviar_correo_notificacion


def main():
    # Definir ruta raíz
    path_root = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'src')
    ruta_reportes_vencidos = os.path.join(path_root, 'Consolidado_Mensual')
    ruta_reportes_ausentismos = os.path.join(path_root, 'Reportes_Ausentismos')

    data = get_datos_id('1')
    # print(data)


    # Crear carpeta por si no existen
    folder = os.path.join(path_root, "Reportes_Ausentismos")
    if not os.path.exists(folder):
        os.mkdir(folder)

    cruce = Cruce_datos()
    dia_habil = cruce.obtener_ultimo_dia_anterior()
    #dia_anterior = '2025-09-04'  # para pruebas

    descargar_correo = DescargaCorreo()


    # bucle para procesar correos uno por uno
    while True:
        exito, mensaje = descargar_correo.descargar_correo(dia_habil)

        if not exito:
            print(f"Error al descargar el correo: {mensaje}")
            break

        print(f"Procesando archivo descargado.....")
        data_consolidada =  cruce.cruce_datos_cedula_vs_cedula(dia_habil)

        if data_consolidada:
                print(f"Datos consolidados: {data_consolidada}")
                # tomamos los datos de data_consolidada y guardamos en un archivo excel con la misma estructura de ausentismos_sin_justificacion
                cruce.reporte_consolidado_diario(data_consolidada)
        else:
            print("No se encontraron datos para consolidar.")
    
        # Llamar a la función para generar el reporte consolidado mensual
        reporte_consolidado_mensual(dia_habil)
        # llamar a la función para insertar los datos en la base de datos SQL
        insert_consolidado_mensual_sql()


    # validamos de que si es la 1:00 pm , que envie un correo con el archivo ausentismos sin justificacion
    enviar_correo_sin_respuesta(dia_habil)

    zonas = filtrar_zonas_ausentismos()
    if zonas:
        enviar_correo_zonas(zonas, dia_habil)

    # LLAMAMOS A LA FUNCION DE ELIMINAR ARCHIVOS ANTIGUOS
    eliminar_archivos_antiguos(dia_habil, ruta_reportes_vencidos, ruta_reportes_ausentismos)

    
    # Llamar a la función para enviar el correo de notificación
    ruta_consolidado_diario = os.path.join(ruta_reportes_ausentismos, f"REPORTE_CONSOLIDADO_AUSENTISMO_DIARIO_{dia_habil}.xlsx")
    enviar_correo_notificacion(ruta_consolidado_diario, dia_habil)

    try:
        datos_no_coincidentes = cruce_do(dia_habil)
        if datos_no_coincidentes is False:
            print("El cruce de datos no se pudo completar debido a un error.")
        else:
            print("Cruce de datos completado exitosamente.")

    except Exception as e:
        print(f"Error durante el cruce de datos: {e}")

    print("Proceso finalizado correctamente.")
if __name__ == "__main__":
    main()