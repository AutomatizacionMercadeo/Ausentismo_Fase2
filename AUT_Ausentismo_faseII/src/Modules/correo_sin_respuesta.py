import os
from datetime import datetime
from src.Emails.crear_correos import crearCorreos

def enviar_correo_sin_respuesta(dia_habil):
    # obtenemos la ruta de solo src sin meternos en Modules

    path_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


    hora = datetime.now()  # obtiene la fecha y hora actual
    print(f"Hora actual: {hora}")
    # hora = datetime(2025, 9, 16, 13, 10, 0)  # para pruebas
    ultima_ejecucion = None  # Variable para rastrear la última ejecución

    # Ejecutar solo si estamos en la hora 13 y aún no se ejecutó en el día
    if hora.hour == 13 and (ultima_ejecucion is None or ultima_ejecucion.date() != hora.date()):
        hora_str = hora.strftime("%H:%M")
        

        print(f"Es la {hora_str}, enviando correos de recordatorio al proceso...")

        # llamamos a la funcion de enviar correo
        correo = [
            'cristian.avendano@gruporeditos.com',
            'jose.chaverra@gruporeditos.com',
            'aprendiz.funcional@gruporeditos.com'
        ]
        crear_correo = crearCorreos(correo)

        asunto = f"RECORDATORIO AUSENTISMO PARA LA FECHA - {dia_habil}, HORA - {hora_str}"
        cuerpo = f"""
            <strong>Cordial Saludo,</strong>
            <br><br>
            Este es un recordatorio automático sobre los registros de <strong>AUSENTISMOS SIN RESPUESTA</strong> 
            detectados para su centro de costo en la fecha <strong>{dia_habil}</strong>.
            <br><br>
            Por favor, revise y tome las acciones necesarias.
        """

        ruta_excel = os.path.join(path_root,'reportes', f"Ausentismos_SIN_JUSTIFICACION_GENERAL - {dia_habil}.xlsx")

        crear_correo.enviar_correo(asunto, cuerpo, ruta_excel)

        # Guardamos la fecha de la última ejecución
        ultima_ejecucion = hora