import smtplib
import datetime
from src.Emails.crear_correos import crearCorreos



#ahora creamos una funcion para enviar errores de envio de correos
def enviar_correo_notificacion(ruta_consolidado_diario, dia_habil):

    to_email_notificacion = ['aprendiz.funcional@gruporeditos.com', 'cristian.avendano@gruporeditos.com', 'jose.chaverra@gruporeditos.com', 'natalia.vanegas@gruporeditos.com']
    #to_email_notificacion = ['aprendiz.funcional@gruporeditos.com', 'cristian.avendano@gruporeditos.com', 'jose.chaverra@gruporeditos.com']
    # Obtener las credenciales de la cuenta de correo
    correo = crearCorreos(to_email_notificacion)
    smtp_server, smtp_port, smtp_username, smtp_password = correo.conexion_correo()

    hora_actual = datetime.datetime.now()
    HORA_ACTUAL = hora_actual.strftime("%H:00")
    # Definir el asunto y el mensaje del correo
    asunto = f"AUSENTISMOS FASE 2 - EJECUCION EXITOSA - {dia_habil} del corte de la {HORA_ACTUAL}"
    message = f"""
            <strong>Cordial Saludo,</strong>
            <br><br>
            Les informamos que los ausentismos correspondientes al día {dia_habil} ya han sido procesados correctamente.
            <br><br>
            """
        
    # Contenido del correo (puede incluir HTML)
    cuerpo = f"""
    <html>
    <body>
        {message}<br>
        <br>
        Por favor no responder ni enviar correos de respuesta a la cuenta correo.automatizacion@gruporeditos.com.
        </p>
    </body>
    </html>
    """

    # Crear el mensaje
    msg = correo.crear_mensaje(asunto, cuerpo)
    # Adjuntar el archivo de justificación
    correo.adjuntar_archivos(msg, ruta_consolidado_diario)

    try:
        # Conectar al servidor SMTP y enviar el correo
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(smtp_username, smtp_password)
            server.send_message(msg)
            print("Correo de notificacion enviado correctamente.")
    except Exception as e:
        print(f"Error al enviar el correo de notificacion. Mensaje de error: {e}")