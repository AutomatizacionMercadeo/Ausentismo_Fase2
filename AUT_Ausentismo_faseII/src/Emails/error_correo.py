import smtplib
from src.Emails.crear_correos import crearCorreos


#ahora creamos una funcion para enviar errores de envio de correos
def enviar_error_correo(error_message, ruta_justificacion, dia_anterior, correo_destinatario):
    # Obtener las credenciales de la cuenta de correo
    correo = crearCorreos(correo_destinatario)
    smtp_server, smtp_port, smtp_username, smtp_password = correo.conexion_correo()

    error_message_html = error_message.replace("\n", "<br>")

    # Definir el asunto y el mensaje del correo
    asunto = f"Error: Problema al enviar correo de justificación de ausentismo {dia_anterior}"
    cuerpo = f"""
    <html>
    <body>
        <strong>Cordial Saludo,</strong>
        <br><br>
        <strong>Error:</strong> {error_message_html}
        <br><br>
        <br>
        Por favor no responder ni enviar correos de respuesta a la cuenta correo.automatizacion@gruporeditos.com.
        </p>
    </body>
    </html>
    """

    # Crear el mensaje
    msg = correo.crear_mensaje(asunto, cuerpo)
    # Adjuntar el archivo de justificación
    correo.adjuntar_archivos(msg, ruta_justificacion)

    try:
        # Conectar al servidor SMTP y enviar el correo
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(smtp_username, smtp_password)
            server.send_message(msg)
            print("Correo de error enviado correctamente.")
    except Exception as e:
        print(f"Error al enviar el correo de error. Mensaje de error: {e}")