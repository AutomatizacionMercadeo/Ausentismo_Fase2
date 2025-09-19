#enviar correo de error
import smtplib
import locale
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from src.Fuji.get_data import get_datos_id
from datetime import datetime
from datetime import datetime, timedelta
from email.mime.application import MIMEApplication


class crearCorreos():
    def __init__(self, correo_destinatario):

        #Dirección de correo electrónico del remitente

        self.to_email = correo_destinatario
        self.from_email = ''

    def conexion_correo(self):
        try:
            # Obtener las variables de la DB
            data = get_datos_id('1')

            # Extraer las variables del objeto data
            smtp_server = data['server_smtp']
            smtp_port = data['port_smtp']
            smtp_username = data['user_smtp']
            smtp_password = data['pass_smtp']
            self.from_email = smtp_username

            # Devolver cuenta con las credenciales
            return smtp_server, smtp_port, smtp_username, smtp_password

        except Exception as e:
            print(f"Error al conectar al correo. Mensaje de error: {e}")

    def crear_mensaje(self, asunto, cuerpo):
        # Crear el mensaje
        msg = MIMEMultipart()
        msg['From'] = self.from_email
        msg['To'] = ', '.join(self.to_email)
        msg['Subject'] = asunto

        # Adjuntar el cuerpo del mensaje
        msg.attach(MIMEText(cuerpo, 'html'))

        return msg
    

    def preparar_correo(self, dia_habil, zona, municipio):
    
        # Definir el asunto y el mensaje del correo
        asunto = f"RECORDATORIO AUSENTISMO_{municipio}_{zona} PARA LA FECHA - {dia_habil}"
        message = f"""
            <strong>Cordial Saludo,</strong>
            <br><br>
            Este es un recordatorio automático sobre los registros de <strong>ausentismo SIN JUSTIFICACIÓN</strong> 
            detectados para su centro de costo en la fecha <strong>{dia_habil}</strong>.
            <br><br>
            Adjuntamos el archivo en Excel con el detalle de los empleados relacionados, 
            incluyendo nombre y cédula, para su revisión.
            <br><br>
            Le solicitamos validar la información y tomar las acciones correspondientes. 
            En caso de existir una justificación, por favor remitirla al correo: 
            <strong>correo.automatizacion@gruporeditos.com</strong>.
            <br><br>
            Gracias por su atención y gestión oportuna.
        """

        # Contenido del correo (HTML)
        cuerpo = f"""
        <html>
        <body>
            {message}
            <br>
            <p style="color: gray; font-size: 12px;">
                Por favor no responder ni enviar correos de respuesta directamente a esta cuenta.
            </p>
        </body>
        </html>
        """


        return asunto, cuerpo


    #ahora creamos una funcion para enviar el correo con su adjunto, asunto y cuerpo
    def enviar_correo(self, asunto, cuerpo, ruta_excel):
        # Obtener las credenciales de la cuenta de correo
        smtp_server, smtp_port, smtp_username, smtp_password = self.conexion_correo()

        # Crear el mensaje
        msg = self.crear_mensaje(asunto, cuerpo)
        # Adjuntar el archivo Excel
        self.adjuntar_archivos(msg, ruta_excel)
        try:
            # Conectar al servidor SMTP y enviar el correo
            with smtplib.SMTP(smtp_server, smtp_port) as server:
                server.starttls()
                server.login(smtp_username, smtp_password)
                server.send_message(msg)
                print("Correo enviado correctamente.")
        except Exception as e:
            print(f"Error al enviar el correo con adjunto. Mensaje de error: {e}")


    def adjuntar_archivos(self, msg , ruta_justificacion):

        #validamos si la ruta del archivo existe
        if not os.path.exists(ruta_justificacion):
            print(f"El archivo {ruta_justificacion} no existe.")
            return False

        # Adjuntar el archivo Excel al mensaje
        try:
            with open(ruta_justificacion, 'rb') as f:
                part = MIMEApplication(f.read(), Name=os.path.basename(ruta_justificacion))
                part['Content-Disposition'] = f'attachment; filename="{os.path.basename(ruta_justificacion)}"'
                msg.attach(part)
            print(f"Archivo adjuntado correctamente: {ruta_justificacion}")
            return True
        except Exception as e:
            print(f"Error al adjuntar el archivo. Mensaje de error: {e}")
            return False