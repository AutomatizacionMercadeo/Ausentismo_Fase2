import os
import re
import datetime as dt
import sys
from O365 import Account, FileSystemTokenBackend
from src.Fuji.get_data import get_datos_id
from src.Emails.crear_correos import crearCorreos
from datetime import datetime, timedelta



class DescargaCorreo():
    
    def __init__(self):

        # Rutas
        self.path_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        self.path_email = os.path.join(self.path_root, "Emails","Estado.txt")
        self.path_download = os.path.join(self.path_root, "Reportes_Ausentismos")
        self.crear_correo = crearCorreos(correo_destinatario=None)

    def conexion_correo(self):

        try:
            # Obtener las variables de la DB
            data = get_datos_id('1')

            # Extraer las variables del objeto data
            client_id = data['client_id']
            client_secret = data['secret_id']
            tenant_id_ = data['tenant_id']

            # print(f"ID de cliente: {client_id}")
            # print(f"Secreto de cliente: {client_secret}")
            # print(f"ID de inquilino: {tenant_id_}")

            # Definir ruta raíz
            path_root = os.path.dirname(os.path.abspath(__file__))
            # Construir la ruta del archivo de token
            token_path = os.path.join(path_root,'o365_token.txt')

            # Definir credenciales de acceso a la cuenta de correo
            credentials = (client_id, client_secret)
            token_backend = FileSystemTokenBackend(token_path=token_path)
            account = Account(credentials, tenant_id=tenant_id_, token_backend=token_backend)
            
            # Devolver cuenta con las credenciales
            return account

        except Exception as e:
            print(f"Error al conectar al correo. Mensaje de error: {e}")



    def convertir_asunto(self, texto):

        reemplazos = {
            'á': 'a', 'é': 'e', 'í': 'i', 'ó': 'o', 'ú': 'u',
            'Á': 'A', 'É': 'E', 'Í': 'I', 'Ó': 'O', 'Ú': 'U'
        }
        for acento, sin_acento in reemplazos.items():
            texto = texto.replace(acento, sin_acento)
            # Limpiar el asunto eliminando los prefijos "RE:", "FW:", "RV:"
        subject = texto.strip()
        for prefix in ["RE:", "FW:", "RV:"]:
            if subject.startswith(prefix):
                subject = subject[len(prefix):].strip()
        
        # Eliminar texto entre paréntesis
        subject = re.sub(r"\(.*?\)", "", subject).strip()

        # Eliminar espacios duplicados intermedios
        subject = ' '.join(subject.split())
        
        return subject


    def descargar_correo(self, dia_anterior):
        
        # Inicio y fin del día actual
        hoy = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
        mañana = hoy + timedelta(days=1)

        subject = f"AUSENTISMO {dia_anterior}"
        # name_files = [
        #     f"AUSENTISMO_{dia_anterior}.xls",
        #     f"AUSENTISMO_{dia_anterior}.xlsx"
        # ]

        account = self.conexion_correo()
        if not account:
            return False, "No se pudo establecer conexión con el correo."
        try:
            if not account.is_authenticated:
                if not account.authenticate(scopes=['basic', 'message_all']):
                    return False, "Error de autenticación"
                print("- Autenticado correctamente.")
            else:
                print("- Autenticado con token almacenado.")

            folder = account.mailbox().get_folder(folder_name='Ausentismo')
            
            query = folder.new_query('receivedDateTime').greater_equal(hoy).less(mañana)
            mensajes = list(folder.get_messages(query=query, order_by='receivedDateTime asc', download_attachments=True))

            if not mensajes:
                print("No hay mensajes en la carpeta.")
            
                return False, "No hay mensajes en la carpeta Ausentismo en el correo de automatizacion."

            for mensaje in mensajes:
                if mensaje.is_read:
                    continue  # Ignorar mensajes ya leídos

                #Valida el asunto del mensaje
                asunto_limpio = self.convertir_asunto(mensaje.subject).strip()

                tiene_base = "AUSENTISMO" in asunto_limpio
                tiene_fecha = str(dia_anterior) in asunto_limpio


                if not (tiene_base and tiene_fecha):
                    mensaje.mark_as_read()
                    print(f"Asunto incorrecto: '{mensaje.subject}'")

                    # Reenviar el mismo correo al remitente y con copia a soporte
                    rv = mensaje.forward()
                    rv.to.add(mensaje.sender.address)              # remitente original
                    rv.cc.add("jose.chaverra@gruporeditos.com", "cristian.avendano@gruporeditos.com")      # copia al soporte
                    rv.subject = f"RV: {mensaje.subject}"

                    # Aquí agregamos el mensaje antes del contenido reenviado
                    rv.body = (
                        f"""
                            <p style="font-family: Arial, sans-serif; font-size: 14px; color: #333;">
                                Estimado usuario,<br><br>
                                Este correo ha sido reenviado automáticamente debido a una discrepancia en la fecha detectada.<br><br>

                                <b>Asunto recibido:</b> {mensaje.subject}<br>
                                <b>Asunto esperado:</b> {subject}<br><br>

                                Le solicitamos amablemente verificar la informacion correspondiente y reenviar el correo a la siguiente dirección: 
                                <a href="mailto:correo.automatizacion@gruporeditos.com" style="color: #1a73e8; text-decoration: none;">
                                    correo.automatizacion@gruporeditos.com
                                </a>.<br><br>

                                Agradecemos su colaboración.
                            </p>
                            <hr style="border: 0; border-top: 1px solid #ccc; margin-top: 20px;">
                            <p style="font-family: Arial, sans-serif; font-size: 12px; color: #777;">
                                Este es un mensaje generado automáticamente por el sistema de automatización de ausentismos.<br>
                                Por favor, no responda a este correo.
                            </p>
                            """

                        + rv.body
                    )
                    rv.send()

                    print(f"Correo reenviado a {mensaje.sender.address} y copia a soporte.")
                    continue  # Ignorar este mensaje y pasar al siguiente
                
                #Valida si el mensaje tiene archivos adjuntos
                if not mensaje.attachments or len(mensaje.attachments) == 0:
                    mensaje.mark_as_read()
                    print("El mensaje no tiene archivos adjuntos.")

                    # Aquí puedes notificar el error
                    rv = mensaje.forward()
                    rv.to.add(mensaje.sender.address)
                    rv.cc.add("jose.chaverra@gruporeditos.com", "cristian.avendano@gruporeditos.com")
                    rv.subject = f"RV: {mensaje.subject}"
                    rv.body = (
                        f"""<p>
                        Este correo fue reenviado automáticamente porque no contiene archivos adjuntos.<br>
                        <b>Error:</b> El mensaje no tiene archivos adjuntos.<br>
                        Por favor, verifique y vuelva a enviar el correo con el archivo requerido, al correo: 
                        <a href="mailto:correo.automatizacion@gruporeditos.com">correo.automatizacion@gruporeditos.com</a>.<br>
                        Gracias.
                        </p>
                        <hr>
                        """
                        + rv.body
                    )
                    rv.send()
                    print(f"Correo reenviado a {mensaje.sender.address} y copia a soporte.")
                    continue  # Ignorar este mensaje y pasar al siguiente


                # Recorremos los archivos adjuntos del mensaje
                adjuntos_validos = []

                for adjunto in mensaje.attachments:
                    print(f"Revisando adjunto: {adjunto.name}")
                    # validar la extensión del archivo
                    if adjunto.name.endswith('.xlsx') or adjunto.name.endswith('.xls'):
                        
                        # validar el nombre del archivo
                        if "AUSENTISMO" in adjunto.name.upper():
                            adjuntos_validos.append(adjunto)  #guardamos adjuntos válidos
                        else:
                            mensaje.mark_as_read()
                            print(f"Error: Nombre de archivo no esperado: {adjunto.name}")
                            
                            # Aquí notificamos el error a la oficina que envió el correo
                            rv = mensaje.forward()
                            rv.to.add(mensaje.sender.address)
                            rv.cc.add("jose.chaverra@gruporeditos.com", "cristian.avendano@gruporeditos.com")
                            rv.subject = f"RV: {mensaje.subject}"
                            rv.body = (
                                f"""<p>
                                Este correo fue reenviado automáticamente porque el nombre o la fecha del archivo adjunto no es el esperado.<br><br>
                                <b>Error:</b> El archivo adjunto {adjunto.name} no tiene el nombre esperado.<br>
                                <b>Nombre esperado que contenga:</b> AUSENTISMO<br><br>
                                Por favor, verifique y vuelva a enviar el correo con el nombre o la fecha del archivo requerido al correo: 
                                <a href="mailto:correo.automatizacion@gruporeditos.com">correo.automatizacion@gruporeditos.com</a>.<br>
                                Gracias.
                                </p>
                                <hr>
                                """
                                + rv.body
                            )
                            rv.send()
                            print(f"Correo error reenviado a {mensaje.sender.address} y copia a soporte.")   
                            continue  # Ignorar este adjunto y pasar al siguiente

                    else:
                        # Ignoramos archivos que no sean xls o xlsx
                        print(f"Error: archivo no tiene la extensión correcta {adjunto.name}")
                        mensaje.mark_as_read()
                        rv = mensaje.forward()
                        rv.to.add(mensaje.sender.address)
                        rv.cc.add("jose.chaverra@gruporeditos.com", "cristian.avendano@gruporeditos.com")
                        rv.subject = f"RV: {mensaje.subject}"
                        rv.body = (
                            f"""<p>
                            Este correo fue reenviado automáticamente porque la extensión del archivo adjunto no es la esperada.<br><br>
                            <b>Error:</b> El archivo adjunto {adjunto.name} no tiene la extensión esperada.<br>
                            <b>Extensiones esperadas:</b> .xls o .xlsx<br><br>
                            Por favor, verifique y vuelva a enviar el correo con la extensión correcta al correo: 
                            <a href="mailto:correo.automatizacion@gruporeditos.com">correo.automatizacion@gruporeditos.com</a>.<br>
                            Gracias.
                            </p>
                            <hr>
                            """
                            + rv.body
                        )
                        rv.send()
                        print(f"Correo error reenviado a {mensaje.sender.address} y copia a soporte.")   
                        continue  # Ignorar este adjunto y pasar al siguiente

                # Al terminar el bucle, si hay archivos válidos descargamos el último
                if adjuntos_validos:
                    ultimo_adjunto = adjuntos_validos[-1]  # último válido de la lista
                    ultimo_adjunto.save(self.path_download)
                    mensaje.mark_as_read()  # marca como leído el mensaje
                    print(f"- Último archivo {ultimo_adjunto.name} descargado correctamente.")
                    return True, f"Archivo descargado: {ultimo_adjunto.name}"
                else:
                    mensaje.mark_as_read()
                    print("No se encontró ningún archivo válido en el mensaje.")
                    return False, "No se encontró archivo válido"

    
            
            print("No se encontró mensaje válido o NO leido en el ultimo correo recibido.")
            
            
            ############# FIN DEL PROCESO DE DESCARGA DE CORREO #############    
            return False, "No se encontro mas correos NO LEIDOS para procesar"

        except Exception as e:
            error_msg = f"Error al procesar correos: {e}"
            print(error_msg)
            return False, error_msg


        
            

