from Fuji.get_data import get_datos_id

def env_email():
    
    data = get_datos_id('1')
    try:
        # Extraer las variables del objeto data para el envío de correos
        smtp_server = data['server_smtp']
        smtp_port = data['port_smtp']
        smtp_username = data['user_smtp']
        smtp_password = data['pass_smtp']

        print("Credenciales SMTP obtenidas correctamente.")
        return smtp_server, smtp_port, smtp_username, smtp_password
    
    except Exception as e:
        print(f"Error al obtener las credenciales SMTP. Mensaje de error: {e}")
        return None, None, None, None


