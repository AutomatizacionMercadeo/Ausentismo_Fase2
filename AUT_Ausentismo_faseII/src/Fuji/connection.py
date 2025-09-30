import pyodbc, os, time
from dotenv import load_dotenv


def connection(max_retries=3, retry_delay=5):
    
    # Cargar las variables de entorno desde el archivo .env
    load_dotenv()
    
    # Obtener las credenciales de las variables de entorno
    server = os.getenv('SERVER')
    database = os.getenv('DATABASE')
    usr = os.getenv('USR')
    pwd = os.getenv('PASS')
    
    # Construir cadena de conexión
    connection_string = f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={server};DATABASE={database};UID={usr};PWD={pwd};'
    
    for attempt in range(max_retries):
        try:
            conn = pyodbc.connect(connection_string)
            print("Conexión exitosa a la base de datos.")
            return conn
        except pyodbc.Error as ex:
            print(f"Error al conectar a la base de datos: {ex}")
            if attempt < max_retries - 1:
                print(f"Reintentando conexión ({attempt + 1}/{max_retries})...")
                time.sleep(retry_delay)
            else:
                print("Se ha alcanzado el número máximo de intentos. No se pudo conectar a la base de datos.")
                return None      

