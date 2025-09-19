import pyodbc
from src.Fuji.connection import connection

def get_datos_id(id_unico):
    conn = connection()
    if conn:
        cursor = conn.cursor()
        try:
            # Llamar al procedimiento almacenado con el parámetro
            cursor.execute("{CALL sp_GetData(?)}", id_unico)
            # Obtener todos los registros
            rows = cursor.fetchall()
            
            # Obtener los nombres de las columnas
            columns = [column[0] for column in cursor.description]
            
            # Procesar cada registro sin imprimir datos sensibles
            for row in rows:
                data = {column: getattr(row, column) for column in columns}
                # Usa los datos de manera segura en tu aplicación
                return data
        except pyodbc.Error as ex:
            print("Error al obtener datos:", ex)
        finally:
            cursor.close()
            conn.close()
    else:
        print("No se pudo establecer conexión a la base de datos.")
