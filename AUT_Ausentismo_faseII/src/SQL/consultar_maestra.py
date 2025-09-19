import pyodbc
from src.Fuji.connection import connection

def consultar_maestra_DB():

    conn = connection()
    try:
        cursor = conn.cursor()
        query = f"""
        SELECT * FROM [DB_Automatizacion_Reportes].[dbo].[MAESTRA];
        """
        cursor.execute(query)

        datos = cursor.fetchall()
        columnas = [col[0] for col in cursor.description]
        data_maestra = [dict(zip(columnas, fila)) for fila in datos]

        # imprimimos los resultados
        print(f"Se han encontrado {len(data_maestra)} registros en la DB maestra.")
        # for fila in data_maestra:
        #     print(fila)

        return data_maestra

    except pyodbc.Error as ex:
        print(f"Error al extraer datos de la tabla: {ex}")
        return []

    finally:
        cursor.close()
        conn.close()

