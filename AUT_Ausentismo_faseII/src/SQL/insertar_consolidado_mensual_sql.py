import os
import sys
import pandas as pd
from datetime import datetime, timedelta

# 👇 Agregamos la raíz del proyecto al sys.path
proyecto_raiz = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.append(proyecto_raiz)

from src.Modules.consolidado_mensual import obtener_nombre_mes
from src.Fuji.connection import connection

def insert_consolidado_mensual_sql():
    """Insertar los datos de reporte consolidado mensual en la base de datos SQL"""
    
    año_actual = datetime.now().year
    ayer = datetime.now() - timedelta(days=1)
    # obtener el mes actual en mayusculas
    mes_actual = obtener_nombre_mes(ayer)

    try:
        print(f"Ruta del proyecto: {proyecto_raiz}")

        # ruta de la carpeta reportes
        carpeta_reportes = os.path.join(proyecto_raiz, 'src', 'Consolidado_Mensual')
        print(f"Ruta de la carpeta reportes: {carpeta_reportes}")

        # validamos que la carpeta exista
        if not os.path.exists(carpeta_reportes):
            print(f"La carpeta 'reportes' no existe en: {carpeta_reportes}")
            return

        ruta_reporte_mensual = os.path.join(
            carpeta_reportes,
            f'REPORTE_CONSOLIDADO_AUSENTISMO_{mes_actual}_{año_actual}.xlsx'
        )
        print(f"Ruta del archivo reporte mensual: {ruta_reporte_mensual}")

        if not os.path.exists(ruta_reporte_mensual):
            print(f"No se encontró el archivo de reporte mensual en: {ruta_reporte_mensual}")
            return
        
        # Cargar el archivo Excel con pandas
        print("Cargando el archivo Excel...")
        df_consolidado = pd.read_excel(ruta_reporte_mensual, sheet_name= "Consolidado", header=0)

        # asegurar que todas las columnas sean string
        df_consolidado = df_consolidado.astype(str)

        print("columnas leidas por pandas:")
        for i, col in enumerate(df_consolidado.columns):
            print(f"{i}: {repr(col)}")

        columnas_esperadas = ["ZONA","CENTRO_COSTOS", "OFICINA", "CEDULA", "NOMBRE", "MOTIVO_AUSENCIA", "DIAS", "FECHA_INICIAL", "FECHA_FINAL"]
        for col in columnas_esperadas:
            if col not in df_consolidado.columns:
                print(f"Error: No se encontró la columna '{col}' en el archivo Excel.")
                return False
            
        # Conexión a la base de datos SQL Server
        print("Conectando a la base de datos SQL Server...")
        conn = connection()
        if conn is None:
            print("No se pudo conectar a la base de datos.")
            return False

        cursor = conn.cursor()

        # validamos si la tabla consolidado_ausentismo existe
        cursor.execute("SELECT COUNT(*) FROM TblConsolidadoAusentismos")
        resultado = cursor.fetchone()
        if resultado is None:
            print("La tabla 'TblConsolidadoAusentismos' no existe en la base de datos.")
            return False

        # insertar los datos en la tabla consolidado_ausentismo
        for index, row in df_consolidado.iterrows():
            try:
                # Validar que no exista la misma combinación de CEDULA + FECHA_INICIAL
                cursor.execute("""
                    SELECT COUNT(*) FROM TblConsolidadoAusentismos 
                    WHERE CEDULA = ? AND FECHA_INICIAL = ?
                """, row['CEDULA'], row['FECHA_INICIAL'])
                
                existe = cursor.fetchone()[0]

                if existe == 0:
                    cursor.execute("""
                        INSERT INTO TblConsolidadoAusentismos (ZONA, CENTRO_COSTOS, OFICINA, CEDULA, NOMBRE, MOTIVO_AUSENCIA, DIAS, FECHA_INICIAL, FECHA_FINAL)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, 
                    row['ZONA'], 
                    row['CENTRO_COSTOS'], 
                    row['OFICINA'], 
                    row['CEDULA'], 
                    row['NOMBRE'], 
                    row['MOTIVO_AUSENCIA'],
                    row['DIAS'], 
                    row['FECHA_INICIAL'],
                    row['FECHA_FINAL']
                    )
                    print(f"Fila {index} insertada correctamente.")
                else:
                    print(f"Fila {index} con CEDULA {row['CEDULA']} y FECHA {row['FECHA_INICIAL']} ya existe, no se inserta.")
            
            except Exception as e:
                print(f"Error al insertar la fila {index} con cédula {row['CEDULA']}: {e}")

        #cierre después de recorrer todas las filas
        conn.commit()
        cursor.close()
        conn.close()



    

    except Exception as e:
        print(f"Error al insertar los datos de reporte consolidado mensual: {e}")


# if __name__ == "__main__":
#     insert_consolidado_mensual_sql()
