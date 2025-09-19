import os
import re
import shutil
from datetime import datetime
import datetime as dt
from src.Modules.consolidado_mensual import MESES_DIC


def eliminar_archivos_antiguos(dia_habil, ruta_reportes):
        
        fecha_valida = dia_habil
        if not fecha_valida:
            return

        patrones = [
            re.compile(r"Ausentismos_SIN_JUSTIFICACION_GENERAL - (\d{4}-\d{2}-\d{2})\.xlsx"),
            re.compile(r"REPORTE_CONSOLIDADO_AUSENTISMO_DIARIO_(\d{4}-\d{2}-\d{2})\.xlsx")
        ]

        for archivo in os.listdir(ruta_reportes):
            for p in patrones:
                match = p.match(archivo)
                if match:
                    fecha_archivo = match.group(1)
                    if fecha_archivo != fecha_valida:
                        ruta_completa = os.path.join(ruta_reportes, archivo)
                        print(f"Eliminando archivo antiguo: {ruta_completa}")
                        os.remove(ruta_completa)
                    break  # si ya hizo match con un patrón, no seguir probando con otros


        # Regex que captura la fecha dentro de la carpeta Zona_...
        patron = re.compile(r"Zona.*?(\d{4}-\d{2}-\d{2})")
        for nombre in os.listdir(ruta_reportes):
            ruta_completa = os.path.join(ruta_reportes, nombre)

            match = patron.search(nombre)   # usar search en lugar de match
            if match:
                fecha_archivo = match.group(1)
                print(f"Detectado: {nombre} → fecha {fecha_archivo}")

                if fecha_archivo != fecha_valida:
                    try:
                        if os.path.isdir(ruta_completa):
                            print(f"Eliminando carpeta antigua: {ruta_completa}")
                            shutil.rmtree(ruta_completa)
                    except Exception as e:
                        print(f"Error al eliminar {ruta_completa}: {e}")

        # 🔹 Nuevo bloque: limpiar CONSOLIDADO_MENSUAL por mes/año
        try:
            hoy = dt.date.today()
            mes_actual = hoy.month
            anio_actual = hoy.year

            # Calcular el mes de hace 2 meses
            mes_prev = mes_actual - 2
            anio_prev = anio_actual
            if mes_prev <= 0:
                mes_prev += 12
                anio_prev -= 1

            mes_a_eliminar = (MESES_DIC[mes_prev], anio_prev)

            # Buscar en nombres de consolidado mensual
            for archivo in os.listdir(ruta_reportes):
                if archivo.startswith("REPORTE_CONSOLIDADO_AUSENTISMO_") and archivo.endswith(".xlsx"):
                    patron = f"{mes_a_eliminar[0]}_{mes_a_eliminar[1]}"
                    if patron in archivo:
                        ruta_archivo = os.path.join(ruta_reportes, archivo)
                        os.remove(ruta_archivo)
                        print(f"Eliminando consolidado mensual antiguo: {ruta_archivo}")

        except Exception as e:
            print(f"Error al limpiar consolidados mensuales: {e}")
