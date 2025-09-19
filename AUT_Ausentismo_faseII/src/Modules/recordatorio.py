import os
from openpyxl import load_workbook
from datetime import datetime
from src.Modules.procesos import Cruce_datos
from src.Emails.crear_correos import crearCorreos
import re
from openpyxl import Workbook


def filtrar_zonas_ausentismos():

    # correo = crearCorreos()
    cruce = Cruce_datos()

    datos_ausentismo = cruce.extraer_datos_ausentismo_sin_justificacion()
    # datos_maestra = cruce.extraer_datos_maestra()
    dia_habil = cruce.obtener_ultimo_dia_anterior()



    zonas = set()
    resultados = {}  # <-- guardar aquí municipios y rutas

    for fila in datos_ausentismo:
        zonas.add(fila[0])
        # print(f"Zona encontrada: {fila[0]}")

    print(f"Total de zonas encontradas: {len(zonas)}")
    print(f"Total de filas en datos_ausentismo: {len(datos_ausentismo)}")

    for zona in zonas:
        zonas_filtradas = []
        for fila in datos_ausentismo:
            if fila[0] == zona:
                zonas_filtradas.append(fila)
                # print(f"Fila encontrada para la zona {zona}: {fila}")

        carpeta_zona = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'Reportes_Ausentismos', f'Zona_{zona}_{dia_habil}'))
        if not os.path.exists(carpeta_zona):
            os.makedirs(carpeta_zona)
            #print(f"Carpeta creada: {carpeta_zona}")

        municipios = set()
        for fila in zonas_filtradas:
            municipios.add(fila[2])
            # print(f"Municipio encontrado: {fila[2]}")

        resultados[zona] = {}

        for municipio in municipios:
            municipio_filtrado = []
            for fila in zonas_filtradas:
                if fila[2] == municipio:
                    municipio_filtrado.append(fila)
                    #print(f"Fila encontrada para el municipio {municipio}: {fila}")

            if municipio_filtrado:
                # limpiamos el nombre del municipio para que no tenga caracteres especiales
                municipio = re.sub(r"[^A-Za-z0-9 _-]", "", municipio)
                workbook = Workbook()
                hoja = workbook.active
                hoja.title = f"{municipio}"

                headers = ['ZONA', 'CENTRO COSTOS', 'OFICINA', 'CEDULA', 'NOMBRE(completo)', 
                        'MOTIVO DE AUSENCIA', 'DIAS', 'FECHA INICIAL', 'FECHA FINAL']
                hoja.append(headers)

                for fila in municipio_filtrado:
                    hoja.append(fila)

                cruce.estilo_formato_excel(hoja)

                carpeta_municipio = os.path.join(carpeta_zona, f'Municipio_{municipio}')
                if not os.path.exists(carpeta_municipio):
                    os.makedirs(carpeta_municipio)
                    #print(f"Carpeta creada: {carpeta_municipio}")

                ruta_guardado = os.path.join(carpeta_municipio, f'AUSENTISMO_{municipio}_{dia_habil}.xlsx')
                workbook.save(ruta_guardado)
                #print(f"Archivo guardado para el municipio {municipio} en {ruta_guardado}")

                resultados[zona][municipio] = ruta_guardado  # <-- guardamos ruta del archivo

    return resultados












































































































