import os
import pandas as pd
from src.Emails.crear_correos import crearCorreos
from src.SQL.consultar_maestra import consultar_maestra_DB

def enviar_correo_zonas(zonas, dia_habil):

    # creamos un contador
    contador = 0

    # Traemos la maestra con centros de costo y correos
    data_maestra = consultar_maestra_DB()
    df = pd.DataFrame(data_maestra)
    centros_costo_maestra = df['CENTRO_COSTOS'].astype(str).str.replace(r'\.0$', '', regex=True).to_list()

    for zona in zonas:
        carpeta_zona = os.path.abspath(os.path.join(
            os.path.dirname(__file__), '..', 'Reportes_Ausentismos', f'Zona_{zona}_{dia_habil}'
        ))

        # Recorremos las carpetas de municipios dentro de la zona
        for carpeta_mpio in os.listdir(carpeta_zona):
            if carpeta_mpio.startswith("Municipio_"):
                municipio = carpeta_mpio.replace("Municipio_", "")

                ruta_excel_mpio = os.path.join(carpeta_zona, carpeta_mpio, f'REPORTE_AUSENTISMO_{municipio}_{dia_habil}.xlsx')

                if not os.path.exists(ruta_excel_mpio):
                    print(f"No existe archivo para {municipio}, se omite")
                    continue

                # Leemos el Excel del municipio
                df_mpio = pd.read_excel(ruta_excel_mpio)

                # Extraemos centros de costo de ese municipio
                centros_costo_mpio = df_mpio['CENTRO COSTOS'].astype(str).str.replace(r'\.0$', '', regex=True).to_list()

                # Buscamos correos en la maestra
                data_correos_mpio = ["jose.chaverra@gruporeditos.com"]
                for cc in centros_costo_mpio:
                    if cc in centros_costo_maestra:
                        correo = df.loc[
                            df['CENTRO_COSTOS'].astype(str).str.replace(r'\.0$', '', regex=True) == cc, 'CORREO'
                        ].to_list()[0]

                        if correo not in data_correos_mpio:
                            data_correos_mpio.append(correo)

                # Enviar correo
                if data_correos_mpio:
                    print(f" Enviando correo de {municipio} con archivo {ruta_excel_mpio}")
                    correo = crearCorreos(data_correos_mpio)
                    asunto, cuerpo = correo.preparar_correo(dia_habil,zona, municipio)
                    correo.enviar_correo(asunto, cuerpo, ruta_excel_mpio)

                    contador += 1
                else:
                    print(f"No hay correos para enviar en {municipio}")
    print(f"Correos enviados: {contador}")