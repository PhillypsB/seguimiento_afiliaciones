import os
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from datetime import datetime
import time
from prettytable import PrettyTable

def limpiar_pantalla():
    print("\033c", end="")

def obtener_estadisticas(df, ruta_csv, ruta_excel, tiempo_lectura_csv, tiempo_escritura_excel):
    try:
        tamano_csv = os.path.getsize(ruta_csv) / (1024 * 1024)
        filas_csv = df.shape[0]

        tamano_excel = os.path.getsize(ruta_excel) / (1024 * 1024)
        filas_excel = pd.read_excel(ruta_excel, sheet_name='Resultados').shape[0]

        print("\nEstadísticas:")
        print(f"Tamaño del archivo CSV original: {tamano_csv:.2f} MB")
        print(f"Número de filas en el archivo CSV original: {filas_csv:,}")
        print(f"Tamaño del archivo Excel generado: {tamano_excel:.2f} MB")
        print(f"Número de filas en el archivo Excel: {filas_excel:,}")

    except Exception as e:
        print(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - Error al obtener estadísticas: {e}")

def mostrar_tiempo_transcurrido(start_time):
    elapsed_time = time.time() - start_time
    minutes, seconds = divmod(elapsed_time, 60)
    
    if minutes > 0:
        print(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - Tiempo de ejecución: {int(minutes)} minutos y {seconds:.1f} segundos")
    else:
        print(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - Tiempo de ejecución: {seconds:.1f} segundos")

def leer_csv(ruta_csv):
    try:
        global start_time
        start_time = time.time()
        df = pd.read_csv(ruta_csv, encoding="latin-1", delimiter=";")
        tiempo_lectura_csv = time.time() - start_time
        print(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - Tiempo de lectura del CSV con codificación latin-1: {tiempo_lectura_csv:.1f} segundos")
        return df, tiempo_lectura_csv
    except Exception as e:
        print(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - Error al leer el archivo CSV: {e}")
        return None, None

def escribir_excel(df, ruta_excel):
    try:
        with pd.ExcelWriter(ruta_excel, engine='openpyxl') as writer:
            writer.book = Workbook()
            df.to_excel(writer, index=False, sheet_name='Resultados')
            writer.save()
        tiempo_escritura_excel = time.time() - start_time
        print(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - Tiempo de escritura del Excel: {tiempo_escritura_excel:.1f} segundos")
        return tiempo_escritura_excel
    except Exception as e:
        print(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - Error al escribir en el archivo Excel: {e}")
        return None

def leer_excel(ruta_excel):
    try:
        df_resultados = pd.read_excel(ruta_excel, sheet_name='Resultados')
        print(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - Tiempos de lectura del nuevo Excel: {time.time() - start_time:.1f} segundos")
        return df_resultados
    except Exception as e:
        print(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - Error al leer el archivo Excel: {e}")
        return None

def mostrar_resultados(df_resultados):
    if df_resultados is not None:
        pretty_table = PrettyTable()
        pretty_table.field_names = df_resultados.columns[:3]
        for _, row in df_resultados.iloc[:5, :3].iterrows():
            pretty_table.add_row(row)
        print(pretty_table)

def main():
    limpiar_pantalla()
    print(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - LIMPIANDO TABLA DE RESULTADOS")

    ruta_csv = r'G:\\BRAVO\\DATA\\sql.csv'
    ruta_excel = r'G:\\BRAVO\\RESULT\\nuevo_excel.xlsx'

    df, tiempo_lectura_csv = leer_csv(ruta_csv)

    if df is not None:
        tiempo_escritura_excel = escribir_excel(df, ruta_excel)
        mostrar_tiempo_transcurrido(start_time)

        print(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - LEYENDO EXCEL Y MOSTRANDO RESULTADOS")

        df_resultados = leer_excel(ruta_excel)
        mostrar_resultados(df_resultados)

        obtener_estadisticas(df, ruta_csv, ruta_excel, tiempo_lectura_csv, tiempo_escritura_excel)

if __name__ == "__main__":
    main()
