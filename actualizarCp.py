import jpype
jpype.startJVM()
from asposecells.api import Workbook
import pandas as pd
import openpyxl
from sqlalchemy import create_engine
import sqlalchemy
import os
import datetime
import psycopg2
import subprocess

subprocess.run(['python','descarga.py'])

archivo_xls = Workbook('Sonora.xls')
archivo_xls.save("Sonora.xlsx")
jpype.shutdownJVM()
os.remove('Sonora.xls')

workbook = openpyxl.load_workbook('Sonora.xlsx')

# Obtener la hoja que se desea eliminar
Hoja1 = workbook['Nota']
Hoja2 = workbook['Evaluation Warning']

# Eliminar la hoja
workbook.remove(Hoja1)
workbook.remove(Hoja2)

# Genera un nombre único para el archivo
now = datetime.datetime.now()
timestamp = now.strftime('%Y-%m-%d')
filename = f'RegistrosBD_{timestamp}.xlsx'
# Guarda el archivo con el nuevo nombre
workbook.save('Sonora.xlsx')
# Cierra el archivo
workbook.close()
#Eliminacion de columnas
excel_file = pd.read_excel('Sonora.xlsx')
excel_file.drop(columns=['d_CP', 'c_estado','c_oficina','c_CP','c_tipo_asenta','c_mnpio','id_asenta_cpcons','c_cve_ciudad'], inplace =True) 
#Creacion Archivo limpio
sheet_name = "Sonora"
excel_file.to_excel("Sonora.xlsx",sheet_name=sheet_name,index = False)
engine = create_engine('postgresql://postgres:%40Hleyva21@localhost:5432/sac')
query = '''
    SELECT
        settlements.postal_code as d_codigo,
        settlements.name as d_asenta,
        initcap(type_settlements.name) AS d_tipo_asenta,
        municipalities.name AS D_mnpio,
        states.name AS d_estado,
        localities.name AS d_ciudad,
        settlements.ambit as d_zona
    FROM cat.settlements
    LEFT JOIN cat.states ON settlements.state_id = states.id
    LEFT JOIN cat.municipalities ON settlements.municipality_id = municipalities.id
    LEFT JOIN cat.localities ON settlements.locality_id = localities.id
    LEFT JOIN cat.type_settlements ON settlements.type_settlement_id = type_settlements.id;
'''
df_bd = pd.read_sql(query,engine)
df_bd.to_excel(filename,index=False) 

# Lee los datos desde los archivos Excel
df_sonora = pd.read_excel('Sonora.xlsx')
df_registros_bd = pd.read_excel(filename)

# Combina los DataFrames usando 'merge' y 'indicator'
merged_df = pd.merge(df_sonora, df_registros_bd, on=['d_codigo', 'd_asenta'], how='left', indicator=True)

# Filtra los registros de Sonora que no están en registros_bd
nuevos_cp = merged_df.loc[merged_df['_merge'] == 'left_only'].drop('_merge', axis=1)

# Elimina columnas específicas
columnas_a_eliminar = ['d_tipo_asenta_y', 'd_mnpio', 'd_estado_y', 'd_ciudad_y', 'd_zona_y']
nuevos_cp = nuevos_cp.drop(columnas_a_eliminar, axis=1)

# Quita los sufijos "_x" e "_y" de los nombres de las columnas
nuevos_cp.columns = nuevos_cp.columns.str.replace('_x', '')

# Renombra la columna 'D_mnpio' a 'd_mnpio'
nuevos_cp = nuevos_cp.rename(columns={'D_mnpio': 'd_mnpio'})

nuevos_cp.to_sql('nuevos', con=engine,schema='migrations',if_exists='append', index=False,dtype={'d_codigo': sqlalchemy.types.VARCHAR(5), 'd_asenta': sqlalchemy.types.VARCHAR(250), 'd_zona': sqlalchemy.types.VARCHAR(250)})

os.remove('Sonora.xlsx')
os.remove(filename)

conn = psycopg2.connect(
host="localhost",
database="sac",
user="postgres",
password="@Hleyva21"
)

# Creación de un cursor
cur = conn.cursor()

# Ejecución del procedimiento almacenado
cur.execute("CALL migrations.new_cp()")

# Confirmación de los cambios
conn.commit()

# Cierre del cursor y la conexión
cur.close()
conn.close()

