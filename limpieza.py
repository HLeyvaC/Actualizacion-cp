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

archivo_xls = Workbook('D:\Gob\Actualizacion cp\Sonora.xls')##Tu Path
archivo_xls.save("Sonora.xlsx")
jpype.shutdownJVM()
os.remove('D:\Gob\Actualizacion cp\Sonora.xls')

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
filename = f'Archivo_{timestamp}.xlsx'
# Guarda el archivo con el nuevo nombre
workbook.save(filename)
# Cierra el archivo
workbook.close()
#Eliminacion de columnas
excel_file = pd.read_excel(filename)
excel_file.drop(columns=['d_CP', 'c_estado','c_oficina','c_CP','c_tipo_asenta','c_mnpio','id_asenta_cpcons','c_cve_ciudad'], inplace =True) 
#Creacion Archivo limpio
sheet_name = "Sonora"
excel_file.to_excel(filename,sheet_name=sheet_name,index = False)

if os.path.exists("registro.txt"):
    with open("registro.txt","r") as archivo:
        nombre = archivo.read()
    
    df1 = pd.read_excel(nombre)
    df2 = pd.read_excel(filename)

    os.remove('cp_nuevos.xlsx')
    concat_df = pd.concat([df1, df2], axis=0, ignore_index=True)
    unique_df = concat_df.drop_duplicates(keep=False)
    unique_df.rename(columns={'D_mnpio': 'd_mnpio'}, inplace=True)
    unique_df.to_excel('cp_nuevos.xlsx', index=False)
    os.remove(nombre)
    engine = create_engine('postgresql://postgres:%40Hleyva21@localhost:5432/sac')
    unique_df.to_sql('nuevos', con=engine,schema='cat',if_exists='append', index=False,dtype={'d_codigo': sqlalchemy.types.VARCHAR(5), 'd_asenta': sqlalchemy.types.VARCHAR(250), 'd_zona': sqlalchemy.types.VARCHAR(250)})

    conn = psycopg2.connect(
        host="localhost",
        database="sac",
        user="postgres",
        password="@Hleyva21"
    )

    # Creación de un cursor
    cur = conn.cursor()

    # Ejecución del procedimiento almacenado
    cur.callproc('cat.new_cp')

    # Confirmación de los cambios
    conn.commit()

    # Cierre del cursor y la conexión
    cur.close()
    conn.close()

    with open("registro.txt","w") as file:
        file.write(filename)

else:
    df = pd.read_excel(filename)
    engine = create_engine('postgresql://postgres:%40Hleyva21@localhost:5432/sac')
    df.to_sql('nuevos', con=engine,schema='cat',if_exists='append', index=False,dtype={'d_codigo': sqlalchemy.types.VARCHAR(5), 'd_asenta': sqlalchemy.types.VARCHAR(250), 'd_zona': sqlalchemy.types.VARCHAR(250)})

    conn = psycopg2.connect(
        host="localhost",
        database="sac",
        user="postgres",
        password="@Hleyva21"
    )

    # Creación de un cursor
    cur = conn.cursor()
    
    # Ejecución del procedimiento almacenado
    cur.callproc('cat.new_cp')
    
    # Confirmación de los cambios
    conn.commit()

    # Cierre del cursor y la conexión
    cur.close()
    conn.close()

    with open("registro.txt","w") as file:
        file.write(filename)