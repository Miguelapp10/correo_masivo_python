import pandas as pd
import time
import datetime as dt  # Para fechas (opcional)
import glob  # Para jalar todo los archivos en una carpeta
import os  # Para trabajar con rutas
import locale
import warnings  # Para evitar que salgan errores en formatos de archivo (no altera el producto)
from datetime import datetime, timedelta # Para fechas (opcional)
from io import StringIO  # Usado para definir función nueva
from pathlib import Path
import numpy as np  # Para operaciones matemáticas
import xlsxwriter  # Funcionalidad para trabajar con archivos Excel
from fpdf import FPDF  # Para crear archivos PDF
from pandas import ExcelWriter  # Para exportar tabla a Excel
from xlsx2csv import Xlsx2csv  # Usado para definir función nueva
from decimal import Decimal
import unidecode
from UsuarioContraTienda import ruta_base_Lima_Plan,ruta_consolidado_Provincia,fecha_hoy,usuario,ruta_post,ruta_base_Post,credentials_path,email_usuario ,ruta_consolidado_Provincia,contrasena_usuario, fecha_hoy,formato_hoy
##ruta_base_Plan
### reporte DEVOLUCION Y NO SHOW
ruta_SP_Lima = glob.glob(os.path.join( ruta_base_Lima_Plan,('Tienda de Lima - Store Pick Up.xlsx')))
ruta_SP_Lima_ = pd.DataFrame()
ruta_SP_Lima_ = []  # Initialize an empty list instead of a DataFrame
x = pd.DataFrame()
for i in range(len(ruta_SP_Lima)):
    x = pd.read_excel(ruta_SP_Lima[i], 'Hoja1', dtype={"RASTREO": str,"ORDER_NUMBER": str})
    ruta_SP_Lima_.append(x)
ruta_SP_Lima_ = pd.concat(ruta_SP_Lima_, ignore_index=True)
# Convert 'fecha' column to datetime and filter rows where 'fecha' is less than today
ruta_SP_Lima_['FECHA PICKUP'] = pd.to_datetime(ruta_SP_Lima_['FECHA PICKUP'])
ruta_SP_Lima_ = ruta_SP_Lima_[ruta_SP_Lima_['FECHA PICKUP'] == pd.to_datetime(fecha_hoy)]

### Datos de los tienda 
ruta_TIENDAS = glob.glob(os.path.join(credentials_path,"Dato_tienda_Dev_C&C.xlsx"))
ruta_TIENDAS_ = pd.DataFrame()
ruta_TIENDAS_ = []  # Initialize an empty list instead of a DataFrame
x = pd.DataFrame()
for i in range(len(ruta_TIENDAS)):
    x = pd.read_excel(ruta_TIENDAS[i], 'Datos',dtype={"CELULAR": str,"CELULAR_SEGUNDO_CONTACTO":str,"CELULAR_2":str,'Bu_tienda':str})
    ruta_TIENDAS_.append(x)
ruta_TIENDAS_ = pd.concat(ruta_TIENDAS_, ignore_index=True)

# Quitar inconsistencias en nombres de columnas

ruta_SP_Lima = ruta_SP_Lima_.rename(columns=lambda x: x.strip()).copy()
ruta_TIENDAS = ruta_TIENDAS_.rename(columns=lambda x: x.strip()).copy()

# Usando la función format()
ruta_SP_Lima['RASTREO'] = ruta_SP_Lima['RASTREO'].astype(str)
ruta_SP_Lima['ORDER_NUMBER'] = ruta_SP_Lima['ORDER_NUMBER'].astype(str)
ruta_SP_Lima['FLUJO'] = ruta_SP_Lima['FLUJO'].astype(str).str.upper()
ruta_SP_Lima['TIENDA_ORIGEN'] = ruta_SP_Lima['TIENDA_ORIGEN'].astype(str).str.upper() 
ruta_TIENDAS['Bu_tienda'] = ruta_TIENDAS['Bu_tienda'].astype(str).str.upper()


SP_LIMA_TIENDA = pd.merge(left=ruta_SP_Lima ,right= ruta_TIENDAS , how='left', left_on='ID_TIENDA', right_on='COD_TIENDA')
                           #left_on='TIENDA_ORIGEN', right_on='Bu_tienda' ) 
SP_LIMA_TIENDA= SP_LIMA_TIENDA[['RASTREO','ORDER_NUMBER','RLO_ID','COD_TIENDA','TIENDA_ORIGEN', 'FLUJO',
                                'BU','PROVEEDOR','CORREOS']]


# Convertir a mayúsculas, eliminar tildes y reemplazar 'Ñ' por 'N'
SP_LIMA_TIENDA['TIENDA_ORIGEN'] = SP_LIMA_TIENDA['TIENDA_ORIGEN'].astype(str).str.upper().apply(lambda x: unidecode.unidecode(x.replace('Ñ', 'N')))
SP_LIMA_TIENDA['FLUJO'] = SP_LIMA_TIENDA['FLUJO'].astype(str).str.upper().apply(lambda x: unidecode.unidecode(x.replace('Ñ', 'N')))
with pd.ExcelWriter(os.path.join(credentials_path, f"PRDTiendaLima_{fecha_hoy}.xlsx"), engine='xlsxwriter') as writer:
    # Guarda los resultados de CNCRD en la primera pestaña
    SP_LIMA_TIENDA.to_excel(writer, sheet_name='Dato', index=False)
###########################################################################################################################################################################
###########################################################################################################################################################
import time
import datetime
import pandas as pd
import smtplib
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.header import Header
from email import encoders
import math  # Necesario para verificar NaN
import os
import requests
import urllib3

# Obtener la fecha actual y calcular la semana correspondiente
week_number = fecha_hoy.strftime("%V")
dia_semana = fecha_hoy.strftime('%A')
archivo_registro = "Registro_Correos.xlsx"

###############################################################################################################################################################
##################################################################################################################################################################
# Datos del correo
correo_origen = ""
correo_destino = []
asunto = "No esta actualizado reporte Consolidado store pickup"
mensaje = """
Hola,
No está actualizado el reporte Consolidado store pickup del dia de hoy.
Por favor validar.
Saludos,
Miguel Pazos
"""
# Desactivar advertencias de seguridad (opcional)
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
# Verificar si el DataFrame está vacío
if SP_LIMA_TIENDA.empty:
    print("No hay datos para enviar, se procederá a enviar un correo de alerta y mensaje a Teams.")

    # Enviar correo
    msg = MIMEMultipart()
    msg['From'] = correo_origen
    msg['To'] = ", ".join(correo_destino)
    msg['Subject'] = Header(asunto, 'utf-8')
    msg.attach(MIMEText(mensaje, 'plain', 'utf-8'))

    try:
        server = smtplib.SMTP('smtp.office365.com', 587)
        server.starttls()
        server.login(email_usuario, contrasena_usuario)  # Define estas variables con tus credenciales
        server.sendmail(correo_origen, correo_destino, msg.as_string())
        server.quit()
        print("Correo enviado exitosamente.")
    except Exception as e:
        print(f"Error al enviar correo: {e}")

    # Enviar mensaje a Teams via Power Automate
    url_flujo = ""
    payload = {
        "mensaje": " No está actualizado el reporte Consolidado store pickup del día de hoy.\nPor favor validar y actualizar el archivo. \nSaludos, "
    }
    res = requests.post(url_flujo, json=payload, verify=False)
    if res.status_code == 200:
        print("Mensaje enviado por Power Automate a Teams.")
    else:
        print(f"Error al enviar mensaje: {res.status_code}\n{res.text}")

else:
    print("Sí hay datos, no se envía correo ni mensaje a Teams.")

#############################################################################################################################################################################
##############################################################################################################################################################################
# Crear un DataFrame vacío para almacenar el registro de envío de correos electrónicos
registro_correos = pd.DataFrame(columns=["Tienda", "Resultado"])

# Iterar sobre las tiendas únicas en el DataFrame
for tienda in SP_LIMA_TIENDA['TIENDA_ORIGEN'].unique():
    df_tienda = SP_LIMA_TIENDA[SP_LIMA_TIENDA['TIENDA_ORIGEN']== tienda]
    if df_tienda.empty:
        print(f"No hay datos para la tienda {tienda}. No se enviará el correo.")
        continue
    correo_tienda = df_tienda['CORREOS'].iloc[0]
    origen_tienda = df_tienda['TIENDA_ORIGEN'].iloc[0]
    ID_Tienda = df_tienda['COD_TIENDA'].iloc[0]
    if pd.isna(correo_tienda):
        print(f"La dirección de correo electrónico para la tienda {tienda} es nula (NaN). No se enviará el correo.")
        continue
    
    # Create a pivot table
    pivot_table = pd.pivot_table(df_tienda,
                             values=['ORDER_NUMBER'],
                             index=['TIENDA_ORIGEN'],
                             columns=['FLUJO','BU'],
                             aggfunc={'ORDER_NUMBER': 'count'}, #'count'
                             #margins=True,
                             #margins_name='Total'
                             ).rename(columns={'FLUJO': 'Tipo'}
                                      ).rename_axis(columns={'TIENDA_ORIGEN': 'Tienda'})

    # Rellenar los valores NaN con ''
    pivot_table.fillna('', inplace=True)
    # Create HTML representation of the styled pivot table
    html_table = pivot_table.to_html(classes='styled-table', escape=False)
    # Filtrado y columnas necesarias
    filtro_cols = ['RASTREO','ORDER_NUMBER','RLO_ID','COD_TIENDA','TIENDA_ORIGEN', 'FLUJO',
                                'BU','PROVEEDOR']
    Devo_df = df_tienda[df_tienda['FLUJO'].str.contains('DEVOLUCIÓN|DEVOLUCION', na=False)][filtro_cols]
    NoShow_df = df_tienda[df_tienda['FLUJO'].str.contains('NO SHOW', na=False)][filtro_cols]
    Offline_df = df_tienda[df_tienda['FLUJO'].str.contains('OFFLINE', na=False)][filtro_cols]
    # Definir el nombre del archivo Excel con el formato deseado
    nombre_archivo = f"Devoluciones_{origen_tienda}_W{week_number}.xlsx"

    # Escribir los DataFrames en un archivo Excel
    with pd.ExcelWriter(nombre_archivo) as writer:
        # Escribir 'filtered_rlos_tienda' en una hoja llamada 'filtered_rlos_tienda'
        Devo_df.to_excel(writer, sheet_name='Devolucion', index=False)
        NoShow_df.to_excel(writer, sheet_name='NoShow', index=False)
        Offline_df.to_excel(writer, sheet_name='Offline', index=False)
        workbook = writer.book
        Devo_sheet = writer.sheets['Devolucion']
        NoShow_sheet = writer.sheets['NoShow']
        Offline_sheet = writer.sheets['Offline']

        # Formato para los títulos de fondo de color naranja
        header_format = workbook.add_format({'bg_color': '#a4d41e', 'bold': True})
        for df, sheet_name in [(Devo_df, 'Devolucion'), (NoShow_df, 'NoShow') , (Offline_df, 'Offline')]:
            sheet = writer.sheets[sheet_name]
            # Formato encabezado
            for col_num, col_name in enumerate(df.columns):
                sheet.write(0, col_num, col_name, header_format)
            # Autofiltro y tamaño de columna
            sheet.autofilter(0, 0, len(df.index), len(df.columns) - 1)
            sheet.set_column(0, len(df.columns) - 1, 30)     

    if not os.path.exists(nombre_archivo):
        print(f"El archivo {nombre_archivo} no fue creado correctamente.")
        registro_correos = pd.concat([registro_correos, pd.DataFrame({"Tienda": tienda, "Resultado": "Error: Archivo no creado"}, index=[0])])
        continue

    # Configuración
    asunto_base = f"RECOLECCIÓN FLOTA IBIS DEVOLUCIONES CLIENTE + NO SHOW (ABANDONO) // {tienda}// {fecha_hoy.year} - WEEK {week_number} - {fecha_hoy}"

    # Configuración de los servidores SMTP y los puertos
    MAX_RETRIES = 3
    # Intento de conexión SMTP
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            server = smtplib.SMTP('smtp.office365.com', 587)
            server.starttls()
            server.login(email_usuario, contrasena_usuario)
            break  # Si la conexión es exitosa, salir del bucle
        except smtplib.SMTPAuthenticationError as e:
            print(f"[Intento {attempt}] Error de autenticación: {e}")
            if attempt < MAX_RETRIES:
                time.sleep(5)
        except Exception as e:
            print(f"Error inesperado: {e}")
            break
    else:
        print("No se pudo establecer conexión SMTP después de varios intentos.")

   # Dividir la cadena en una lista de direcciones de correo electrónico
    destinatarios_to =  correo_tienda.split(',')

    # Concatenar las direcciones de correo electrónico en una cadena separada por comas
    destinatarios_str_to = ', '.join(destinatarios_to)
    
    # Concatenar las direcciones de correo electrónico en una cadena separada por comas
    destinatarios_CC = []
    # Concatenar las direcciones de correo electrónico en una cadena separada por comas
    destinatarios_str_CC = ', '.join(destinatarios_CC)
    try:
        # Configurar el mensaje
        msg = MIMEMultipart()
        msg['From'] = ''
        msg['To'] =    destinatarios_str_to 
        msg['Subject'] = asunto_base
        msg['CC'] = destinatarios_str_CC

        # Darle un nivel de importancia al correo electrónico (en este caso, Alto)
        msg.add_header('Importance', 'High')

        # Crear el cuerpo del correo en formato HTML
        mensaje_html = f"""
        <html>
        <head>
        <style>
        .styled-table th {{
        background-color: green;
        font-weight: bold;
        }}
        </style>
        </head>
        <body>
        <p>Estimado {tienda}!</p>
        <p>Se adjunta el detalle de los pedidos que serán recolectados el dia {dia_semana} bajo el nuevo flujo (IBIS).</p>
        <p>Tener en cuenta que las recolecciones serán tanto para pedidos por devolución cliente y no show.</p>

        <p><span style="font-size: 16px"><u><strong> Cantidad de devoluciones y/o Noshow de cada tienda</strong></u></span><br></p>
        {html_table}
        <p>Quedo atento</p>
        <p>Saludos</p>
        <p>Equipo Home Delivery</p>
        </body>
        </html>
        """

        # Adjuntar el cuerpo del correo en formato HTML
        msg.attach(MIMEText(mensaje_html, 'html'))

        # Adjuntar el archivo Excel al correo electrónico
        with open(nombre_archivo, "rb") as attachment:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f"attachment; filename= {nombre_archivo}")
            msg.attach(part)

        # Enviar el correo electrónico
        server.sendmail(email_usuario,destinatarios_str_to.split(',') + destinatarios_str_CC.split(','), msg.as_string()) 
        registro_correos = pd.concat([registro_correos, pd.DataFrame({"Tienda": tienda, "Resultado": "Correo enviado correctamente"}, index=[0])])
        print(f"Correo electrónico enviado correctamente a {tienda}")

    except smtplib.SMTPAuthenticationError as e:
        registro_correos = pd.concat([registro_correos, pd.DataFrame({"Tienda": tienda, "Resultado": f"Error de autenticación: {e}"}, index=[0])])
        print(f"Error de autenticación al enviar correo electrónico a {tienda}: {e}")    

    except smtplib.SMTPDataError as e:
        registro_correos = pd.concat([registro_correos, pd.DataFrame({"Tienda": tienda, "Resultado": f"Error al enviar correo electrónico: {e}"}, index=[0])])
        print(f"Error al enviar correo electrónico a {tienda}: {e}")

    except Exception as e:
        registro_correos = pd.concat([registro_correos, pd.DataFrame({"Tienda": tienda, "Resultado": f"Error al enviar correo electrónico: {e}"}, index=[0])])
        print(f"Error al enviar correo electrónico a {tienda}: {e}")
    finally:
        try:
            server.quit()
        except smtplib.SMTPServerDisconnected as e:
            print(f"Error al cerrar la conexión SMTP: {e}")
# Guardar el DataFrame como un archivo Excel
with pd.ExcelWriter(os.path.join(credentials_path, f"RegistroCorreoTiendaProvincia_{fecha_hoy}.xlsx"), engine='xlsxwriter') as writer:
    # Guarda los resultados de CNCRD en la primera pestaña
    registro_correos.to_excel(writer, sheet_name='Datos', index=False)
# Mensaje de confirmación
print(f"Se ha guardado el registro de envío de correos electrónicos: RegistroCorreoTiendaProvincia_{fecha_hoy}.xlsx")

################################################################################################################################################################
################################################################################################################################################################
def generar_archivo_excel(df, nombre_archivo, sheet_name):
    with pd.ExcelWriter(nombre_archivo) as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        workbook = writer.book
        sheet = writer.sheets[sheet_name]
        header_format = workbook.add_format({'bg_color': '#a4d41e', 'bold': True})
        for col_num, value in enumerate(df.columns.values):
            sheet.write(0, col_num, value, header_format)
        sheet.autofilter(0, 0, len(df.index), len(df.columns) - 1)
        sheet.set_column(0, len(df.columns) - 1, 30)

def enviar_correo(asunto, destinatarios_to, destinatarios_cc, html_body, archivo_adjunto, email_usuario, contrasena_usuario):
    msg = MIMEMultipart()
    msg['From'] = '' ##email_usuario
    msg['To'] = ', '.join(destinatarios_to)
    msg['Subject'] = asunto
    msg['CC'] = ', '.join(destinatarios_cc)
    msg.add_header('Importance', 'High')
    msg.attach(MIMEText(html_body, 'html'))

    with open(archivo_adjunto, "rb") as attachment:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f"attachment; filename= {archivo_adjunto}")
        msg.attach(part)

    with smtplib.SMTP('smtp.office365.com', 587) as server:
        server.starttls()
        server.login(email_usuario, contrasena_usuario)
        server.sendmail(email_usuario, ', '.join(destinatarios_to).split(',') + ', '.join(destinatarios_cc).split(','), msg.as_string())

def procesar_envio(nombre, df_filtro, filtros, destinatarios_to, destinatarios_cc, fecha_hoy, email_usuario, contrasena_usuario):
    df = SP_LIMA_TIENDA[SP_LIMA_TIENDA['TIENDA_ORIGEN'].str.contains(filtros[0]) | SP_LIMA_TIENDA['TIENDA_ORIGEN'].str.contains(filtros[1])]
    df = df[['RASTREO','ORDER_NUMBER','RLO_ID','COD_TIENDA','TIENDA_ORIGEN', 'FLUJO','BU','PROVEEDOR','CORREOS']]
    
    if not df.empty:
        df = df.drop_duplicates(subset=['ORDER_NUMBER'])
        pivot_table = pd.pivot_table(
            df,
            values=['ORDER_NUMBER'],
            index=['TIENDA_ORIGEN'],
            columns=['FLUJO','BU'],
            aggfunc={'ORDER_NUMBER': 'count'}, #'count'
            margins=True,
            margins_name='Total'
        ).rename(columns={'FLUJO': 'Tipo'}
                ).rename_axis(columns={'TIENDA_ORIGEN': 'Tienda'}).fillna('')
            
        html_table = pivot_table.to_html(classes='styled-table', escape=False)

        week_number = fecha_hoy.strftime("%V")
        archivo = f"Devoluciones_{nombre}_W{week_number}_{fecha_hoy}.xlsx"
        asunto = f"RECOLECCIÓN FLOTA IBIS_ DEVOLUCIONES CLIENTE + NO SHOW (ABANDONO) {nombre.upper()} // WEEK {week_number} - {fecha_hoy}"
        
        generar_archivo_excel(df, archivo, nombre)

        mensaje_html = f"""
        <html>
        <head>
        <style>
        .styled-table th {{ background-color: green; font-weight: bold; }}
        </style>
        </head>
        <body>
        <p>Estimado {nombre}!</p>
        <p>Se adjunta el detalle de los pedidos que serán recolectados mañana bajo el nuevo flujo (IBIS).</p>
        <p>Tener en cuenta que las recolecciones serán tanto para pedidos por devolución cliente y no show.</p>
        <p><strong><u>Cantidad de devoluciones y/o Noshow de cada tienda:</u></strong></p>
        {html_table}
        <p>Saludos cordiales,<br>Home Delivery</p>
        </body>
        </html>
        """
        
        enviar_correo(asunto, destinatarios_to, destinatarios_cc, mensaje_html, archivo, email_usuario, contrasena_usuario)
        print(f"Correo enviado a {nombre}")
    else:
        print(f"No se encontraron datos para {nombre}. No se enviará correo.")

# ====================== PARÁMETROS ========================
# --- Tot ---
procesar_envio(
    nombre="TOT",
    df_filtro=SP_LIMA_TIENDA,
    filtros=[""],
    destinatarios_to=[],
    destinatarios_cc=[],
    fecha_hoy=fecha_hoy,
    email_usuario=email_usuario,
    contrasena_usuario=contrasena_usuario
)

# --- Sodi ---
procesar_envio(
    nombre="SOD",
    df_filtro=SP_LIMA_TIENDA,
    filtros=[""],
    destinatarios_to=[],
    destinatarios_cc=[],
    fecha_hoy=fecha_hoy,
    email_usuario=email_usuario,
    contrasena_usuario=contrasena_usuario
)


