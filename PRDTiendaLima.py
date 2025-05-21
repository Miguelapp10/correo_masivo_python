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
from UsuarioContraTienda import ruta_lista_PICKUP,usuario,ruta_post,ruta_base_Post,credentials_path,email_usuario ,contrasena_usuario, fecha_hoy,formato_hoy

### reporte diario de urbano diario
ruta_PICKUP = glob.glob(os.path.join(ruta_lista_PICKUP,('Consolidado STORE PICK-UP '+ formato_hoy + '.xlsx')))
ruta_PICKUP_ = pd.DataFrame()
ruta_PICKUP_ = []  # Initialize an empty list instead of a DataFrame
x = pd.DataFrame()
for i in range(len(ruta_PICKUP)):
    x = pd.read_excel(ruta_PICKUP[i], 'CONSOLIDADO TOTAL', dtype={"numero_order": str,"num_rastreo": str,"Lpn_compra":str})#, 
    ruta_PICKUP_.append(x)
ruta_PICKUP_ = pd.concat(ruta_PICKUP_, ignore_index=True)

### Datos de los tienda 

ruta_TIENDAS = glob.glob(os.path.join(credentials_path,"Dato_tienda_Dev_C&C.xlsx"))
ruta_TIENDAS_ = pd.DataFrame()
ruta_TIENDAS_ = []  # Initialize an empty list instead of a DataFrame
x = pd.DataFrame()
for i in range(len(ruta_TIENDAS)):
    x = pd.read_excel(ruta_TIENDAS[i], 'Datos',dtype={"CELULAR": str,"CELULAR_SEGUNDO_CONTACTO":str,"CELULAR_2":str})
    ruta_TIENDAS_.append(x)
ruta_TIENDAS_ = pd.concat(ruta_TIENDAS_, ignore_index=True)
########################################################################################################################################################
########################################################################################################################################################

# Quitar inconsistencias en nombres de columnas
ruta_PICKUP = ruta_PICKUP_.rename(columns=lambda x: x.strip()).copy()
ruta_TIENDAS = ruta_TIENDAS_.rename(columns=lambda x: x.strip()).copy()

PICKUP_TIENDA = pd.merge(left=ruta_PICKUP ,right= ruta_TIENDAS , how='left', left_on='ID Tienda', right_on='COD_TIENDA' )

# Usando la función format()
PICKUP_TIENDA['num_rastreo'] = PICKUP_TIENDA['num_rastreo'].astype(str)
PICKUP_TIENDA['Lpn_compra'] = PICKUP_TIENDA['Lpn_compra'].astype(str)
PICKUP_TIENDA['CORREOS'] = PICKUP_TIENDA['CORREOS'].astype(str) # Convert to string type
print(PICKUP_TIENDA)

# Lista de las tiendas que quieres filtrar 'Sodimac Primavera', 'Tottus El Agustino', 'Tottus Miraflores' 
#filtrar_tiendas = [ 'Tottus Los Olivos']
# Filtrar las tiendas
#PICKUP_TIENDA = PICKUP_TIENDA[PICKUP_TIENDA['Nombre_Tienda'].isin(filtrar_tiendas)]
# Eliminar duplicados en la columna "Nombre"
#PICKUP_TIENDA = PICKUP_TIENDA.drop_duplicates(subset=['numero_order'])
# Guardar el DataFrame como un archivo Excel
with pd.ExcelWriter(os.path.join(credentials_path, f"PRDTiendaLima_{fecha_hoy}.xlsx"), engine='xlsxwriter') as writer:
    # Guarda los resultados de CNCRD en la primera pestaña
    PICKUP_TIENDA.to_excel(writer, sheet_name='Dato', index=False)
########################################################################################################################################################
########################################################################################################################################################
import time
import datetime
import pandas as pd
##import pyautogui
import smtplib
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import math  # Necesario para verificar NaN
import os

# Obtener la fecha actual y calcular la semana correspondiente
#fecha_hoy = datetime.date.today()
week_number = fecha_hoy.strftime("%V")
dia_semana = fecha_hoy.strftime('%A')
#archivo_registro = "Registro_Correos.xlsx"
# Crear un DataFrame vacío para almacenar el registro de envío de correos electrónicos
registro_correos = pd.DataFrame(columns=["Tienda", "Resultado"])

#######################################################################################################################################################
#######################################################################################################################################################

# Iterar sobre las tiendas únicas en el DataFrame
for tienda in PICKUP_TIENDA['Nombre_Tienda'].unique():
    df_tienda = PICKUP_TIENDA[PICKUP_TIENDA['Nombre_Tienda'] == tienda]
    if df_tienda.empty:
        print(f"No hay datos para la tienda {tienda}. No se enviará el correo.")
        continue
    correo_tienda = df_tienda['CORREOS'].iloc[0]
    ID_Tienda = df_tienda['ID Tienda'].iloc[0]
    if pd.isna(correo_tienda):
        print(f"La dirección de correo electrónico para la tienda {tienda} es nula (NaN). No se enviará el correo.")
        continue
    # Create a pivot table
    pivot_table = pd.pivot_table(df_tienda,
                             values=['numero_order'],
                             index=['Nombre_Tienda'],
                             columns=['Tipo Devolución','BU'],
                             aggfunc={'numero_order': 'count'}, #'count'
                             #margins=True,
                             #margins_name='Total'
                             ).rename(columns={'Tipo Devolución': 'Tipo'}
                                      ).rename_axis(columns={'Nombre_Tienda': 'Tienda'})
    # Rellenar los valores NaN con ''
    pivot_table.fillna('', inplace=True)
    # Create HTML representation of the styled pivot table
    html_table = pivot_table.to_html(classes='styled-table', escape=False)
    # Filtrado y columnas necesarias
    filtro_cols = ['numero_order', 'num_rastreo', 'Lpn_compra', 'ID Tienda', 'Nombre_Tienda', 'BU', 'Tipo Devolución', 'Vehiculo']
    Devo_df = df_tienda[df_tienda['Tipo Devolución'].str.contains('', na=False)][filtro_cols]
    NoShow_df = df_tienda[df_tienda['Tipo Devolución'].str.contains('', na=False)][filtro_cols]
    # Definir el nombre del archivo Excel con el formato deseado
    nombre_archivo = f"Devoluciones_{ID_Tienda}_W{week_number}.xlsx"
    # Crear archivo Excel
    with pd.ExcelWriter(nombre_archivo, engine='xlsxwriter') as writer:
    # Escribir hojas
        Devo_df.to_excel(writer, sheet_name='Devolucion', index=False)
        NoShow_df.to_excel(writer, sheet_name='NoShow', index=False)
        workbook = writer.book
        header_format = workbook.add_format({'bg_color': '#a4d41e', 'bold': True})

        for df, sheet_name in [(Devo_df, 'Devolucion'), (NoShow_df, 'NoShow')]:
            sheet = writer.sheets[sheet_name]
            # Formato encabezado
            for col_num, col_name in enumerate(df.columns):
                sheet.write(0, col_num, col_name, header_format)
            # Autofiltro y tamaño de columna
            sheet.autofilter(0, 0, len(df.index), len(df.columns) - 1)
            sheet.set_column(0, len(df.columns) - 1, 30)
    # Validación de creación de archivo
    if not os.path.exists(nombre_archivo):
        print(f"El archivo {nombre_archivo} no fue creado correctamente.")
        registro_correos = pd.concat([registro_correos, pd.DataFrame({
            "Tienda": tienda,
            "Resultado": "Error: Archivo no creado"
        }, index=[0])])
        continue
    
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
    destinatarios_to = correo_tienda.split(',')

    # Concatenar las direcciones de correo electrónico en una cadena separada por comas
    destinatarios_str_to = ', '.join(destinatarios_to)
    
    # Concatenar las direcciones de correo electrónico en una cadena separada por comas
    destinatarios_CC = ['']

    # Concatenar las direcciones de correo electrónico en una cadena separada por comas
    destinatarios_str_CC = ', '.join(destinatarios_CC)
    try:
        # Configurar el mensaje
        msg = MIMEMultipart()
        msg['From'] = ''
        msg['To'] =    destinatarios_str_to ##'mpazos@falabella.com', 
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
        #server.sendmail(email_usuario,destinatarios_str_to.split(',') + destinatarios_str_CC.split(','), msg.as_string()) 
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
with pd.ExcelWriter(os.path.join(credentials_path, f"RegistroCorreoTiendaLima_{fecha_hoy}.xlsx"), engine='xlsxwriter') as writer:
    # Guarda los resultados de CNCRD en la primera pestaña
    registro_correos.to_excel(writer, sheet_name='Datos', index=False)
# Mensaje de confirmación
print(f"Se ha guardado el registro de envío de correos electrónicos: RegistroCorreoTiendaLima_{fecha_hoy}.xlsx")
########################################################################################################################################################
########################################################################################################################################################

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
    msg['From'] = '' #email_usuario
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
        #server.sendmail(email_usuario, ', '.join(destinatarios_to).split(',') + ', '.join(destinatarios_cc).split(','), msg.as_string())

def procesar_envio(nombre, df_filtro, filtros, destinatarios_to, destinatarios_cc, fecha_hoy, email_usuario, contrasena_usuario):
    df = ruta_PICKUP[ruta_PICKUP['Nombre_Tienda'].str.contains(filtros[0]) | ruta_PICKUP['Nombre_Tienda'].str.contains(filtros[1])]
    df = df[['numero_order','num_rastreo','Lpn_compra','sku_simple','sku_name','item_shipment_method','ID Tienda',
             'Nombre_Tienda','BU','Tipo Devolución','Vehiculo']] 
    if not df.empty:
        df = df.drop_duplicates(subset=['numero_order'])
        pivot_table = pd.pivot_table(
            df,
            values=['numero_order'],
            index=['Nombre_Tienda'],
            columns=['Tipo Devolución','BU'],
            aggfunc={'numero_order': 'count'}, #'count'
            margins=True,
            margins_name='Total'
        ).rename(columns={'Tipo Devolución': 'Tipo'}
                ).rename_axis(columns={'Nombre_Tienda': 'Tienda'}).fillna('')
        
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
# ------
procesar_envio(
    nombre="",
    df_filtro=ruta_PICKUP,
    filtros=[],
    destinatarios_to=[''],
    destinatarios_cc=[],
    fecha_hoy=fecha_hoy,
    email_usuario=email_usuario,
    contrasena_usuario=contrasena_usuario
)


