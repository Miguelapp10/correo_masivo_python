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
import pandas as pd
import datetime as dt  # Para fechas (opcional)
import glob  # Para jalar todo los archivos en una carpeta
import os  # Para trabajar con rutas
import locale
import time

# Usuario
usuario = r''
# Ruta
ruta_plan = r'Planificación UM CTs - NO SHOW' 
Ruta_lima_plan = r'Planificación IL - 1.2.9. Tiendas de Lima - Correo Masivo' 
ruta_post = r'Devolucion_Postventa - Documentos' 

# 1. Rutas base
ruta_base_Plan = r'C:\\Users\\' + usuario + r'\\\\' + ruta_plan 
ruta_base_Lima_Plan = r'C:\\Users\\' + usuario + r'\\\\' + Ruta_lima_plan 
ruta_base_Post = r'C:\\Users\\' + usuario + r'\\\\' + ruta_post 
## lista Personal de recolección diaria
credentials_path =  ruta_base_Post + r'\\PRD_Tienda'
ruta_PRD_Provincia = ruta_base_Post + r'\\PRD_Tienda\\Lista_PRD_Provincia'
ruta_lista_PRD= ruta_base_Post + r'\\PRD_Tienda\\Lista_PRD'
ruta_lista_PICKUP = ruta_base_Post + r'\\PRD_Tienda\\Lista_STORE_PICKUP'
ruta_lista_LIMASUR = ruta_base_Post + r'\\PRD_Tienda\\Lista_LIMASUR'
ruta_consolidado_Provincia = ruta_base_Post + r'\\PRD_Tienda\\Lista_PRD_Provincia\\Consolidado_Provincia'

# Opciones
## fechas
# Establecer la configuración regional en español
locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')
fecha_hoy = pd.to_datetime('today').date() + pd.Timedelta(days=1)
#fecha_hoy = pd.to_datetime('2025-07-10').date()
## formatear fecha
formato_hoy = fecha_hoy.strftime('%d.%m.%y')
## Warnings
warnings.filterwarnings(action='ignore') # Eliminar mensajes de warning (no elimina mensajes de error)
# Configuración para deshabilitar la notación científica al mostrar el DataFrame
pd.set_option('display.float_format', None)
# Funciones
## Función para leer Excel más rápido
def read_excel(path: str, sheet_name: str) -> pd.DataFrame:
    buffer = StringIO()
    Xlsx2csv(path, outputencoding="utf-8", sheet_name=sheet_name).convert(buffer)
    buffer.seek(0)
    df = pd.read_csv(buffer)
    return df
#################################################################################################################################################
import os.path
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from googleapiclient.errors import HttpError
import io
import ssl
import requests

# Desactivar la verificación de certificados SSL
requests.packages.urllib3.disable_warnings(requests.packages.urllib3.exceptions.InsecureRequestWarning)
ssl._create_default_https_context = ssl._create_unverified_context

# Definir los alcances de la API
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets.readonly',
    'https://www.googleapis.com/auth/bigquery',
    'https://www.googleapis.com/auth/drive.readonly',
    'https://www.googleapis.com/auth/cloud-platform',
    'https://www.googleapis.com/auth/drive', 
    'https://www.googleapis.com/auth/drive.metadata.readonly'
]

# Función para descargar la hoja de Google Sheets
def download_sheet(service, sheet_id, output_filename):
    try:
        request = service.files().export_media(
            fileId=sheet_id,
            mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        file_data = io.BytesIO()
        downloader = MediaIoBaseDownload(file_data, request)
        done = False
        while done is False:
            status, done = downloader.next_chunk()

        with open(output_filename, 'wb') as f:
            f.write(file_data.getvalue())
        print(f"Google Sheets file '{sheet_id}' exported successfully as Excel (XLSX) format.")
    except HttpError as error:
        print(f"An error occurred while downloading sheet '{sheet_id}': {error}")

# Función principal para autenticarse y ejecutar el script
def main():
    creds = None
    API_SERVICE_NAME = 'drive'
    API_VERSION = 'v3'

    # Cargar credenciales desde el archivo token.json
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)

    # Si las credenciales no son válidas, realizar el flujo de autenticación
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                "credentials.json", SCOPES
            )
            creds = flow.run_local_server(port=0)

        # Guardar las credenciales para el futuro uso
        with open("token.json", "w") as token:
            token.write(creds.to_json())

    try:
        # Construir el servicio de Google Drive
        service = build(API_SERVICE_NAME, API_VERSION, credentials=creds)

        # Descargar el segundo archivo
        sheet_id_2 = ''
        download_sheet(service, sheet_id_2, 'Dato_tienda_Dev_C&C.xlsx')

    except HttpError as error:
        print(f"An error occurred: {error}")

if __name__ == "__main__":
    main()

#################################################################################################################################################
import keyring
# Almacenar las credenciales
email= ''
contrasena = ''  # << Actualiza la contraseña si es necesario
# Guardar en keyring
keyring.set_password("sistema_smtp", email, contrasena)
# Recuperar las credenciales de forma segura
email_usuario = email
contrasena_usuario = keyring.get_password("sistema_smtp", email_usuario)
