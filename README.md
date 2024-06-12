# Automatización de Reportes Diarios y Distribución de Correos Electrónicos

## Descripción General
Este script automatiza el proceso de descarga de reportes diarios desde Google Drive, procesamiento de datos desde archivos Excel, generación de reportes resumen y envío de estos reportes por correo electrónico a las tiendas correspondientes. Realiza tres tareas principales:

1. **Descargar Archivos desde Google Drive**
2. **Procesar Archivos Excel**
3. **Enviar Reportes Resumen por Correo Electrónico**

## Requisitos Previos
- Python 3.6 o superior
- Bibliotecas de Python requeridas:
  - `pandas`
  - `xlsxwriter`
  - `fpdf`
  - `google-auth`
  - `google-auth-oauthlib`
  - `google-auth-httplib2`
  - `google-api-python-client`
  - `smtplib`
  - `email`
  - `numpy`
  - `glob`
  - `warnings`
  - `xlsx2csv`

## Instrucciones de Configuración

1. **Configuración de la API de Google**:
   - Asegúrese de tener `credentials.json` para la API de Google en su directorio de trabajo.
   - Configure un ID de cliente OAuth 2.0 desde la Consola de Google Cloud.

2. **Configuración del Entorno**:
   - Cree un entorno virtual de Python y actívelo:
     ```sh
     python -m venv venv
     source venv/bin/activate  # En Windows: venv\Scripts\activate
     ```
   - Instale las bibliotecas requeridas:
     ```sh
     pip install pandas xlsxwriter fpdf google-auth google-auth-oauthlib google-auth-httplib2 google-api-python-client smtplib email numpy glob warnings xlsx2csv
     ```

3. **Estructura del Directorio**:
   - Asegúrese de que exista la siguiente estructura de directorios:
     ```
     C:\Users\<username>\Falabella\<ruta>
     ├── PRD_Tienda
     │   ├── Lista_PRD
     │   └── Lista_STORE_PICKUP
     ```

4. **Credenciales**:
   - Reemplace los marcadores de posición en el script con valores reales:
     - `usuario`: Su nombre de usuario.
     - `ruta`: Subdirectorio para Falabella.
     - Credenciales de correo electrónico y configuraciones SMTP.
     - IDs y URLs de los archivos de Google Sheets.

## Descripción del Script

### 1. Descargar Archivos desde Google Drive

La función `main` maneja la autenticación de Google Drive y descarga hojas de cálculo específicas de Google como archivos Excel:

```python
SCOPES = ["https://www.googleapis.com/auth/drive"]

def download_sheet(service, sheet_id, output_filename):
    # Función para descargar una hoja de cálculo de Google como un archivo Excel
    ...

def main():
    creds = None
    ...
    try:
        service = build(API_SERVICE_NAME, API_VERSION, credentials=creds)
        download_sheet(service, sheet_id_1, os.path.join(ruta_lista_PRD,'Personal Recoleccion diaria '+ str(fecha_hoy)+'.xlsx'))
        download_sheet(service, sheet_id_2, 'Dato_tienda_Dev_C&C.xlsx')
    except HttpError as error:
        print(f"An error occurred: {error}")

if __name__ == "__main__":
    main()
```
###  2. Procesar Archivos Excel

El script lee y procesa datos de los archivos Excel descargados, fusiona la información necesaria y crea reportes resumen:

```python
def read_excel(path: str, sheet_name: str) -> pd.DataFrame:
    buffer = StringIO()
    Xlsx2csv(path, outputencoding="utf-8", sheet_name=sheet_name).convert(buffer)
    buffer.seek(0)
    df = pd.read_csv(buffer)
    return df

# Procesamiento de datos
...
```
###  3. Enviar Reportes Resumen por Correo Electrónico

El script genera un reporte resumen para cada tienda, adjunta los archivos Excel correspondientes y envía un correo electrónico:

```python
# Configuración del correo electrónico
email_usuario = 'su_email@example.com'
contrasena = 'su_contraseña'
servidores_smtp = 'smtp.office365.com'
puertos_smtp = 587

# Iterar sobre las tiendas y enviar correos electrónicos
for tienda in PICKUP_TIENDA['Nombre_Tienda'].unique():
    ...
    # Crear el contenido del correo y adjuntar archivos
    ...
    try:
        # Enviar el correo electrónico
        server.sendmail(email_usuario, destinatarios_str_to.split(',') + destinatarios_str_CC.split(','), msg.as_string())
        registro_correos = pd.concat([registro_correos, pd.DataFrame({"Tienda": tienda, "Resultado": "Correo enviado correctamente"}, index=[0])])
        print(f"Correo electrónico enviado correctamente a {tienda}")
    except smtplib.SMTPDataError as e:
        registro_correos = pd.concat([registro_correos, pd.DataFrame({"Tienda": tienda, "Resultado": f"Error al enviar correo electrónico: {e}"}, index=[0])])
        print(f"Error al enviar correo electrónico a {tienda}: {e}")
    finally:
        server.quit()
```
### Conclusión

Este script automatiza la descarga, el procesamiento y la distribución de reportes diarios, asegurando que los interesados reciban datos actualizados de manera oportuna. Asegúrese de actualizar los marcadores de posición con datos y credenciales reales antes de ejecutar el script.
