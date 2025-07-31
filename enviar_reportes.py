import pandas as pd
import openpyxl
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import os

# -----------------------------------------------------------------------------
# CONFIGURACIÓN (MODIFICAR ESTOS VALORES)
# -----------------------------------------------------------------------------

# --- Configuración del Correo Electrónico ---
# IMPORTANTE: Para Gmail/Outlook, usa una "Contraseña de Aplicación", no tu contraseña principal.
EMAIL_REMITENTE = "tu_correo@ejemplo.com"  # Tu dirección de correo electrónico
PASSWORD_REMITENTE = "tu_password_de_aplicacion"    # Tu contraseña de aplicación generada
SMTP_SERVER = "smtp.gmail.com"  # Servidor de Gmail. Cambiar si usas otro (ej. "smtp.office365.com" para Outlook)
SMTP_PORT = 587                 # Puerto para TLS

# --- Configuración del Archivo Excel ---
NOMBRE_ARCHIVO_EXCEL = "datos_clientes.xlsx" # Nombre exacto de tu archivo Excel

# ¡¡CRÍTICO!! Define la celda en la hoja "E2" donde se coloca el número de apartamento
# para que la plantilla actualice sus datos (ej. 'B2', 'J7', etc.)
CELDA_APTO_EN_E2 = 'A1' # <--- CAMBIA ESTO POR LA CELDA CORRECTA

# --- Nombres de las hojas y columnas ---
HOJA_CORREOS = "CORREOS"
HOJA_PLANTILLA = "E2"
COLUMNA_EMAIL = "Correo"  # Nombre de la columna con los emails en la hoja "CORREOS"
COLUMNA_UNIDAD = "Unidad" # Nombre de la columna con las unidades en la hoja "CORREOS"

# -----------------------------------------------------------------------------
# INICIO DEL SCRIPT (NO MODIFICAR DE AQUÍ EN ADELANTE)
# -----------------------------------------------------------------------------

def enviar_correo(destinatario, asunto, cuerpo_html):
    """Establece conexión con el servidor SMTP y envía el correo."""
    try:
        # Configuración del mensaje
        msg = MIMEMultipart('alternative')
        msg['From'] = EMAIL_REMITENTE
        msg['To'] = destinatario
        msg['Subject'] = asunto
        msg.attach(MIMEText(cuerpo_html, 'html'))

        # Conexión y envío
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()  # Habilitar seguridad
        server.login(EMAIL_REMITENTE, PASSWORD_REMITENTE)
        texto = msg.as_string()
        server.sendmail(EMAIL_REMITENTE, destinatario, texto)
        server.quit()
        print(f"✅ Correo enviado exitosamente a {destinatario}")
        return True
    except smtplib.SMTPAuthenticationError:
        print(f"❌ Error de autenticación. Revisa tu email y contraseña de aplicación.")
        return False
    except Exception as e:
        print(f"❌ Error al enviar correo a {destinatario}: {e}")
        return False

def procesar_y_enviar():
    """Función principal que lee el Excel, procesa los datos y llama al envío."""
    # 1. Leer la lista de correos y unidades
    try:
        df_correos = pd.read_excel(NOMBRE_ARCHIVO_EXCEL, sheet_name=HOJA_CORREOS)
        print(f"Se encontraron {len(df_correos)} registros en la hoja '{HOJA_CORREOS}'.")
    except FileNotFoundError:
        print(f"🚨 Error: No se encontró el archivo '{NOMBRE_ARCHIVO_EXCEL}'. Asegúrate de que esté en la misma carpeta.")
        return
    except Exception as e:
        print(f"🚨 Error al leer la hoja '{HOJA_CORREOS}': {e}")
        return

    # 2. Cargar el libro de trabajo para modificarlo
    try:
        workbook = openpyxl.load_workbook(NOMBRE_ARCHIVO_EXCEL)
        hoja_plantilla = workbook[HOJA_PLANTILLA]
    except KeyError:
        print(f"🚨 Error: No se encontró la hoja '{HOJA_PLANTILLA}' en el archivo.")
        return
    except Exception as e:
        print(f"🚨 Error al cargar el libro de Excel: {e}")
        return

    # 3. Iterar sobre cada cliente para generar y enviar su reporte
    for index, fila in df_correos.iterrows():
        try:
            email_cliente = fila[COLUMNA_EMAIL]
            unidad_cliente = fila[COLUMNA_UNIDAD]

            print(f"\n Procesando Unidad: {unidad_cliente} | Destinatario: {email_cliente}")

            # 3.1. Actualizar el número de apartamento en la hoja de plantilla
            hoja_plantilla[CELDA_APTO_EN_E2] = unidad_cliente

            # 3.2. Guardar el libro en un archivo temporal para que las fórmulas se recalculen
            # Usamos un archivo temporal para no sobreescribir el original en cada paso.
            archivo_temporal = f"temp_report_{unidad_cliente}.xlsx"
            workbook.save(archivo_temporal)

            # 3.3. Leer SÓLO el rango deseado (B1:H42) del archivo temporal ya actualizado
            df_datos_apt = pd.read_excel(
                archivo_temporal,
                sheet_name=HOJA_PLANTILLA,
                header=None, # No hay encabezados en el rango que leemos
                usecols='B:H', # Columnas de la B a la H
                nrows=42,      # Leer las primeras 42 filas
                engine='openpyxl'
            )
            
            # 3.4. Eliminar el archivo temporal
            os.remove(archivo_temporal)

            # 3.5. Convertir los datos a una tabla HTML para el cuerpo del correo
            html_tabla = df_datos_apt.to_html(index=False, header=False, na_rep='', border=1)

            # 3.6. Componer el correo y enviarlo
            asunto = f"Información de su unidad: {unidad_cliente}"
            cuerpo_mensaje = f"""
            <html>
            <head>
                <style>
                    body {{ font-family: Arial, sans-serif; }}
                    table {{ border-collapse: collapse; width: 100%; }}
                    th, td {{ border: 1px solid #dddddd; text-align: left; padding: 8px; }}
                    tr:nth-child(even) {{ background-color: #f2f2f2; }}
                </style>
            </head>
            <body>
                <p>Estimado residente de la unidad <strong>{unidad_cliente}</strong>,</p>
                <p>A continuación, encontrará la información solicitada:</p>
                {html_tabla}
                <br>
                <p>Saludos cordiales,</p>
                <p><strong>La Administración</strong></p>
            </body>
            </html>
            """
            
            if not enviar_correo(email_cliente, asunto, cuerpo_mensaje):
                print("🚨 Deteniendo el script debido a un error de envío.")
                break # Detiene el bucle si falla el envío (ej. mala contraseña)

        except KeyError as e:
            print(f"🚨 Error: La columna {e} no existe en la hoja '{HOJA_CORREOS}'. Revisa la configuración.")
            break
        except Exception as e:
            print(f"🚨 Ocurrió un error inesperado procesando la unidad {unidad_cliente}: {e}")
            continue

    print("\n🎉 Proceso finalizado.")

if __name__ == "__main__":
    procesar_y_enviar()