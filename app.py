import streamlit as st
import pandas as pd
import datetime
import smtplib
from email.message import EmailMessage
from io import BytesIO
from openpyxl import Workbook
from openpyxl.workbook.protection import WorkbookProtection
# --- CONFIGURACIÓN SMTP DESDE SECRETS ---
smtp_user = st.secrets["SMTP_USER"]
smtp_pass = st.secrets["SMTP_PASS"]
# --- CONSTANTES ---
COMPRADOR = "612539"
OB_POR_PROVEEDOR = {
   "Efficold": "31005315",
   "Docriluc": "31005264"
}
# --- ARTÍCULOS ---
articulos = [
   {"Nº artículo": "1009250", "Descripción": "1009250 Mostradores Mahou 2025 (ROJO ESTRELLAS)", "limite": 1000, "multiplo": 10},
   {"Nº artículo": "1009248", "Descripción": "1009248 Mostradores ALH 2025 (VERDE)", "limite": 1000, "multiplo": 10},
   {"Nº artículo": "1003102", "Descripción": "1003102 Mostradores Mahou 2024 (ROJO)", "limite": 1000, "multiplo": 10},
   {"Nº artículo": "1001727", "Descripción": "1001727 Mostradores SM 2024 (VERDE)", "limite": 500, "multiplo": 10},
   {"Nº artículo": "1000511", "Descripción": "1000511 Mostradores ALH 2024 (GRIS)", "limite": 500, "multiplo": 10}
]
# --- PROVEEDORES ---
proveedor_opciones = {
   "Efficold": "10573",
   "Docriluc": "1083828"
}
# --- DESTINOS ---
destinos = {
   "CHAMANSER": "8751",
   "FEDUVIR": "8251",
   "CANFRIBUR": "8004",
   "ECESA GETAFE": "8214",
   "ECESA LEVANTE": "8255",
   "SILESTEC": "8005",
   "9VIPESET": "8071",
   "OTRO DESTINO": None
}
# --- INTERFAZ ---
st.title("PETICION DE MOSTRADORES Y BOTELLEROS")
# Selección de proveedor
proveedor_nombre = st.selectbox("Selecciona el proveedor:", list(proveedor_opciones.keys()))
proveedor_codigo = proveedor_opciones[proveedor_nombre]
ob_proveedor = OB_POR_PROVEEDOR[proveedor_nombre]
# Dirección de entrega
destino_seleccionado = st.selectbox("Selecciona el destino:", list(destinos.keys()))
if destino_seleccionado == "OTRO DESTINO":
   codigo_entrega = st.text_input("Introduce un código de entrega (4 cifras, empieza por 8):")
   if not (codigo_entrega and codigo_entrega.isdigit() and len(codigo_entrega) == 4 and codigo_entrega.startswith("8")):
       st.error("Código no válido")
       st.stop()
else:
   codigo_entrega = destinos[destino_seleccionado]
# Selección de cantidades
st.subheader("Selecciona las cantidades:")
pedido = []
for articulo in articulos:
   codigo = articulo["Nº artículo"]
   descripcion = articulo["Descripción"]
   limite = articulo["limite"]
   multiplo = articulo["multiplo"]
   cantidad = st.number_input(f"{descripcion}:", min_value=0, max_value=limite, step=multiplo, value=0)
   if cantidad > 0:
       pedido.append({
           "Fecha solicitud": datetime.date.today(),
           "OB": ob_proveedor,
           "Comprador": COMPRADOR,
           "LM aux": ob_proveedor,
           "Cód Prov": proveedor_codigo,
           "Proveedor": proveedor_nombre,
           "Suc/planta": 8040,
           "Dir entr": codigo_entrega,
           "Nº artículo": codigo,
           "Descripción": descripcion,
           "Autorizar cant": cantidad,
       })
# Crear Excel protegido
def crear_excel_protegido(df):
   wb = Workbook()
   ws = wb.active
   ws.append(df.columns.tolist())
   for _, row in df.iterrows():
       ws.append(row.tolist())
   wb.security = WorkbookProtection(workbookPassword="NESTARES_24", lockStructure=True)
   output = BytesIO()
   wb.save(output)
   output.seek(0)
   return output.read()
# Enviar correo con manejo de errores
def enviar_correo(destinatario, asunto, adjunto_bytes):
   try:
       msg = EmailMessage()
       msg["Subject"] = asunto
       msg["From"] = smtp_user
       msg["To"] = destinatario
       msg.set_content("Adjunto encontrarás el archivo de pedido.")
       msg.add_attachment(adjunto_bytes, maintype="application", subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename="pedido_materiales.xlsx")
       # Conectar al servidor SMTP de Brevo
       with smtplib.SMTP("smtp-relay.brevo.com", 587) as server:
           server.starttls()  # Iniciar la conexión segura
           server.login(smtp_user, smtp_pass)
           server.send_message(msg)
       return True
   except Exception as e:
       st.error(f"Error al enviar el correo: {e}")
       return False
# Acciones finales
if st.button("Generar y Enviar Pedido"):
   if not pedido:
       st.warning("No has seleccionado ninguna cantidad.")
       st.stop()
   df = pd.DataFrame(pedido)
   excel_bytes = crear_excel_protegido(df)
   if enviar_correo("davidvictoresg@hotmail.com", "Pedido de Materiales", excel_bytes):
       st.success("Correo enviado correctamente.")
       st.download_button("Descargar Pedido", data=excel_bytes, file_name="pedido_materiales.xlsx")
