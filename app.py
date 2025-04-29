import streamlit as st
import pandas as pd
import datetime
import smtplib
from email.message import EmailMessage
from io import BytesIO
from openpyxl import Workbook
from openpyxl.workbook.protection import WorkbookProtection
import os
# --- CONFIGURACIÓN SMTP ---
smtp_user = os.getenv("SMTP_USER")
smtp_pass = os.getenv("SMTP_PASS")
# --- DATOS FIJOS ---
COMPRADOR = "612539"
OB_POR_ARTICULO = {
   "1009250": "24341155",
   "1009248": "24341161",
   "1003102": "24341155",
   "1001727": "24341159",
   "1000511": "24341161"
}
# --- ARTÍCULOS ---
articulos = [
   {"Nº artículo": "1009250", "Descripción": "1009250 Mostradores Mahou 2025 (ROJO ESTRELLAS)"},
   {"Nº artículo": "1009248", "Descripción": "1009248 Mostradores ALH 2025 (VERDE)"},
   {"Nº artículo": "1003102", "Descripción": "1003102 Mostradores Mahou 2024 (ROJO)"},
   {"Nº artículo": "1001727", "Descripción": "1001727 Mostradores SM 2024 (VERDE)"},
   {"Nº artículo": "1000511", "Descripción": "1000511 Mostradores ALH 2024 (GRIS)"}
]
# --- PROVEEDORES ---
proveedor_opciones = {
   "Efficold": "10573",
   "Docriluc": "1083828"
}
proveedor_nombre = st.selectbox("Selecciona el proveedor:", list(proveedor_opciones.keys()))
proveedor_codigo = proveedor_opciones[proveedor_nombre]
# --- DIRECCIONES DE ENTREGA ---
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
st.title("PETICION DE MOSTRADORES Y BOTELLEROS")
destino_seleccionado = st.selectbox("Selecciona el destino:", list(destinos.keys()))
if destino_seleccionado == "OTRO DESTINO":
   codigo_entrega = st.text_input("Introduce un código de entrega (4 cifras, empieza por 8):")
   if not (codigo_entrega and codigo_entrega.isdigit() and len(codigo_entrega) == 4 and codigo_entrega.startswith("8")):
       st.error("Código no válido")
       st.stop()
else:
   codigo_entrega = destinos[destino_seleccionado]
# --- PEDIDO ---
st.subheader("Selecciona las cantidades:")
pedido = []
for articulo in articulos:
   codigo = articulo["Nº artículo"]
   descripcion = articulo["Descripción"]
   ob = OB_POR_ARTICULO.get(codigo)
   cantidad = st.number_input(f"{descripcion}:", min_value=0, step=1, value=0)
   if cantidad > 0:
       pedido.append({
           "Fecha solicitud": datetime.date.today(),
           "OB": ob,
           "Comprador": COMPRADOR,
           "LM aux": ob,
           "Cód Prov": proveedor_codigo,
           "Proveedor": proveedor_nombre,
           "Suc/planta": 8040,
           "Dir entr": codigo_entrega,
           "Nº artículo": codigo,
           "Descripción": descripcion,
           "Autorizar cant": cantidad,
       })
# --- EXCEL ---
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
# --- CORREO ---
def enviar_correo(destinatario, asunto, adjunto_bytes):
   msg = EmailMessage()
   msg["Subject"] = asunto
   msg["From"] = smtp_user
   msg["To"] = destinatario
   msg.set_content("Adjunto encontrarás el archivo de pedido.")
   msg.add_attachment(adjunto_bytes, maintype="application", subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename="pedido_materiales.xlsx")
   with smtplib.SMTP("relay.brevo.com", 587) as server:
       server.starttls()
       server.login(smtp_user, smtp_pass)
       server.send_message(msg)
# --- ACCIONES ---
if st.button("Generar y Enviar Pedido"):
   if not pedido:
       st.warning("No has seleccionado ninguna cantidad.")
       st.stop()
   df = pd.DataFrame(pedido)
   excel_bytes = crear_excel_protegido(df)
   enviar_correo("dvictoresg@mahou-sanmiguel.com", "Pedido de Materiales", excel_bytes)
   st.success("Correo enviado correctamente.")
   st.download_button("Descargar Pedido", data=excel_bytes, file_name="pedido_materiales.xlsx")
