import streamlit as st
import pandas as pd
import datetime
from io import BytesIO
from openpyxl import Workbook
from openpyxl.workbook.protection import WorkbookProtection
import base64
import sib_api_v3_sdk
from sib_api_v3_sdk.rest import ApiException
# --- CONFIGURACIÓN BREVO API ---
api_key = st.secrets["BREVO"]["API_KEY"]
# --- CONSTANTES ---
COMPRADOR = "612539"
OB_POR_PROVEEDOR = {
   "Efficold": "31005315",
   "Docriluc": "31005264"
}
articulos = [
   {"Nº artículo": "1009250", "Descripción": "1009250 Mostradores Mahou 2025 (ROJO ESTRELLAS)", "limite": 1000, "multiplo": 10},
   {"Nº artículo": "1009248", "Descripción": "1009248 Mostradores ALH 2025 (VERDE)", "limite": 1000, "multiplo": 10},
   {"Nº artículo": "1003102", "Descripción": "1003102 Mostradores Mahou 2024 (ROJO)", "limite": 1000, "multiplo": 10},
   {"Nº artículo": "1001727", "Descripción": "1001727 Mostradores SM 2024 (VERDE)", "limite": 500, "multiplo": 10},
   {"Nº artículo": "1000511", "Descripción": "1000511 Mostradores ALH 2024 (GRIS)", "limite": 500, "multiplo": 10}
]
proveedor_opciones = {
   "Efficold": "10573",
   "Docriluc": "1083828"
}
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
# Enviar correo con Brevo API
def enviar_correo(destinatario, asunto, adjunto_bytes):
   configuration = sib_api_v3_sdk.Configuration()
   configuration.api_key['api-key'] = api_key
   api_instance = sib_api_v3_sdk.TransactionalEmailsApi(sib_api_v3_sdk.ApiClient(configuration))
   adjunto_b64 = base64.b64encode(adjunto_bytes).decode()
   email = sib_api_v3_sdk.SendSmtpEmail(
       to=[{"email": destinatario}],
       sender={"name": "App Pedidos", "email": "pedidosmaterialmsm@gmail.com"},  # Email verificado en Brevo
       subject=asunto,
       html_content="<p>Adjunto encontrarás el archivo de pedido.</p>",
       attachment=[{
           "content": adjunto_b64,
           "name": "pedido_materiales.xlsx"
       }]
   )
   try:
       api_instance.send_transac_email(email)
       return True
   except ApiException as e:
       st.error(f"Error al enviar correo: {e}")
       return False
# Botón de acción final
if st.button("Generar y Enviar Pedido"):
   if not pedido:
       st.warning("No has seleccionado ninguna cantidad.")
       st.stop()
   df = pd.DataFrame(pedido)
   excel_bytes = crear_excel_protegido(df)
   if enviar_correo("dvictoresg@mahou-sanmiguel.com", "Pedido de Materiales", excel_bytes):
       st.success("Correo enviado correctamente.")
       st.download_button("Descargar Pedido", data=excel_bytes, file_name="pedido_materiales.xlsx")
