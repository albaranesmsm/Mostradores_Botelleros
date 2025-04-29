import streamlit as st
import pandas as pd
import datetime
import smtplib
from email.message import EmailMessage
from io import BytesIO
from openpyxl import Workbook
from openpyxl.workbook.protection import WorkbookProtection
# --- DATOS FIJOS ---
COMPRADOR = "612539"
# --- RELACIÓN ARTÍCULOS Y LM AUX ---
lm_aux_por_articulo = {
   "1009250": "24341155",
   "1009248": "24341161",
   "1003102": "24341155",
   "1001727": "24341159",
   "1000511": "24341161"
}
# --- LISTADO DE ARTÍCULOS ---
articulos = [
   {"Nº artículo": "1009250", "Descripción": "1009250 Mostradores Mahou 2025 (ROJO ESTRELLAS)"},
   {"Nº artículo": "1009248", "Descripción": "1009248 Mostradores ALH 2025 (VERDE)"},
   {"Nº artículo": "1003102", "Descripción": "1003102 Mostradores Mahou 2024 (ROJO)"},
   {"Nº artículo": "1001727", "Descripción": "1001727 Mostradores SM 2024 (VERDE)"},
   {"Nº artículo": "1000511", "Descripción": "1000511 Mostradores ALH 2024 (GRIS)"}
]
# --- RESTRICCIONES POR ARTÍCULO ---
restricciones = {
   "1600043": {"multiplo": 10, "max": 1000},
   "1600050": {"multiplo": 10, "max": 1000},
   "1600051": {"multiplo": 10, "max": 1000},
   "1600052": {"multiplo": 10, "max": 1000},
   "1600053": {"multiplo": 10, "max": 1000}
}
# --- RELACIÓN DIRECCIONES DE ENTREGA ---
destinos = {
   "CHAMANSER": "8751",
   "FEDUVIR": "8251",
   "CANFRIBUR": "8004",
   "ECESA GETAFE": "8214",
   "ECESA LEVANTE": "8255",
   "SILESTEC": "8005",
   "9VIPESET": "8071",
   "OTRO DESTINO": ""
}
# --- RELACIÓN PROVEEDORES ---
proveedores_disponibles = {
   "Efficold": "10573",
   "Docriluc": "1083828"
}
# --- INTERFAZ ---
st.title("PETICION DE MOSTRADORES Y BOTELLEROS")
# --- Dirección de entrega ---
destino_seleccionado = st.selectbox("Selecciona Dirección de Entrega:", list(destinos.keys()))
if destino_seleccionado == "OTRO DESTINO":
   dir_entrega = st.text_input("Introduce manualmente un código de Dirección de Entrega (4 cifras empezando por 8):", max_chars=4)
   if not dir_entrega or not (dir_entrega.isdigit() and len(dir_entrega) == 4 and dir_entrega.startswith("8")):
       st.error("Debe introducir un código de 4 cifras que empiece por 8")
       st.stop()
else:
   dir_entrega = destinos[destino_seleccionado]
# --- Selección de proveedor general para el pedido ---
proveedor_seleccionado = st.selectbox("Selecciona el Proveedor:", list(proveedores_disponibles.keys()))
codigo_proveedor = proveedores_disponibles[proveedor_seleccionado]
st.subheader("Selecciona las cantidades:")
pedido = []
for articulo in articulos:
   codigo = str(articulo["Nº artículo"])
   descripcion = articulo["Descripción"]
   maximo = restricciones.get(codigo, {}).get("max", 1000)
   multiplo = restricciones.get(codigo, {}).get("multiplo", 1)
   cantidad = st.number_input(
       f"{descripcion} (Múltiplo: {multiplo}, Máx: {maximo})",
       min_value=0, max_value=maximo, step=multiplo, value=0,
   )
   if cantidad > 0:
       pedido.append({
           "Fecha solicitud": datetime.date.today(),
           "OB": lm_aux_por_articulo[codigo],
           "Comprador": COMPRADOR,
           "LM aux": lm_aux_por_articulo[codigo],
           "Cód Prov": codigo_proveedor,
           "Proveedor": proveedor_seleccionado,
           "Suc/planta": 8040,
           "Dir entr": dir_entrega,
           "Nº artículo": codigo,
           "Descripción": descripcion,
           "Autorizar cant": cantidad,
       })
# --- Excel protegido ---
def crear_excel_protegido(df):
   wb = Workbook()
   ws = wb.active
   ws.append(df.columns.tolist())
   for _, row in df.iterrows():
       ws.append(row.tolist())
   wb.security = WorkbookProtection(workbookPassword="NESTARES_24", lockStructure=True)
   # Guardar en BytesIO
   excel_buffer = BytesIO()
   wb.save(excel_buffer)
   excel_buffer.seek(0)
   return excel_buffer
# --- Envío de correo ---
def enviar_correo(destinatario, asunto, adjunto_bytes):
   msg = EmailMessage()
   msg["Subject"] = asunto
   msg["From"] = "pedidosmaterialesmsm@gmail.com"
   msg["To"] = destinatario
   msg.set_content("Adjunto encontrarás el archivo de pedido.")
   msg.add_attachment(adjunto_bytes, maintype='application', subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet', filename="pedido_materiales.xlsx")
   with smtplib.SMTP("smtp.gmail.com", 587) as server:
       server.starttls()
       server.login("pedidosmaterialesmsm@gmail.com", "iquiuslwxribewal")
       server.send_message(msg)
# --- Botón de generación ---
if st.button("Generar Pedido"):
   if pedido:
       df = pd.DataFrame(pedido)
       excel_bytes = crear_excel_protegido(df)
       st.success("Pedido generado correctamente.")
       st.download_button("Descargar Pedido", data=excel_bytes, file_name="pedido_materiales.xlsx")
       if st.button("Enviar Pedido por Email"):
           enviar_correo("dvictoresg@mahou-sanmiguel.com", "Pedido de Materiales", excel_bytes.getvalue())
           st.success("Correo enviado correctamente.")
   else:
       st.warning("No se ha seleccionado ningún artículo.")