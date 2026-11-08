import streamlit as st
import google.generativeai as genai
import json
import os
import tempfile
from io import BytesIO
import openpyxl
import pandas as pd
from datetime import datetime
import traceback
from PIL import Image

st.set_page_config(page_title="Gestor de Gastos (Reporte Mensual)", page_icon="💸", layout="wide")

st.title("💸 Tu Gestor de Gastos (Reporte Mensual)")
st.markdown("Procesamiento rápido de comprobantes y actualización automática de tu Excel.")

with st.sidebar:
    st.header("⚙️ Configuración")
    mes_actual = st.selectbox("Mes de carga", ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"], index=datetime.now().month - 1)
    anio_actual = st.number_input("Año de carga", min_value=2020, max_value=2100, value=datetime.now().year)

st.markdown("### 1. Sube tu Reporte de Cuentas actual")
master_file = st.file_uploader("Sube tu archivo Excel (ej. Reporte de tus cuentas - septiembre de 2026.xlsx)", type=["xlsx"])

st.markdown("### 2. Sube tus resúmenes o capturas")
uploaded_files = st.file_uploader("Imágenes o PDFs de Wise, MP, ARQ", type=["png", "jpg", "jpeg", "pdf"], accept_multiple_files=True)

if st.button("Procesar y Actualizar Excel", type="primary"):
    try:
        api_key = st.secrets["API_KEY"]
    except Exception:
        api_key = ""
    
    if not api_key:
        st.error("No se encontró la 'API_KEY' en los Secrets de Streamlit. Por favor, configúrala en el panel de tu app en la nube.")
    elif not master_file:
        st.error("Por favor, sube tu reporte Excel actual.")
    elif not uploaded_files:
        st.warning("Sube al menos un comprobante de gastos.")
    else:
        try:
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel('gemini-1.5-flash')
            all_expenses = []
            
            with st.spinner("Procesando y comprimiendo comprobantes..."):
                for file in uploaded_files:
                    try:
                        file_bytes = file.getvalue()
                        ext = file.name.split('.')[-1].lower()
                        mime_type = "image/jpeg"
                        
                        # Convertimos RGBA a RGB para evitar el error de transparencia en capturas de Mac
                        if ext in ["png", "jpg", "jpeg"]:
                            img = Image.open(BytesIO(file_bytes))
                            if img.mode in ("RGBA", "P"):
                                img = img.convert("RGB")
                            img.thumbnail((1600, 1600))
                            buffered = BytesIO()
                            img.save(buffered, format="JPEG", quality=85)
                            file_bytes = buffered.getvalue()
                            mime_type = "image/jpeg"
                        elif ext == "pdf":
                            mime_type = "application/pdf"
                        
                        prompt = f'''
                        Analiza este comprobante o captura. Extrae los movimientos de SALIDA de dinero.
                        Devuelve ÚNICAMENTE un arreglo JSON con esta estructura exacta para cada gasto:
                        [{{
                            "Tipo": "Gasto",
                            "Tipo de Gasto": "variable",
                            "Dia": 15,
                            "Mes": "{mes_actual}",
                            "Año": {anio_actual},
                            "Detalle": "Comercio o motivo limpio (ej. La gallega, Starbucks, Uber, etc)",
                            "Categoria": "Supermercado, Extras, Delivery, Salidas, Transporte, Salud, etc",
                            "Presupuesto": "Personal",
                            "Monto": 1500.5
                        }}]
                        Reglas:
                        - 'Dia' debe ser solo el número del día (ej. 5). Si no puedes deducirlo exacto, usa 1.
                        - 'Monto' numérico positivo, sin comillas, sin signos.
                        - Ignora ingresos.
                        - Devuelve SOLO el JSON válido, sin markdown adicional.
                        '''
                        
                        response = model.generate_content([
                            {
                                'mime_type': mime_type,
                                'data': file_bytes
                            },
                            prompt
                        ])
                        
                        json_str = response.text.strip()
                        if json_str.startswith("```json"): json_str = json_str[7:-3]
                        elif json_str.startswith("```"): json_str = json_str[3:-3]
                            
                        parsed_data = json.loads(json_str)
                        if isinstance(parsed_data, list):
                            all_expenses.extend(parsed_data)
                        elif isinstance(parsed_data, dict):
                            all_expenses.append(parsed_data)
                            
                    except Exception as ex:
                        st.error(f"Error detallado en {file.name}: {str(ex)}")
                        st.code(traceback.format_exc())
            
            if all_expenses:
                st.success("¡Extracción veloz completada! Inyectando en tu Excel...")
                
                wb = openpyxl.load_workbook(master_file)
                
                # Búsqueda flexible de la pestaña de ingresos/egresos
                target_sheet = None
                for sheet in wb.sheetnames:
                    if "ingreso" in sheet.lower() or "egreso" in sheet.lower() or "gasto" in sheet.lower():
                        target_sheet = sheet
                        break
                
                if not target_sheet:
                    target_sheet = wb.sheetnames[0] # Usa la primera por defecto si no hay coincidencias
                
                ws = wb[target_sheet]
                
                for exp in all_expenses:
                    ws.append([
                        exp.get('Tipo', 'Gasto'),
                        exp.get('Tipo de Gasto', 'variable'),
                        exp.get('Dia', 1),
                        exp.get('Mes', mes_actual),
                        exp.get('Año', anio_actual),
                        exp.get('Detalle', ''),
                        exp.get('Categoria', 'Extras'),
                        exp.get('Presupuesto', 'Personal'),
                        exp.get('Monto', 0)
                    ])
                
                output = BytesIO()
                wb.save(output)
                output.seek(0)
                
                st.balloons()
                st.download_button(
                    label="📥 Descargar tu Reporte Actualizado",
                    data=output,
                    file_name=f"Reporte_Actualizado_{mes_actual}_{anio_actual}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
                
                st.subheader(f"Gastos agregados a la hoja '{target_sheet}':")
                st.dataframe(pd.DataFrame(all_expenses))
                
        except Exception as general_err:
            st.error(f"Error general en la ejecución: {str(general_err)}")
            st.code(traceback.format_exc())
