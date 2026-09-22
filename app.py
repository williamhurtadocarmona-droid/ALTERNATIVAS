import streamlit as st
import pandas as pd
import openpyxl
import io
import os

st.set_page_config(page_title="Generador GFPI-F-165 SENA", page_icon="📊", layout="centered")

st.title("📊 Generador de Formato GFPI-F-165 (Etapa Productiva)")
st.write("Sube el reporte de Sofia Plus para generar automáticamente el libro de Excel con las pestañas individuales de los aprendices activos (**EN FORMACIÓN**).")

# 1. Cargar el reporte de Sofia Plus
excel_reporte = st.file_uploader("Sube el reporte de Sofia Plus (.xlsx / .xls)", type=["xlsx", "xls"])

PLANTILLA_BASE = "GFPI-F-165.xlsx"  # Nombre del archivo base en tu GitHub

if excel_reporte is not None:
    try:
        if not os.path.exists(PLANTILLA_BASE):
            st.error(f"⚠️ No se encontró la plantilla base '{PLANTILLA_BASE}' en el servidor/GitHub. Por favor asegúrate de subir el archivo con ese nombre exacto.")
        else:
            # 2. Leer Encabezado del reporte de Sofia Plus (filas 1-12)
            df_encabezado = pd.read_excel(excel_reporte, nrows=12, header=None)
            
            regional = str(df_encabezado.iloc[0, 2]).strip() if len(df_encabezado) > 0 else ""
            centro = str(df_encabezado.iloc[1, 2]).strip() if len(df_encabezado) > 1 else ""
            ficha = str(df_encabezado.iloc[2, 2]).strip() if len(df_encabezado) > 2 else ""
            programa = str(df_encabezado.iloc[5, 2]).strip() if len(df_encabezado) > 5 else ""

            # 3. Leer la tabla de aprendices (desde la fila 13 en adelante)
            df_aprendices = pd.read_excel(excel_reporte, skiprows=12)
            df_aprendices.columns = [str(c).strip() for c in df_aprendices.columns]

            # Identificar columnas estándar de Sofia Plus
            col_estado = [c for c in df_aprendices.columns if 'Estado' in c or 'ESTADO' in c]
            col_estado = col_estado[0] if col_estado else "Estado"

            # 4. FILTRAR SOLO APRENDICES EN FORMACION
            df_filtrado = df_aprendices[df_aprendices[col_estado].astype(str).str.upper().str.contains("EN FORMAC", na=False)]

            cant_total = len(df_aprendices)
            cant_activos = len(df_filtrado)

            st.info(f"📋 **Resumen del Reporte:** Total registros: {cant_total} | Aprendices **EN FORMACIÓN**: {cant_activos}")

            if cant_activos == 0:
                st.warning("⚠️ No se encontraron aprendices con estado 'EN FORMACIÓN' en el archivo subido.")
            else:
                st.dataframe(df_filtrado[['Tipo de Documento', 'Número de Documento', 'Nombres', 'Apellidos', col_estado]].head(10))

                if st.button("🚀 Generar Excel con Pestañas Individuales"):
                    with st.spinner("Generando pestañas individuales por cada aprendiz..."):
                        
                        # Cargar plantilla Excel con openpyxl
                        wb = openpyxl.load_workbook(PLANTILLA_BASE)
                        
                        # Ubicar la pestaña plantilla individual (ej. 'Instrucciones F2. Individual' o 'F2. Individual')
                        hoja_plantilla_nombre = None
                        for name in wb.sheetnames:
                            if "F2" in name or "Individual" in name:
                                hoja_plantilla_nombre = name
                                break
                        
                        if not hoja_plantilla_nombre:
                            hoja_plantilla_nombre = wb.sheetnames[-1]  # Usar la última si no encuentra coincidencia

                        hoja_base = wb[hoja_plantilla_nombre]

                        col_tipo_doc = [c for c in df_filtrado.columns if 'Tipo' in c][0]
                        col_num_doc = [c for c in df_filtrado.columns if 'Número' in c or 'Documento' in c][0]
                        col_nombres = [c for c in df_filtrado.columns if 'Nombre' in c][0]
                        col_apellidos = [c for c in df_filtrado.columns if 'Apellido' in c][0]
                        col_correo = [c for c in df_filtrado.columns if 'Correo' in c or 'Email' in c]
                        col_correo = col_correo[0] if col_correo else None
                        col_tel = [c for c in df_filtrado.columns if 'Teléfono' in c or 'Celular' in c]
                        col_tel = col_tel[0] if col_tel else None

                        # Recorrer cada aprendiz en formación y duplicar la pestaña
                        for idx, row in df_filtrado.iterrows():
                            nombre = str(row[col_nombres]).strip()
                            apellido = str(row[col_apellidos]).strip()
                            
                            # Nombre de pestaña válido (máximo 30 caracteres)
                            titulo_pestaña = f"{nombre.split()[0]} {apellido.split()[0]}"[:30]

                            # Crear copia de la pestaña plantilla
                            target_sheet = wb.copy_worksheet(hoja_base)
                            target_sheet.title = titulo_pestaña

                            # Rellenar los campos correspondientes
                            target_sheet['D11'] = programa          # Nivel y Programa de formación
                            target_sheet['Q11'] = ficha             # Número de Grupo - ficha
                            target_sheet['D12'] = centro            # Centro de formación
                            
                            target_sheet['C16'] = nombre            # Nombres
                            target_sheet['G16'] = apellido          # Apellidos
                            target_sheet['K16'] = str(row[col_tipo_doc]) # Tipo Doc
                            target_sheet['M16'] = str(row[col_num_doc])  # Número Doc
                            
                            if col_correo:
                                target_sheet['Q16'] = str(row[col_correo]) # Correo
                            if col_tel:
                                target_sheet['U16'] = str(row[col_tel])    # Teléfono

                        # Guardar resultado en memoria
                        output = io.BytesIO()
                        wb.save(output)
                        output.seek(0)

                        st.download_button(
                            label="📥 Descargar Libro Excel Completo (.xlsx)",
                            data=output,
                            file_name=f"GFPI-F-165_Ficha_{ficha}_Consolidado.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                        st.success(f"¡Éxito! Se generaron {cant_activos} pestañas individuales correctamente.")

    except Exception as e:
        st.error(f"Ocurrió un error al procesar el archivo: {e}")