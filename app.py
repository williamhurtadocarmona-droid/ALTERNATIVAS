import streamlit as st
import pandas as pd
import openpyxl
import io
import os

st.set_page_config(page_title="Generador GFPI-F-165 SENA", page_icon="📊", layout="centered")

st.title("📊 Generador de Formato GFPI-F-165 (Etapa Productiva)")
st.write("Sube el reporte de Sofia Plus para generar automáticamente las pestañas individuales de los aprendices activos (**EN FORMACIÓN**).")

# 1. Cargar el reporte de Sofia Plus
excel_reporte = st.file_uploader("Sube el reporte de Sofia Plus (.xlsx / .xls)", type=["xlsx", "xls"])

PLANTILLA_BASE = "GFPI-F-165.xlsx"  # Archivo base en GitHub

if excel_reporte is not None:
    try:
        if not os.path.exists(PLANTILLA_BASE):
            st.error(f"⚠️ No se encontró la plantilla base '{PLANTILLA_BASE}' en GitHub. Asegúrate de que el archivo tenga ese nombre exacto.")
        else:
            # 2. Leer Encabezado (filas 1-12)
            df_encabezado = pd.read_excel(excel_reporte, nrows=12, header=None)
            
            regional = str(df_encabezado.iloc[0, 2]).strip() if len(df_encabezado) > 0 and pd.notna(df_encabezado.iloc[0, 2]) else ""
            centro = str(df_encabezado.iloc[1, 2]).strip() if len(df_encabezado) > 1 and pd.notna(df_encabezado.iloc[1, 2]) else ""
            ficha = str(df_encabezado.iloc[2, 2]).strip() if len(df_encabezado) > 2 and pd.notna(df_encabezado.iloc[2, 2]) else ""
            programa = str(df_encabezado.iloc[5, 2]).strip() if len(df_encabezado) > 5 and pd.notna(df_encabezado.iloc[5, 2]) else ""

            # 3. Leer la tabla de aprendices (desde la fila 13)
            df_aprendices = pd.read_excel(excel_reporte, skiprows=12)
            df_aprendices.columns = [str(c).strip() for c in df_aprendices.columns]

            # Detectar columna de Estado
            col_estado = [c for c in df_aprendices.columns if 'Estado' in c or 'ESTADO' in c]
            col_estado = col_estado[0] if col_estado else df_aprendices.columns[-1]

            # Detectar Número de Documento (excluyendo 'Tipo')
            col_num_doc = [c for c in df_aprendices.columns if ('Número' in c or 'Numero' in c or 'Documento' in c or 'Identificación' in c) and 'Tipo' not in c]
            col_num_doc = col_num_doc[0] if col_num_doc else None

            # Detectar Nombres y Apellidos
            col_nombres = [c for c in df_aprendices.columns if 'Nombre' in c or 'Aprendiz' in c]
            col_apellidos = [c for c in df_aprendices.columns if 'Apellido' in c]
            
            col_tipo_doc = [c for c in df_aprendices.columns if 'Tipo' in c]
            col_tipo_doc = col_tipo_doc[0] if col_tipo_doc else None

            col_correo = [c for c in df_aprendices.columns if 'Correo' in c or 'Email' in c]
            col_correo = col_correo[0] if col_correo else None

            col_tel = [c for c in df_aprendices.columns if 'Teléfono' in c or 'Celular' in c or 'Móvil' in c]
            col_tel = col_tel[0] if col_tel else None

            # 4. FILTRAR POR "EN FORMACIÓN" Y ELIMINAR DUPLICADOS
            df_activos = df_aprendices[df_aprendices[col_estado].astype(str).str.upper().str.contains("EN FORMAC", na=False)].copy()

            col_clave = col_num_doc if col_num_doc else (col_nombres[0] if col_nombres else None)

            if col_clave:
                df_unicos = df_activos.drop_duplicates(subset=[col_clave]).copy()
            else:
                df_unicos = df_activos.copy()

            cant_totales_filas = len(df_aprendices)
            cant_unicos_activos = len(df_unicos)

            st.info(f"📋 **Resumen:** Filas en reporte: {cant_totales_filas} | Aprendices ÚNICOS en **EN FORMACIÓN**: **{cant_unicos_activos}**")

            if cant_unicos_activos == 0:
                st.warning("⚠️ No se encontraron aprendices con estado 'EN FORMACIÓN' en el archivo.")
            else:
                # Mostrar lista completa en pantalla
                cols_mostrar = [c for c in [col_tipo_doc, col_num_doc, col_nombres[0] if col_nombres else None, col_estado] if c is not None]
                st.dataframe(df_unicos[cols_mostrar])

                if st.button("🚀 Generar Excel con Pestañas Individuales"):
                    with st.spinner(f"Generando {cant_unicos_activos} pestañas a partir de 'Selección Modificación F2 Indiv'..."):
                        
                        wb = openpyxl.load_workbook(PLANTILLA_BASE)
                        
                        # Nombre exacto de la pestaña plantilla a duplicar
                        NOMBRE_HOJA_PLANTILLA = "Selección Modificación F2 Indiv"

                        if NOMBRE_HOJA_PLANTILLA in wb.sheetnames:
                            hoja_base = wb[NOMBRE_HOJA_PLANTILLA]
                        else:
                            # Búsqueda alternativa por coincidencia parcial si el nombre varía levemente
                            hoja_encontrada = [s for s in wb.sheetnames if "F2" in s or "Indiv" in s]
                            if hoja_encontrada:
                                hoja_base = wb[hoja_encontrada[0]]
                            else:
                                st.error(f"⚠️ No se encontró la pestaña '{NOMBRE_HOJA_PLANTILLA}' en el archivo de plantilla Excel.")
                                st.stop()

                        # Recorrer a todos los aprendices únicos activos
                        for idx, row in df_unicos.iterrows():
                            # Obtener Nombres y Apellidos
                            if col_apellidos:
                                nombre = str(row[col_nombres[0]]).strip()
                                apellido = str(row[col_apellidos[0]]).strip()
                            else:
                                nombre_completo = str(row[col_nombres[0]]).strip().split()
                                nombre = " ".join(nombre_completo[:2]) if len(nombre_completo) > 1 else nombre_completo[0]
                                apellido = " ".join(nombre_completo[2:]) if len(nombre_completo) > 2 else ""

                            tipo_doc = str(row[col_tipo_doc]) if col_tipo_doc and pd.notna(row[col_tipo_doc]) else ""
                            num_doc = str(row[col_num_doc]).strip() if col_num_doc and pd.notna(row[col_num_doc]) else ""

                            # Definir nombre corto para la pestaña (máx. 31 caracteres)
                            id_pestana = num_doc[-4:] if num_doc else nombre.split()[0]
                            titulo_pestaña = f"{nombre.split()[0]} {id_pestana}"[:30]

                            # Copiar la pestaña 'Selección Modificación F2 Indiv'
                            target_sheet = wb.copy_worksheet(hoja_base)
                            target_sheet.title = titulo_pestaña

                            # Rellenar campos en la plantilla copiada
                            target_sheet['D11'] = programa          # Programa
                            target_sheet['Q11'] = ficha             # Ficha
                            target_sheet['D12'] = centro            # Centro
                            
                            target_sheet['C16'] = nombre            # Nombres
                            target_sheet['G16'] = apellido          # Apellidos
                            target_sheet['K16'] = tipo_doc          # Tipo Doc
                            target_sheet['M16'] = num_doc           # Número Doc
                            
                            if col_correo and pd.notna(row[col_correo]):
                                target_sheet['Q16'] = str(row[col_correo])
                            if col_tel and pd.notna(row[col_tel]):
                                target_sheet['U16'] = str(row[col_tel])

                        # Guardar el libro generado
                        output = io.BytesIO()
                        wb.save(output)
                        output.seek(0)

                        st.download_button(
                            label="📥 Descargar Libro Excel Consolidado (.xlsx)",
                            data=output,
                            file_name=f"GFPI-F-165_Ficha_{ficha}_Consolidado.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                        st.success(f"¡Listo! Se copiaron y rellenaron {cant_unicos_activos} pestañas individuales basándose en 'Selección Modificación F2 Indiv'.")

    except Exception as e:
        st.error(f"Ocurrió un error al procesar el archivo: {e}")
