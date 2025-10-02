import streamlit as st
from utils import process_files_and_generate_report, generate_ratios_charts_pdf 
import tempfile, os, shutil

st.set_page_config(page_title="Analizador Financiero SMV", layout="wide")
st.title("📊 Analizador Financiero (Análisis Vertical, Horizontal y Ratios)")
st.markdown("Estimado usuario: Carga los archivos de los estados financieros por año (uno por año) para generar el reporte.")

# Inicializar Session State
if 'report_path' not in st.session_state: 
    st.session_state.update({
        'report_path': None, 'pdf_path': None, 'temp_dir': None,
        'excel_data': None, 'pdf_data': None,
        'excel_filename': "REPORTE_ANALISIS_FINANCIERO.xlsx",
        'pdf_filename': "GRAFICOS_RATIOS_FINANCIEROS.pdf"
    })

def clear_session_files():
    st.session_state.update({'report_path': None, 'pdf_path': None, 'excel_data': None, 'pdf_data': None})
    if st.session_state['temp_dir'] and os.path.exists(st.session_state['temp_dir']):
        shutil.rmtree(st.session_state['temp_dir'])
        st.session_state['temp_dir'] = None

uploaded = st.file_uploader("📂 Sube los archivos Excel (múltiples años)", accept_multiple_files=True, type=['xlsx','xlsm','xls'])
st.info("El libro modelo (BASE.xlsx) ya está incluido. Los años deben ser consecutivos (Ej: 2023, 2022, 2021).")

if uploaded:
    st.write(f"Cargados: **{len(uploaded)} archivo(s)**.")
    
    if st.button("🚀 Generar Reporte Completo (Excel)"):
        clear_session_files() 
        with st.spinner("Procesando archivos y generando reporte Excel..."):
            tmpdir = tempfile.mkdtemp()
            paths = []
            for f in uploaded:
                p = os.path.join(tmpdir, f.name)
                with open(p, "wb") as out: out.write(f.getbuffer())
                paths.append(p)
            try:
                out_path = process_files_and_generate_report(paths, model_path="BASE.xlsx", output_dir=tmpdir)
                st.session_state.update({'report_path': out_path, 'temp_dir': tmpdir})
                with open(out_path, "rb") as fh: st.session_state['excel_data'] = fh.read()
                st.success("✅ Reporte Excel generado exitosamente. Puedes descargarlo a continuación.")
            except Exception as e:
                st.error(f"❌ Error al generar el reporte: {str(e)}")
                clear_session_files()
    
    if st.session_state['excel_data']:
        st.markdown("---")
        st.subheader("Descargar Resultados")
        col_excel, col_pdf_ctrl = st.columns([1, 1])

        with col_excel:
            st.download_button(
                label=f"⬇️ Descargar Reporte Excel ({st.session_state['excel_filename']})",
                data=st.session_state['excel_data'],
                file_name=st.session_state['excel_filename'],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                help="Reporte con Análisis Vertical, Horizontal y Ratios calculado."
            )
            
        with col_pdf_ctrl:
            if st.session_state['pdf_data']:
                st.download_button(
                    label=f"⬇️ Descargar Gráficos PDF ({st.session_state['pdf_filename']})", 
                    data=st.session_state['pdf_data'], 
                    file_name=st.session_state['pdf_filename'],
                    mime="application/pdf", key='download_pdf_btn'
                )
            elif st.button("📈 Generar Gráficas de Ratios (PDF)", key='gen_pdf_btn'):
                with st.spinner("Generando visualizaciones en PDF con nuevo diseño..."):
                    try:
                        pdf_path = generate_ratios_charts_pdf(st.session_state['report_path'], st.session_state['temp_dir'])
                        st.session_state['pdf_path'] = pdf_path
                        with open(pdf_path, "rb") as fh_pdf: st.session_state['pdf_data'] = fh_pdf.read()
                        st.success("✅ Gráficas generadas. Descarga el PDF ahora.")
                        st.rerun()
                    except Exception as e:
                        st.error(f" Error al generar el PDF de gráficas: {str(e)}")
else:
    clear_session_files()
    st.write("Aún no has subido archivos. Por favor, selecciona los archivos Excel para comenzar el análisis.")