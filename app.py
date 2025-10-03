import streamlit as st
from utils import process_files_and_generate_report, generate_ratios_charts_pdf 
import tempfile, os, shutil

st.set_page_config(page_title="Analizador Financiero SMV", layout="wide")
st.title("📊 Analizador Financiero")
st.markdown("Carga los archivos de los estados financieros por año (consecutivos) para generar el reporte.")

# --- Configuración y Limpieza del Estado ---
# Consolidar estado: 'dir' (temp dir), 'xlsx' (datos Excel), 'pdf' (datos PDF)
if 'state' not in st.session_state: 
    st.session_state['state'] = {'dir': None, 'xlsx': None, 'pdf': None}

def reset_state():
    """Limpia el directorio temporal y resetea la sesión."""
    if st.session_state['state']['dir']:
        shutil.rmtree(st.session_state['state']['dir'], ignore_errors=True)
    st.session_state['state'] = {'dir': None, 'xlsx': None, 'pdf': None}

uploaded = st.file_uploader("📂 Sube archivos Excel (múltiples años)", accept_multiple_files=True, type=['xlsx','xlsm','xls'])
st.info("El libro modelo (BASE.xlsx) ya está incluido. Los años deben ser consecutivos.")

if uploaded:
    st.write(f"Cargados: **{len(uploaded)} archivo(s)**.")
    
    # --- Generación de Excel ---
    if st.button("🚀 Generar Reporte Excel"):
        reset_state()
        try:
            tmpdir = tempfile.mkdtemp()
            paths = []
            for f in uploaded: # Guardar archivos subidos al disco temporal
                p = os.path.join(tmpdir, f.name)
                with open(p, "wb") as out: out.write(f.getbuffer())
                paths.append(p)

            out_path = process_files_and_generate_report(paths, output_dir=tmpdir)
            
            with open(out_path, "rb") as fh: # Leer archivo generado al estado
                st.session_state.state['xlsx'] = {'data': fh.read(), 'name': os.path.basename(out_path)}
            st.session_state.state['dir'] = tmpdir
            st.success("✅ Reporte Excel generado. Descarga disponible.")
        except Exception as e:
            st.error(f"❌ Error al generar el reporte: {str(e)}")
            reset_state()

    # --- Descarga y Generación de PDF ---
    if (xlsx_data := st.session_state.state['xlsx']):
        st.markdown("---")
        pdf_data = st.session_state.state['pdf']
        col_excel, col_pdf_ctrl = st.columns([1, 1])

        with col_excel:
            # Botón de descarga de Excel (usa el nombre dinámico del reporte)
            st.download_button("⬇️ Descargar Reporte Excel", data=xlsx_data['data'], file_name=xlsx_data['name'], mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            
        with col_pdf_ctrl:
            if pdf_data:
                # Botón de descarga de PDF (si ya fue generado)
                st.download_button("⬇️ Descargar Gráficos PDF", data=pdf_data['data'], file_name=pdf_data['name'], mime="application/pdf")
            elif st.button("📈 Generar Gráficas PDF"):
                with st.spinner("Generando visualizaciones en PDF..."):
                    try:
                        pdf_path = generate_ratios_charts_pdf(os.path.join(st.session_state.state['dir'], xlsx_data['name']), st.session_state.state['dir'])
                        with open(pdf_path, "rb") as fh_pdf: 
                            st.session_state.state['pdf'] = {'data': fh_pdf.read(), 'name': os.path.basename(pdf_path)}
                        st.success("✅ Gráficas generadas. Descarga disponible.")
                        st.rerun() # Rerun para mostrar el botón de descarga
                    except Exception as e:
                        st.error(f"❌ Error al generar el PDF: {str(e)}")
else:
    reset_state()
    st.write("Sube los archivos Excel para comenzar el análisis.")
