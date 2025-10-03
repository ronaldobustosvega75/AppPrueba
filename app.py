import streamlit as st
from utils import process_files_and_generate_report, generate_ratios_charts_pdf
import tempfile, os, shutil, re
from style import load_styles, show_alert

hide_streamlit_style = """
    <style>
    footer {visibility: hidden;}
    header {visibility: hidden;}
    .stDeployButton {display:none;}
    </style>
"""
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

st.set_page_config(page_title="Analizador Financiero SMV", layout="wide")
load_styles()

st.markdown("""
    <div class="main">
        <h1 style='font-size:2.8em; font-family:Arial, sans-serif; text-align:center; margin-bottom:10px;'>
            📊 Analizador Financiero <span style='font-size:0.7em;'>(Análisis Vertical, Horizontal y Ratios)</span>
        </h1>
        <p class="subtitle" style='text-align:center; margin-bottom:30px;'>
            Estimado usuario: Carga los archivos de los estados financieros por año (uno por año).<br>
            Los archivos deben ser consecutivos en años, empezando por <b>2024</b>.
        </p>
    </div>
""", unsafe_allow_html=True)

# Estado de sesión
if 'state' not in st.session_state:
    st.session_state['state'] = {
        'temp_dir': None,
        'report_path': None,
        'pdf_path': None,
        'excel_data': None,
        'pdf_data': None,
        'excel_filename': None,
        'pdf_filename': None
    }

def clear_session_files():
    """Limpia archivos temporales y resetea el estado"""
    if st.session_state['state']['temp_dir'] and os.path.exists(st.session_state['state']['temp_dir']):
        shutil.rmtree(st.session_state['state']['temp_dir'], ignore_errors=True)
    st.session_state['state'] = {
        'temp_dir': None,
        'report_path': None,
        'pdf_path': None,
        'excel_data': None,
        'pdf_data': None,
        'excel_filename': None,
        'pdf_filename': None
    }

def validar_archivos(archivos):
    """Valida que los archivos sean consecutivos en años"""
    if not all(f.name.endswith('.xlsx') for f in archivos):
        return False, "Todos los archivos deben ser en formato .xlsx", "error", []
    
    try:
        anios = []
        archivos_con_anio = []
        for f in archivos:
            match = re.search(r'(\d{4})', f.name)
            if not match:
                return False, "Los nombres de los archivos deben contener un año válido (ejemplo: Leche_Gloria_2024.xlsx).", "warning", []
            anio = int(match.group(1))
            anios.append(anio)
            archivos_con_anio.append((anio, f))
        
        if len(anios) < 3:
            return False, "Debes subir al menos 3 archivos.", "warning", []
    except ValueError:
        return False, "Los nombres de los archivos deben contener un año válido.", "error", []
    
    # Ordenar por año descendente
    archivos_ordenados = sorted(archivos_con_anio, key=lambda x: x[0], reverse=True)
    anios_ordenados = [a for a, _ in archivos_ordenados]
    
    # Verificar que sean consecutivos
    consecutivos = all(anios_ordenados[i] - anios_ordenados[i + 1] == 1 for i in range(len(anios_ordenados) - 1))
    
    if not consecutivos:
        return False, f"Los archivos deben ser de años consecutivos. Se detectaron: {', '.join(map(str, anios_ordenados))}", "warning", []
    
    return True, f"✅ Archivos validados correctamente: {', '.join(map(str, anios_ordenados))}", "success", [f for _, f in archivos_ordenados]

# Subida de archivos
uploaded = st.file_uploader(
    "📂 Sube los archivos SMV de tu empresa (mínimo 3 años)",
    accept_multiple_files=True,
    type=['xlsx', 'xlsm', 'xls'],
    key="uploader"
)

st.info("💡 El libro modelo (BASE.xlsx) ya está incluido en la aplicación y se usará como plantilla.")

if uploaded:
    n = len(uploaded)
    st.markdown(f"<b>📁 {n} archivo(s) cargado(s)</b>", unsafe_allow_html=True)
    
    # Validar archivos
    validacion, mensaje, tipo, archivos_ordenados = validar_archivos(uploaded)
    
    if not validacion:
        show_alert(mensaje, tipo)
        clear_session_files()
    else:
        show_alert(mensaje, "success")
        
        # Botón para generar Excel
        if st.button("🚀 Generar Reporte Completo (Excel)", type="primary"):
            clear_session_files()
            
            with st.spinner("⏳ Procesando archivos y generando reporte Excel..."):
                try:
                    tmpdir = tempfile.mkdtemp()
                    paths = []
                    
                    # Guardar archivos temporalmente
                    for f in archivos_ordenados:
                        p = os.path.join(tmpdir, f.name)
                        with open(p, "wb") as out:
                            out.write(f.getbuffer())
                        paths.append(p)
                    
                    # Generar reporte
                    out_path = process_files_and_generate_report(
                        paths,
                        model_path="BASE.xlsx",
                        output_dir=tmpdir
                    )
                    
                    # Guardar en sesión
                    with open(out_path, "rb") as fh:
                        st.session_state['state']['excel_data'] = fh.read()
                    
                    st.session_state['state'].update({
                        'temp_dir': tmpdir,
                        'report_path': out_path,
                        'excel_filename': os.path.basename(out_path)
                    })
                    
                    st.success("✅ Reporte Excel generado exitosamente. Puedes descargarlo a continuación.")
                    st.rerun()
                    
                except Exception as e:
                    st.error(f"❌ Error al generar el reporte: {str(e)}")
                    import traceback
                    st.code(traceback.format_exc())
                    clear_session_files()
        
        # Sección de descargas
        if st.session_state['state']['excel_data']:
            st.markdown("---")
            st.subheader("📥 Descargar Resultados")
            
            col1, col2, col3 = st.columns(3)
            
            # Columna 1: Descargar Excel
            with col1:
                st.download_button(
                    label=f"⬇️ Descargar Reporte Excel",
                    data=st.session_state['state']['excel_data'],
                    file_name=st.session_state['state']['excel_filename'],
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    help="Reporte con Análisis Vertical, Horizontal y Ratios",
                    use_container_width=True
                )
            
            # Columna 2: Generar/Descargar PDF
            with col2:
                if st.session_state['state']['pdf_data']:
                    st.download_button(
                        label=f"⬇️ Descargar Análisis PDF",
                        data=st.session_state['state']['pdf_data'],
                        file_name=st.session_state['state']['pdf_filename'],
                        mime="application/pdf",
                        help="PDF con gráficos de ratios y análisis narrativo",
                        use_container_width=True
                    )
                else:
                    if st.button("📊 Generar Análisis PDF", type="secondary", use_container_width=True):
                        with st.spinner("⏳ Generando PDF con gráficos y análisis con IA..."):
                            try:
                                report_path = st.session_state['state']['report_path']
                                
                                # Verificar que el archivo existe
                                if not os.path.exists(report_path):
                                    # Recrear desde datos en memoria
                                    report_path = os.path.join(
                                        st.session_state['state']['temp_dir'],
                                        st.session_state['state']['excel_filename']
                                    )
                                    with open(report_path, "wb") as fh:
                                        fh.write(st.session_state['state']['excel_data'])
                                    st.session_state['state']['report_path'] = report_path
                                
                                # Generar PDF
                                pdf_path = generate_ratios_charts_pdf(
                                    report_path,
                                    st.session_state['state']['temp_dir']
                                )
                                
                                # Guardar en sesión
                                with open(pdf_path, "rb") as fh_pdf:
                                    st.session_state['state']['pdf_data'] = fh_pdf.read()
                                
                                st.session_state['state'].update({
                                    'pdf_path': pdf_path,
                                    'pdf_filename': os.path.basename(pdf_path)
                                })
                                
                                st.success("✅ Análisis PDF generado exitosamente.")
                                st.rerun()
                                
                            except Exception as e:
                                st.error(f"❌ Error al generar PDF: {str(e)}")
                                import traceback
                                st.code(traceback.format_exc())
            
            # Columna 3: Nuevo análisis
            with col3:
                if st.button("🔄 Nuevo Análisis", use_container_width=True):
                    clear_session_files()
                    st.rerun()

else:
    clear_session_files()
    st.write("👆 Aún no has subido archivos. Por favor, selecciona los archivos Excel para comenzar el análisis.")

