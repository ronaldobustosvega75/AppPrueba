import streamlit as st

#Estilos de la app
def load_styles():
    st.markdown("""
        <style>
            body {background-color: #f4f6f8;}
            .main {background-color: #fff; border-radius: 18px; padding: 30px;}
            h1 {color: #003366;}
            .subtitle {color: #555; font-size: 1.2em;}
            .stButton>button {
                background-color: #003366;
                color: white;
                font-size: 1.1em;
                border-radius: 8px;
                padding: 10px 24px;
                border: none;
            }
            .stButton>button:hover {
                background-color: #00509e;
                color: #fff;
            }
                
            /* Botón de descarga */
            .stDownloadButton>button {
                background-color: #28a745; /* verde */
                color: white;
                font-size: 1.1em;
                border-radius: 8px;
                padding: 10px 24px;
                border: none;
                font-weight: bold;
                box-shadow: 0 4px 8px rgba(0,0,0,0.2);
            }
            .stDownloadButton>button:hover {
                background-color: #218838; /* verde más oscuro */
                color: #fff;
            }
   
                    
            .stFileUploader label {font-weight: bold;}
        </style>
    """, unsafe_allow_html=True)

# Alertas
def show_alert(message, alert_type="success"):
    colors = {
        "success": "#d4edda",
        "error": "#f8d7da",
        "warning": "#fff3cd",
        "info": "#d1ecf1"
    }
    border_colors = {
        "success": "#28a745",
        "error": "#dc3545",
        "warning": "#ffc107",
        "info": "#17a2b8"
    }
    color = colors.get(alert_type, "#d1ecf1")
    border = border_colors.get(alert_type, "#17a2b8")

    st.markdown(
        f"""
        <div style="padding:15px; border-radius:10px; background-color:{color};
                    border-left:8px solid {border}; margin:10px 0; font-size:16px;">
            {message}
        </div>
        """,
        unsafe_allow_html=True
    )