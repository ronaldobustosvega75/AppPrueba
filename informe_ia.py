import textwrap
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
import os

try:
    from google import genai
    GENAI_AVAILABLE = True
except ImportError:
    GENAI_AVAILABLE = False
    print("⚠️ Google GenAI no está instalado. Instalalo con: pip install google-genai")

GEMINI_API_KEY = "AIzaSyB2z0TruIEMVEaFWiUPHJ0JX0XZeSgZIJI"  

def generar_informe_ia(years, all_ratios):
    if not GENAI_AVAILABLE:
        return generar_informe_local(years, all_ratios)
    
    if not GEMINI_API_KEY or GEMINI_API_KEY == "tu_nueva_api_key_aqui":
        return generar_informe_local(years, all_ratios)
    
    try:
        # Configurar cliente GenAI
        client = genai.Client(api_key=GEMINI_API_KEY)
        
        resumen = ""
        for i, year in enumerate(years):
            resumen += f"\nAño {year}:\n"
            resumen += f"- Liquidez Corriente: {all_ratios[4][i]:.2f}\n"
            resumen += f"- Prueba Ácida: {all_ratios[5][i]:.2f}\n"
            if 6 in all_ratios:
                resumen += f"- Razón Deuda Total: {all_ratios[6][i]:.2f}\n"
            resumen += f"- Deuda/Patrimonio: {all_ratios[7][i]:.2f}\n"
            resumen += f"- Margen Neto: {all_ratios[8][i]*100:.2f}%\n"
            resumen += f"- ROA: {all_ratios[9][i]*100:.2f}%\n"
            resumen += f"- ROE: {all_ratios[10][i]*100:.2f}%\n"
            resumen += f"- Rotación de Activos: {all_ratios[11][i]:.2f}\n"

        prompt = f"""
                Eres un analista financiero profesional.
                Elabora un informe narrativo exclusivamente de análisis, evitando definiciones teóricas de los ratios.

                Estructura el informe en las siguientes secciones:
                1. Análisis de Liquidez
                2. Análisis de Endeudamiento
                3. Análisis de Rentabilidad
                4. Análisis de Eficiencia Operativa
                5. Resumen y Recomendaciones

                Instrucciones:
                - Escribe el informe en texto plano, sin aplicar formato de negritas, cursivas ni encabezados Markdown.
                - No uses asteriscos, guiones, ni ningún carácter especial para resaltar texto.
                - Usa únicamente subtítulos simples escritos de forma normal, por ejemplo: "Análisis de Liquidez".
                - Analiza las tendencias comparando los años entre sí.
                - Señala fortalezas y debilidades en cada área.
                - En la sección de Recomendaciones, ofrece sugerencias concretas y viables (ej. reducción de deuda, mejora en gestión de activos, control de costos, políticas de inversión, etc.).
                - El texto debe ser claro, directo, en español, con párrafos bien estructurados y tono formal.
                - Evita definiciones académicas de los ratios; enfócate solo en la interpretación y el impacto en la empresa.
                - El informe debe tener estilo narrativo, con subtítulos simples para cada sección.
                - Máximo 1200 palabras.

                Datos de ratios financieros:
                {resumen}

                Proporciona el informe completo en español, sin Markdown ni símbolos de formato.
                """


        response = client.models.generate_content(
            model="gemini-2.0-flash-exp",  # Modelo más reciente
            contents=prompt
        )
        
        print("Análisis generado con Gemini 2.0")
        return response.text
        
    except Exception as e:
        print(f"Error con GenAI: {e}")
        print("Usando análisis local como respaldo...")
        return generar_informe_local(years, all_ratios)

def generar_informe_local(years, all_ratios):
    try:
        informe = "ANÁLISIS FINANCIERO DETALLADO\n\n"
        
        informe += "1. ANÁLISIS DE LIQUIDEZ\n\n"
        
        liquidez_corriente = [all_ratios[4][i] for i in range(len(years))]
        prueba_acida = [all_ratios[5][i] for i in range(len(years))]
        
        if len(liquidez_corriente) > 1:
            tendencia_lc = "mejorando" if liquidez_corriente[0] > liquidez_corriente[-1] else "deteriorándose"
            informe += f"La liquidez corriente muestra una tendencia {tendencia_lc} en el período analizado. "
        
        valor_actual_lc = liquidez_corriente[0] if liquidez_corriente else 0
        if valor_actual_lc > 2.0:
            informe += "La empresa mantiene una posición de liquidez sólida, con capacidad suficiente para cubrir sus obligaciones a corto plazo. "
        elif valor_actual_lc > 1.0:
            informe += "La empresa presenta una liquidez adecuada, aunque podría mejorar su gestión de activos corrientes. "
        else:
            informe += "La empresa enfrenta desafíos de liquidez que requieren atención inmediata. "
        
        informe += f"El ratio actual es de {valor_actual_lc:.2f}.\n\n"
        
        informe += "2. ANÁLISIS DE ENDEUDAMIENTO\n\n"
        
        deuda_patrimonio = [all_ratios[7][i] for i in range(len(years))]
        valor_actual_dp = deuda_patrimonio[0] if deuda_patrimonio else 0
        
        if valor_actual_dp > 1.0:
            informe += f"La empresa presenta un alto nivel de endeudamiento con un ratio deuda/patrimonio de {valor_actual_dp:.2f}, "
            informe += "lo que indica que la deuda supera al patrimonio. Esto puede representar un riesgo financiero elevado. "
        elif valor_actual_dp > 0.5:
            informe += f"El nivel de endeudamiento es moderado ({valor_actual_dp:.2f}), "
            informe += "manteniendo un equilibrio razonable entre deuda y patrimonio. "
        else:
            informe += f"La empresa mantiene un endeudamiento conservador ({valor_actual_dp:.2f}), "
            informe += "con bajo riesgo financiero pero posiblemente subutilizando el apalancamiento. "
        
        informe += "\n\n"
        
        informe += "3. ANÁLISIS DE RENTABILIDAD\n\n"
        
        margen_neto = [all_ratios[8][i] * 100 for i in range(len(years))]
        roa = [all_ratios[9][i] * 100 for i in range(len(years))]
        roe = [all_ratios[10][i] * 100 for i in range(len(years))]
        
        valor_actual_mn = margen_neto[0] if margen_neto else 0
        valor_actual_roa = roa[0] if roa else 0
        valor_actual_roe = roe[0] if roe else 0
        
        informe += f"El margen neto actual es del {valor_actual_mn:.2f}%, "
        if valor_actual_mn > 10:
            informe += "indicando una excelente eficiencia en la generación de beneficios. "
        elif valor_actual_mn > 5:
            informe += "mostrando una rentabilidad satisfactoria. "
        elif valor_actual_mn > 0:
            informe += "reflejando una rentabilidad baja que requiere mejoras. "
        else:
            informe += "evidenciando pérdidas que necesitan atención urgente. "
        
        informe += f"\n\nEl ROA ({valor_actual_roa:.2f}%) y ROE ({valor_actual_roe:.2f}%) "
        if valor_actual_roe > valor_actual_roa:
            informe += "muestran que el apalancamiento financiero está beneficiando a los accionistas. "
        else:
            informe += "indican que el apalancamiento podría no estar siendo utilizado eficientemente. "
        
        informe += "\n\n"
        
        informe += "4. ANÁLISIS DE EFICIENCIA OPERATIVA\n\n"
        
        rotacion_activos = [all_ratios[11][i] for i in range(len(years))]
        rotacion_cxc = [all_ratios[12][i] for i in range(len(years))]
        
        valor_actual_ra = rotacion_activos[0] if rotacion_activos else 0
        valor_actual_cxc = rotacion_cxc[0] if rotacion_cxc else 0
        
        informe += f"La rotación de activos totales ({valor_actual_ra:.2f}) "
        if valor_actual_ra > 1.0:
            informe += "demuestra una buena eficiencia en el uso de activos para generar ventas. "
        else:
            informe += "sugiere oportunidades de mejora en la utilización de activos. "
        
        informe += f"La rotación de cuentas por cobrar ({valor_actual_cxc:.2f}) "
        if valor_actual_cxc > 6:
            informe += "indica una excelente gestión de cobranzas. "
        elif valor_actual_cxc > 4:
            informe += "muestra una gestión adecuada del crédito. "
        else:
            informe += "sugiere la necesidad de mejorar las políticas de cobranza. "
        
        informe += "\n\n5. RESUMEN Y RECOMENDACIONES\n\n"
        
        puntos_fuertes = []
        areas_mejora = []
        
        if valor_actual_lc > 1.5:
            puntos_fuertes.append("sólida posición de liquidez")
        else:
            areas_mejora.append("gestión de liquidez")
            
        if valor_actual_dp < 0.6:
            puntos_fuertes.append("nivel de endeudamiento conservador")
        elif valor_actual_dp > 1.0:
            areas_mejora.append("control del endeudamiento")
            
        if valor_actual_mn > 5:
            puntos_fuertes.append("rentabilidad satisfactoria")
        else:
            areas_mejora.append("mejora de márgenes de rentabilidad")
            
        if valor_actual_ra > 0.8:
            puntos_fuertes.append("eficiencia en el uso de activos")
        else:
            areas_mejora.append("optimización del uso de activos")
        
        if puntos_fuertes:
            informe += "FORTALEZAS IDENTIFICADAS:\n"
            for punto in puntos_fuertes:
                informe += f"• {punto.capitalize()}\n"
            informe += "\n"
        
        if areas_mejora:
            informe += "ÁREAS DE MEJORA:\n"
            for area in areas_mejora:
                informe += f"• {area.capitalize()}\n"
            informe += "\n"
        
        return informe
        
    except Exception as e:
        return f"Error al generar el análisis financiero: {str(e)}"


def exportar_informe_pdf(informe_texto, output_dir="."):
    out_file = os.path.join(output_dir, "INFORME_DETALLADO_RATIOS.pdf")

    with PdfPages(out_file) as pdf:
        plt.rcParams.update({'font.size': 10, 'font.family': 'serif'})
        
        fig = plt.figure(figsize=(8.27, 11.69))  # A4 vertical
        
        fig.text(0.5, 0.95, "INFORME DE ANÁLISIS FINANCIERO", 
                ha='center', va='top', fontsize=16, fontweight='bold')
        
        fig.text(0.5, 0.92, "Análisis Detallado de Ratios Financieros", 
                ha='center', va='top', fontsize=12, style='italic')
        
        lineas = informe_texto.split('\n')
        y_position = 0.88
        line_height = 0.025
        
        for linea in lineas:
            # Si la línea es muy larga, dividirla
            if len(linea) > 80:
                sub_lineas = textwrap.wrap(linea, 80)
                for sub_linea in sub_lineas:
                    # Verificar si necesitamos nueva página
                    if y_position < 0.05:
                        pdf.savefig(fig, bbox_inches='tight')
                        plt.close(fig)
                        fig = plt.figure(figsize=(8.27, 11.69))
                        y_position = 0.95
                    
                    # Aplicar formato según el contenido
                    if sub_linea.strip().endswith(':') or sub_linea.strip().isupper():
                        fig.text(0.1, y_position, sub_linea, ha='left', va='top', 
                                fontweight='bold', fontsize=11)
                    elif sub_linea.strip().startswith('•'):
                        fig.text(0.12, y_position, sub_linea, ha='left', va='top', 
                                fontsize=10)
                    else:
                        fig.text(0.1, y_position, sub_linea, ha='left', va='top', 
                                fontsize=10)
                    
                    y_position -= line_height
            else:
                # Verificar si necesitamos nueva página
                if y_position < 0.05:
                    pdf.savefig(fig, bbox_inches='tight')
                    plt.close(fig)
                    fig = plt.figure(figsize=(8.27, 11.69))
                    y_position = 0.95
                
                # Aplicar formato según el contenido
                if linea.strip().endswith(':') or linea.strip().isupper():
                    fig.text(0.1, y_position, linea, ha='left', va='top', 
                            fontweight='bold', fontsize=11)
                elif linea.strip().startswith('•'):
                    fig.text(0.12, y_position, linea, ha='left', va='top', 
                            fontsize=10)
                else:
                    fig.text(0.1, y_position, linea, ha='left', va='top', 
                            fontsize=10)
                
                y_position -= line_height
        
        # Guardar la última página
        pdf.savefig(fig, bbox_inches='tight')
        plt.close(fig)

    return out_file