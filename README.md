# Analizador Financiero SMV (Python / Streamlit)

Proyecto que toma libros SMV (uno por año), extrae rangos específicos y los pega en una plantilla BASE.xlsx
generando análisis vertical, horizontal y una hoja de ratios básica.

## Cómo usar
1. Instala dependencias: `pip install -r requirements.txt`
2. Ejecuta: `streamlit run app.py`
3. Sube los libros SMV (xlsx) — uno por año — y presiona "Generar reporte".
4. Descarga el Excel generado `REPORTE_ANALISIS_FINANCIERO.xlsx`.

## Notas y supuestos
- El archivo `BASE.xlsx` incluido contiene las hojas:
  - INFORMACIÓN
  - ESTADO DE SITUACIÓN FINANCIERA
  - ESTADO DE RESULTADOS
  - ESTADO DE FLUJO DE EFECTIVO
  - RATIOS
- El programa busca el año (YYYY) dentro de cada libro. Asegúrate que los archivos contengan el año en formato 4 dígitos en alguna celda.
- Rangos leídos de cada archivo (según especificación):
  - Estado de situación financiera: `C12:C88`
  - Estado de resultados: `C91:C131`
  - Estado de flujo de efectivo: `C185:C259`
- Los años deben ser consecutivos y no repetidos.
- El análisis vertical y horizontal se calcula y escribe en columnas a la derecha de los datos.
- Este es un scaffold funcional; puedes adaptar fórmulas exactas y posicionamiento según el modelo original.

## Estructura
- app.py — interfaz Streamlit
- utils.py — lógica de lectura/escritura/validaciones
- BASE.xlsx — plantilla modelo (proporcionada)
- requirements.txt