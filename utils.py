import openpyxl as oxl, re, os, numpy as np
import matplotlib.pyplot as plt
from openpyxl.utils import column_index_from_string as col_idx_from_str, get_column_letter as col_let
from openpyxl.styles import NamedStyle
from matplotlib.backends.backend_pdf import PdfPages
from matplotlib.ticker import PercentFormatter
from datetime import datetime

# --- Funciones de Utilidad ---

def find_year_in_workbook(wb):
    """Busca el primer año (20XX) en las primeras 40 filas/12 columnas."""
    pattern = re.compile(r"\b(20\d{2})\b")
    for sheet in wb.worksheets:
        for row in sheet.iter_rows(max_row=40, max_col=12):
            for cell in row:
                if (val := cell.value) and (m := pattern.search(str(val))):
                    return int(m.group(1))
    raise ValueError("No se encontró año. Asegúrate que el archivo contiene el año (YYYY).")

def find_company_name(wb):
    """Extrae el nombre de la empresa de la celda 'A6'."""
    name = str(wb.active['A6'].value or "NOMBRE_DE_LA_EMPRESA").strip()
    return name.split(':', 1)[1].strip() if name.upper().startswith("EMPRESA:") else name

def read_range_values(ws, start_row, end_row, col='C'):
    """Lee valores de un rango de celdas en una columna específica."""
    col_idx = col_idx_from_str(col)
    return [ws.cell(r, col_idx).value or 0 for r in range(start_row, end_row + 1)]

def write_column_into_model(ws, start_row, values, col_idx):
    """Escribe una lista de valores en una columna del modelo, aplicando formato numérico."""
    for i, val in enumerate(values):
        r = start_row + i
        cell = ws.cell(r, col_idx, val)
        if r != 3 and isinstance(val, (int, float)):
            cell.number_format = "#,##0"

def is_file_open(filepath):
    """Verifica si un archivo está abierto (o no existe)."""
    if not os.path.exists(filepath): return False
    try:
        with open(filepath, 'a'): pass
        return False
    except IOError: return True

def safe_float(x):
    """Convierte un valor a float de forma segura, o retorna 0.0."""
    try: return float(x) if x is not None else 0.0
    except: return 0.0

# --- Funciones de Análisis y Reporte (Refactorizadas) ---

def process_files_and_generate_report(input_paths, model_path="BASE.xlsx", output_dir="."):
    """Procesa archivos, valida años, llena el modelo y genera el reporte Excel.
    
    CORRECCIÓN: Se elimina el uso de 'with' para la carga de archivos de entrada 
    y se utiliza try/finally para asegurar .close() y evitar el error del context manager.
    """
    if not input_paths: return

    model_wb = oxl.load_workbook(model_path)
    for name, fmt in [("percentage_style", "0.00%"), ("decimal_style", "0.00")]:
        try: model_wb.add_named_style(NamedStyle(name=name, number_format=fmt))
        except: pass

    # 1. Preparar datos
    first_wb = oxl.load_workbook(input_paths[0], data_only=False)
    company_name = find_company_name(first_wb); first_wb.close()
    
    entries, years = [], []
    for p in input_paths:
        wb = oxl.load_workbook(p, data_only=True) # CORRECCIÓN: Usar load_workbook sin 'with'
        try: 
            ws = wb.active
            year = int(str(ws["C12"].value or find_year_in_workbook(wb)).strip())
            years.append(year)
            entries.append({'year': year, 'bs': read_range_values(ws, 12, 88),
                            'is': read_range_values(ws, 91, 131),
                            'cf': read_range_values(ws, 185, 259)})
        finally:
            wb.close() # Asegurar que el archivo se cierra
            
    years_sorted = sorted(years, reverse=True)
    if len(set(years)) != len(years): raise ValueError("Años repetidos detectados.")
    if any(years_sorted[i] - years_sorted[i+1] != 1 for i in range(len(years_sorted)-1)):
        raise ValueError(f"Los años deben ser consecutivos. Detectados: {', '.join(map(str, years_sorted))}")

    # 2. Llenar el modelo con datos
    ws_info = model_wb['INFORMACIÓN']
    ws_info['B2'].value = f"Análisis Vertical y Horizontal - {company_name}"
    ws_info['B3'].value = f"{years_sorted[-1]}-{years_sorted[0]}" if years_sorted else "YYYY-YYYY"

    base_col, sheets = 3, {'ESTADO DE SITUACIÓN FINANCIERA': 'bs', 'ESTADO DE RESULTADOS': 'is', 'ESTADO DE FLUJO DE EFECTIVO': 'cf'}
    for idx, e in enumerate(sorted(entries, key=lambda e: e['year'], reverse=True)):
        for sheet_name, key in sheets.items():
            write_column_into_model(model_wb[sheet_name], 3, e[key], base_col + idx)

    # 3. Llenar el Análisis Vertical/Horizontal
    def vertical_formula(sheet, row, col, sheet_name):
        coord = sheet.cell(row, col).coordinate
        base_rows = {'ESTADO DE SITUACIÓN FINANCIERA': 40, 'ESTADO DE RESULTADOS': 4}
        if sheet_name in base_rows:
            base_coord = sheet.cell(base_rows[sheet_name], col).coordinate
            return f"=IFERROR({coord}/{base_coord},0)" if row != base_rows[sheet_name] else None
        if sheet_name == 'ESTADO DE FLUJO DE EFECTIVO':
            base_row = next((r for r, end in [(24, 24), (52, 52), (72, 72)] if row <= end), None)
            if base_row and row != base_row:
                return f"=IFERROR({coord}/{sheet.cell(base_row, col).coordinate},0)"
            return None

    report_config = [('ESTADO DE SITUACIÓN FINANCIERA', 3, 77), ('ESTADO DE RESULTADOS', 3, 41), ('ESTADO DE FLUJO DE EFECTIVO', 3, 75)]
    n_entries = len(entries)
    for sheet_name, start_row, rows_count in report_config:
        ws = model_wb[sheet_name]
        vert_start, hor_start = base_col + n_entries + 2, base_col + 2 * n_entries + 4
        ws.cell(2, vert_start, "ANÁLISIS VERTICAL"); ws.cell(2, hor_start, "ANÁLISIS HORIZONTAL")
        for i, y in enumerate(years_sorted): ws.cell(3, vert_start+i, str(y)).number_format = 'General'
        for i, y in enumerate(years_sorted[1:]): ws.cell(3, hor_start+i, str(y)).number_format = 'General'

        for r in range(start_row + 1, start_row + rows_count + 1):
            for j in range(n_entries):
                col_idx = base_col + j
                if formula := vertical_formula(ws, r, col_idx, sheet_name):
                    ws.cell(r, vert_start+j, formula).style = "percentage_style"
                if j + 1 < n_entries:
                    curr_coord, prev_coord = ws.cell(r, col_idx).coordinate, ws.cell(r, base_col+j+1).coordinate
                    formula = f"=IFERROR(({curr_coord}-{prev_coord})/{prev_coord},0)"
                    ws.cell(r, hor_start+j, formula).style = "percentage_style"

    # 4. Llenar la hoja RATIOS
    ws_ratios = model_wb['RATIOS']
    for merged in list(ws_ratios.merged_cells.ranges): ws_ratios.unmerge_cells(str(merged))
    for i, y in enumerate(years_sorted): ws_ratios.cell(3, 5+i, y).number_format = '0'
    
    ratio_formulas = {
        4: f"='ESTADO DE SITUACIÓN FINANCIERA'!{0}20/'ESTADO DE SITUACIÓN FINANCIERA'!{0}55",
        5: f"=('ESTADO DE SITUACIÓN FINANCIERA'!{0}20-'ESTADO DE SITUACIÓN FINANCIERA'!{0}13)/'ESTADO DE SITUACIÓN FINANCIERA'!{0}55",
        6: f"='ESTADO DE SITUACIÓN FINANCIERA'!{0}69/'ESTADO DE SITUACIÓN FINANCIERA'!{0}40",
        7: f"='ESTADO DE SITUACIÓN FINANCIERA'!{0}69/'ESTADO DE SITUACIÓN FINANCIERA'!{0}78",
        8: f"='ESTADO DE RESULTADOS'!{0}28/'ESTADO DE RESULTADOS'!{0}4",
        9: f"='ESTADO DE RESULTADOS'!{0}28/'ESTADO DE SITUACIÓN FINANCIERA'!{0}40",
        10: f"='ESTADO DE RESULTADOS'!{0}28/'ESTADO DE SITUACIÓN FINANCIERA'!{0}78",
        11: f"='ESTADO DE RESULTADOS'!{0}4/'ESTADO DE SITUACIÓN FINANCIERA'!{0}40",
    }
    
    for idx in range(len(years_sorted)):
        col, dc = 5 + idx, col_let(base_col + idx)
        pc = col_let(base_col + idx + 1) if idx < len(years_sorted) - 1 else dc
        
        for row, formula_template in ratio_formulas.items():
            style = "percentage_style" if row in [8, 9, 10] else "decimal_style"
            ws_ratios.cell(row, col, f"=IFERROR({formula_template.format(dc)},0)").style = style
        
        if idx < len(years_sorted) - 1:
            cxc_avg_cells = f"SUM('ESTADO DE SITUACIÓN FINANCIERA'!{dc}8:{dc}11)+SUM('ESTADO DE SITUACIÓN FINANCIERA'!{dc}24:{dc}27)+SUM('ESTADO DE SITUACIÓN FINANCIERA'!{pc}8:{pc}11)+SUM('ESTADO DE SITUACIÓN FINANCIERA'!{pc}24:{pc}27)"
            inv_avg_cells = f"('ESTADO DE SITUACIÓN FINANCIERA'!{dc}13+'ESTADO DE SITUACIÓN FINANCIERA'!{pc}13)+('ESTADO DE SITUACIÓN FINANCIERA'!{dc}30+'ESTADO DE SITUACIÓN FINANCIERA'!{pc}30)"
            ws_ratios.cell(12, col, f"=IFERROR('ESTADO DE RESULTADOS'!{dc}4/(({cxc_avg_cells})/2),0)").style = "decimal_style"
            ws_ratios.cell(13, col, f"=IFERROR(-'ESTADO DE RESULTADOS'!{dc}5/(({inv_avg_cells})/2),0)").style = "decimal_style"
        else:
            for row in [12, 13]: ws_ratios.cell(row, col, 0).style = "decimal_style"

    # 5. Guardar y retornar
    empresa_limpia = re.sub(r'[<>:"/\\|?*]', '', company_name)[:50]
    base_filename = f"REPORTE_{empresa_limpia}_{years_sorted[-1]}-{years_sorted[0]}.xlsx"
    out_file = os.path.join(output_dir, base_filename)
    
    if is_file_open(out_file) or os.path.exists(out_file):
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"REPORTE_{empresa_limpia}_{years_sorted[-1]}-{years_sorted[0]}_{timestamp}.xlsx"
        out_file = os.path.join(output_dir, filename)

    model_wb.save(out_file)
    model_wb.close()
    return out_file

def _add_value_labels(ax, rects, values, ttype):
    """Auxiliar para agregar etiquetas de valor a las barras del gráfico."""
    y_min, y_max = ax.get_ylim()
    for rect, v in zip(rects, values):
        sign = 1 if v >= 0 else -1
        text = f"{v*100:.2f}%" if ttype == 'pct' else f"{v:.2f}"
        ax.text(rect.get_x() + rect.get_width()/2, v + (y_max - y_min) * 0.02 * sign, 
                text, ha='center', va='bottom' if sign >= 0 else 'top', fontsize=9, fontweight='bold')

def _nice_ticks(ax, values, ttype, n_ticks=6):
    """Auxiliar para establecer límites y formato de los ejes Y de forma legible."""
    v_max, v_min = max(values + [0]), min(values + [0])
    span = v_max - v_min
    if span == 0: v_max, v_min = v_max + abs(v_max)*0.1 + 0.1, v_min - abs(v_min)*0.1 - 0.1; span = v_max - v_min
    
    ax.set_ylim(v_min - span*0.05, v_max + span*0.15)
    ax.set_yticks(np.linspace(v_min - span*0.05, v_max + span*0.15, n_ticks))
    ax.yaxis.set_major_formatter(PercentFormatter(1.0) if ttype == 'pct' else plt.FuncFormatter(lambda x, pos: f"{x:.2f}"))

def generate_ratios_charts_pdf(report_path, output_dir):
    """Lee el reporte Excel y genera un PDF con gráficos de ratios.
    
    CORRECCIÓN: Se elimina el uso de 'with' para la carga del archivo de reporte
    y se utiliza try/finally para asegurar .close() y evitar el error del context manager.
    """
    plt.rcParams.update({'font.sans-serif': 'Arial', 'font.size': 10})
    
    wb = oxl.load_workbook(report_path, data_only=True) # CORRECCIÓN: Usar load_workbook sin 'with'
    try:
        ws_ratios, ws_bs, ws_is, ws_info = wb['RATIOS'], wb['ESTADO DE SITUACIÓN FINANCIERA'], wb['ESTADO DE RESULTADOS'], wb['INFORMACIÓN']
        
        company_name = str(ws_info['B2'].value or "EMPRESA").replace("Análisis Vertical y Horizontal - ", "").strip()
        empresa_limpia = re.sub(r'[<>:"/\\|?*]', '', company_name)[:50]

        years = [int(ws_ratios.cell(3, c).value) for c in range(5, ws_ratios.max_column + 1) if ws_ratios.cell(3, c).value]
        if not years: raise ValueError("No se encontraron años en la hoja RATIOS.")

        ratios_info = [
            {'row': r, 'name': n, 'short': s, 'type': t, 'category': c} for r, n, s, t, c in [
                (4, 'Liquidez corriente', 'Liquidez corriente', 'ratio', 'Liquidez'), (5, 'Prueba ácida', 'Prueba ácida', 'ratio', 'Liquidez'),
                (6, 'Razón de deuda total', 'Razón deuda total', 'ratio', 'Endeudamiento'), (7, 'Razón deuda/patrimonio', 'Razón deuda/patrimonio', 'ratio', 'Endeudamiento'),
                (8, 'Margen neto', 'Margen neto', 'pct', 'Rentabilidad'), (9, 'ROA', 'ROA', 'pct', 'Rentabilidad'), (10, 'ROE', 'ROE', 'pct', 'Rentabilidad'),
                (11, 'Rotación de activos totales', 'Rotación activos', 'ratio', 'Actividad'), (12, 'Rotación de cuentas por cobrar', 'Rotación CxC', 'ratio', 'Actividad'),
                (13, 'Rotación de inventarios', 'Rotación invent.', 'ratio', 'Actividad'),
            ]
        ]
        
        # Recálculo de Ratios (para asegurar la data_only)
        all_ratios = {r['row']: [] for r in ratios_info}
        def get_val(ws, r, c): return safe_float(ws.cell(r, 3 + c).value)
        
        for i in range(len(years)):
            bs = {r: get_val(ws_bs, r, i) for r in [13, 20, 30, 40, 55, 69, 78]}
            is_ = {r: get_val(ws_is, r, i) for r in [4, 5, 28]}
            
            # Ratios simples
            all_ratios[4].append(bs[20] / bs[55] if bs[55] else 0)
            all_ratios[5].append((bs[20] - bs[13]) / bs[55] if bs[55] else 0)
            all_ratios[6].append(bs[69] / bs[40] if bs[40] else 0)
            all_ratios[7].append(bs[69] / bs[78] if bs[78] else 0)
            all_ratios[8].append(is_[28] / is_[4] if is_[4] else 0)
            all_ratios[9].append(is_[28] / bs[40] if bs[40] else 0)
            all_ratios[10].append(is_[28] / bs[78] if bs[78] else 0)
            all_ratios[11].append(is_[4] / bs[40] if bs[40] else 0)
            
            # Ratios con promedio (requieren año anterior, si existe)
            if i < len(years) - 1:
                cxc_sum = sum(get_val(ws_bs, r, i) + get_val(ws_bs, r, i+1) for r in list(range(8, 12)) + list(range(24, 28)))
                all_ratios[12].append(is_[4] / (cxc_sum / 2) if cxc_sum else 0)
                inv_sum = bs[13] + get_val(ws_bs, 13, i+1) + bs[30] + get_val(ws_bs, 30, i+1)
                all_ratios[13].append((-is_[5]) / (inv_sum / 2) if inv_sum else 0)
            else:
                all_ratios[12].append(0); all_ratios[13].append(0)


        # Generación del PDF
        pdf_path = os.path.join(output_dir, f"GRAFICOS_{empresa_limpia}_{min(years)}-{max(years)}.pdf")
        with PdfPages(pdf_path) as pdf:
            # Página de Título
            fig = plt.figure(figsize=(11.69, 8.27))
            for y_pos, text, size in [(0.7, "Análisis Detallado de Ratios Financieros", 24), (0.6, company_name, 20),
                                     (0.5, f"Periodo: {min(years)} - {max(years)}", 18), (0.3, "Generado por Analizador Financiero SMV", 12)]:
                fig.text(0.5, y_pos, text, ha='center', fontsize=size)
            pdf.savefig(fig); plt.close(fig)
            
            # Páginas de Gráficos
            groups = [{'title': 'Análisis de Liquidez y Endeudamiento', 'categories': ['Liquidez', 'Endeudamiento'], 'layout': (2, 2)},
                      {'title': 'Análisis de Rentabilidad y Actividad', 'categories': ['Rentabilidad', 'Actividad'], 'layout': (2, 3)}]
            
            for group in groups:
                group_ratios = [r for r in ratios_info if r['category'] in group['categories']]
                n_rows, n_cols = group['layout']
                fig, axes = plt.subplots(n_rows, n_cols, figsize=(11.69, 8.27), squeeze=False)
                fig.suptitle(group['title'], fontsize=18, y=0.98, fontweight='bold', color='darkslategrey')
                
                for i, r in enumerate(group_ratios):
                    ax = axes[i // n_cols, i % n_cols]
                    values = all_ratios[r['row']]
                    
                    # Rotación de CxC y Inventarios solo para n_años - 1
                    x_years = years if r['row'] < 12 else years[:-1] 
                    x_values = values if r['row'] < 12 else values[:-1]
                    
                    rects = ax.bar(np.arange(len(x_years)), x_values, width=0.7, color='#1E90FF', edgecolor='black', linewidth=0.5, alpha=0.95)
                    ax.set_xticks(np.arange(len(x_years)))
                    ax.set_xticklabels([str(y) for y in x_years], fontsize=10, fontweight='medium')
                    ax.set_title(r['name'], fontsize=12, fontweight='bold', color='darkslategrey')
                    _nice_ticks(ax, x_values, r['type'])
                    ax.set_ylabel(r['short'], fontsize=10, color='grey')
                    ax.grid(axis='y', linestyle='-', alpha=0.4)
                    _add_value_labels(ax, rects, x_values, r['type'])
                
                for i in range(len(group_ratios), n_rows * n_cols):
                    fig.delaxes(axes[i // n_cols, i % n_cols])

                plt.tight_layout(rect=[0, 0.05, 1, 0.95])
                pdf.savefig(fig); plt.close(fig)

        return pdf_path
    finally:
        wb.close() # Asegurar que el archivo se cierra al finalizar o si hay error
