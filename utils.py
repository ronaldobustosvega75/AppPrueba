import openpyxl, re, os, numpy as np, matplotlib.pyplot as plt
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.styles import NamedStyle
from matplotlib.backends.backend_pdf import PdfPages 
from matplotlib.ticker import PercentFormatter

def find_year_in_workbook(wb):
    pattern = re.compile(r"\b(20\d{2})\b")
    for sheet in wb.worksheets:
        for row in sheet.iter_rows(min_row=1, max_row=40, min_col=1, max_col=12):
            for cell in row:
                if cell.value and (m := pattern.search(str(cell.value))):
                    return int(m.group(1))
    raise ValueError("No se encontró año en el libro. Asegúrate que el archivo contiene el año (YYYY).")

def read_range_values(wb, sheet_name, start_row, end_row, col='C'):
    ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.active
    return [ws.cell(row=r, column=column_index_from_string(col)).value or 0 for r in range(start_row, end_row+1)]

def write_column_into_model(ws, target_start_row, values, col_idx):
    for i, val in enumerate(values):
        r = target_start_row + i
        cell = ws.cell(row=r, column=col_idx)
        cell.value = val
        if r != 3 and isinstance(val, (int, float)):
            cell.number_format = "#,##0"

def process_files_and_generate_report(input_paths, model_path="BASE.xlsx", output_dir="."):
    model_wb = openpyxl.load_workbook(model_path)

    for style_name, fmt in [("percentage_style", "0.00%"), ("decimal_style", "0.00")]:
        try: model_wb.add_named_style(NamedStyle(name=style_name, number_format=fmt))
        except: pass

    entries, years = [], []
    for p in input_paths:
        wb = openpyxl.load_workbook(p, data_only=True)
        year = int(str(wb.active["C12"].value or find_year_in_workbook(wb)).strip())
        years.append(year)
        entries.append({'year': year, 'bs': read_range_values(wb, wb.sheetnames[0], 12, 88, 'C'),
                       'is': read_range_values(wb, wb.sheetnames[0], 91, 131, 'C'),
                       'cf': read_range_values(wb, wb.sheetnames[0], 185, 259, 'C')})
        wb.close()

    years_sorted = sorted(years, reverse=True)
    if len(set(years)) != len(years): raise ValueError("Años repetidos detectados.")
    for i in range(len(years_sorted)-1):
        if years_sorted[i] - years_sorted[i+1] != 1:
            raise ValueError(f"Los años deben ser consecutivos. Detectados: {', '.join(map(str, years_sorted))}")

    base_col = 3
    sheets = {'ESTADO DE SITUACIÓN FINANCIERA': 'bs', 'ESTADO DE RESULTADOS': 'is', 'ESTADO DE FLUJO DE EFECTIVO': 'cf'}
    for idx, e in enumerate(sorted(entries, key=lambda e: e['year'], reverse=True)):
        for sheet_name, key in sheets.items():
            write_column_into_model(model_wb[sheet_name], 3, e[key], base_col + idx)

    def vertical_formula(sheet, row, col, sheet_name):
        coord = sheet.cell(row=row, column=col).coordinate
        base_rows = {'ESTADO DE SITUACIÓN FINANCIERA': 40, 'ESTADO DE RESULTADOS': 4}
        if sheet_name in base_rows:
            return f"=IFERROR({coord}/{sheet.cell(row=base_rows[sheet_name], column=col).coordinate},0)" if row != base_rows[sheet_name] else None
        if sheet_name == 'ESTADO DE FLUJO DE EFECTIVO':
            base_row = 24 if row <= 24 else (52 if row <= 52 else (72 if row <= 72 else None))
            return f"=IFERROR({coord}/{sheet.cell(row=base_row, column=col).coordinate},0)" if base_row and row != base_row else None

    n_entries = len(entries)
    for sheet_name, start_row, rows_count in [('ESTADO DE SITUACIÓN FINANCIERA', 3, 77), 
                                                ('ESTADO DE RESULTADOS', 3, 41), 
                                                ('ESTADO DE FLUJO DE EFECTIVO', 3, 75)]:
        ws = model_wb[sheet_name]
        vert_start, hor_start = base_col + n_entries + 2, base_col + 2 * n_entries + 4
        ws.cell(row=2, column=vert_start, value="ANÁLISIS VERTICAL")
        ws.cell(row=2, column=hor_start, value="ANÁLISIS HORIZONTAL")
        
        for i, y in enumerate(years_sorted):
            ws.cell(row=3, column=vert_start+i, value=str(y)).number_format = 'General'
        for i, y in enumerate(years_sorted[1:]):
            ws.cell(row=3, column=hor_start+i, value=str(y)).number_format = 'General'
        
        for r in range(start_row+1, start_row+rows_count):
            for j in range(n_entries):
                col_idx = base_col + j
                if formula := vertical_formula(ws, r, col_idx, sheet_name):
                    ws.cell(row=r, column=vert_start+j, value=formula).style = "percentage_style"
                if j+1 < n_entries:
                    ws.cell(row=r, column=hor_start+j, 
                           value=f"=IFERROR(({ws.cell(row=r, column=col_idx).coordinate}-{ws.cell(row=r, column=base_col+j+1).coordinate})/{ws.cell(row=r, column=base_col+j+1).coordinate},0)"
                           ).style = "percentage_style"

    ws_ratios = model_wb['RATIOS']
    for merged in list(ws_ratios.merged_cells.ranges): ws_ratios.unmerge_cells(str(merged))
    
    for i, y in enumerate(years_sorted):
        ws_ratios.cell(row=3, column=5+i, value=y).number_format = '0'
    
    for idx in range(len(years_sorted)):
        col, dc = 5 + idx, get_column_letter(base_col + idx)
        pc = get_column_letter(base_col + idx + 1) if idx < len(years_sorted) - 1 else dc
        
        for row, formula, style in [
            (4, f"=IFERROR('ESTADO DE SITUACIÓN FINANCIERA'!{dc}20/'ESTADO DE SITUACIÓN FINANCIERA'!{dc}55,0)", "decimal_style"),
            (5, f"=IFERROR(('ESTADO DE SITUACIÓN FINANCIERA'!{dc}20-'ESTADO DE SITUACIÓN FINANCIERA'!{dc}13)/'ESTADO DE SITUACIÓN FINANCIERA'!{dc}55,0)", "decimal_style"),
            (6, f"=IFERROR('ESTADO DE SITUACIÓN FINANCIERA'!{dc}69/'ESTADO DE SITUACIÓN FINANCIERA'!{dc}40,0)", "decimal_style"),
            (7, f"=IFERROR('ESTADO DE SITUACIÓN FINANCIERA'!{dc}69/'ESTADO DE SITUACIÓN FINANCIERA'!{dc}78,0)", "decimal_style"),
            (8, f"=IFERROR('ESTADO DE RESULTADOS'!{dc}28/'ESTADO DE RESULTADOS'!{dc}4,0)", "percentage_style"),
            (9, f"=IFERROR('ESTADO DE RESULTADOS'!{dc}28/'ESTADO DE SITUACIÓN FINANCIERA'!{dc}40,0)", "percentage_style"),
            (10, f"=IFERROR('ESTADO DE RESULTADOS'!{dc}28/'ESTADO DE SITUACIÓN FINANCIERA'!{dc}78,0)", "percentage_style"),
            (11, f"=IFERROR('ESTADO DE RESULTADOS'!{dc}4/'ESTADO DE SITUACIÓN FINANCIERA'!{dc}40,0)", "decimal_style")
        ]:
            ws_ratios.cell(row=row, column=col, value=formula).style = style
        
        if idx < len(years_sorted) - 1:
            ws_ratios.cell(row=12, column=col, value=f"=IFERROR('ESTADO DE RESULTADOS'!{dc}4/((SUM('ESTADO DE SITUACIÓN FINANCIERA'!{dc}8:{dc}11)+SUM('ESTADO DE SITUACIÓN FINANCIERA'!{dc}24:{dc}27)+SUM('ESTADO DE SITUACIÓN FINANCIERA'!{pc}8:{pc}11)+SUM('ESTADO DE SITUACIÓN FINANCIERA'!{pc}24:{pc}27))/2),0)").style = "decimal_style"
            ws_ratios.cell(row=13, column=col, value=f"=IFERROR(-'ESTADO DE RESULTADOS'!{dc}5/((('ESTADO DE SITUACIÓN FINANCIERA'!{dc}13+'ESTADO DE SITUACIÓN FINANCIERA'!{pc}13)+('ESTADO DE SITUACIÓN FINANCIERA'!{dc}30+'ESTADO DE SITUACIÓN FINANCIERA'!{pc}30))/2),0)").style = "decimal_style"
        else:
            for row in [12, 13]: ws_ratios.cell(row=row, column=col, value=0).style = "decimal_style"

    out_file = os.path.join(output_dir, "REPORTE_ANALISIS_FINANCIERO.xlsx")
    model_wb.save(out_file)
    model_wb.close()
    return out_file

def safe_float(x):
    try: return float(x) if x is not None else 0.0
    except: return 0.0

def add_value_labels(ax, rects, values, ttype):
    y_min, y_max = ax.get_ylim()
    for rect, v in zip(rects, values):
        sign = 1 if v >= 0 else -1
        ax.text(rect.get_x() + rect.get_width()/2, v + (y_max - y_min) * 0.02 * sign, 
                f"{v*100:.2f}%" if ttype == 'pct' else f"{v:.2f}",
                ha='center', va='bottom' if sign >= 0 else 'top', fontsize=9, fontweight='bold')

def nice_ticks(ax, values, ttype, n_ticks=6):
    v_max, v_min = max(values + [0]), min(values + [0])
    if (span := v_max - v_min) == 0:
        v_max, v_min = v_max + abs(v_max)*0.1 + 0.1, v_min - abs(v_min)*0.1 - 0.1
        span = v_max - v_min
    
    ax.set_ylim(v_min - span*0.05, v_max + span*0.15)
    ax.set_yticks(np.linspace(v_min - span*0.05, v_max + span*0.15, n_ticks))
    ax.yaxis.set_major_formatter(PercentFormatter(1.0) if ttype == 'pct' else plt.FuncFormatter(lambda x, pos: f"{x:.2f}"))

def generate_ratios_charts_pdf(report_path, output_dir):
    plt.rcParams.update({'font.sans-serif': 'Arial', 'font.size': 10})
    
    wb = openpyxl.load_workbook(report_path, data_only=True)
    ws_ratios, ws_bs, ws_is = wb['RATIOS'], wb['ESTADO DE SITUACIÓN FINANCIERA'], wb['ESTADO DE RESULTADOS']

    years, col = [], 5
    while (v := ws_ratios.cell(row=3, column=col).value):
        years.append(int(v))
        col += 1
    if not years:
        wb.close()
        raise ValueError("No se encontraron años en la hoja RATIOS.")

    ratios_info = [
        {'row': 4, 'name': 'Liquidez corriente', 'short': 'Liquidez corriente', 'type': 'ratio', 'category': 'Liquidez'},
        {'row': 5, 'name': 'Prueba ácida', 'short': 'Prueba ácida', 'type': 'ratio', 'category': 'Liquidez'},
        {'row': 6, 'name': 'Razón de deuda total', 'short': 'Razón deuda total', 'type': 'ratio', 'category': 'Endeudamiento'},
        {'row': 7, 'name': 'Razón deuda/patrimonio', 'short': 'Razón deuda/patrimonio', 'type': 'ratio', 'category': 'Endeudamiento'},
        {'row': 8, 'name': 'Margen neto', 'short': 'Margen neto', 'type': 'pct', 'category': 'Rentabilidad'},
        {'row': 9, 'name': 'ROA', 'short': 'ROA', 'type': 'pct', 'category': 'Rentabilidad'},
        {'row': 10, 'name': 'ROE', 'short': 'ROE', 'type': 'pct', 'category': 'Rentabilidad'},
        {'row': 11, 'name': 'Rotación de activos totales', 'short': 'Rotación activos', 'type': 'ratio', 'category': 'Actividad'},
        {'row': 12, 'name': 'Rotación de cuentas por cobrar', 'short': 'Rotación CxC', 'type': 'ratio', 'category': 'Actividad'},
        {'row': 13, 'name': 'Rotación de inventarios', 'short': 'Rotación invent.', 'type': 'ratio', 'category': 'Actividad'},
    ]

    def bs_val(row, col_idx): return safe_float(ws_bs.cell(row=row, column=col_idx).value)
    def is_val(row, col_idx): return safe_float(ws_is.cell(row=row, column=col_idx).value)

    all_ratios = {r['row']: [] for r in ratios_info}

    for i in range(len(years)):
        col_idx = 3 + i
        bs = {k: bs_val(k, col_idx) for k in [20, 55, 13, 30, 69, 40, 78]}
        is_ = {k: is_val(k, col_idx) for k in [28, 4, 5]}

        for row, calc in [
            (4, bs[20] / bs[55] if bs[55] else 0),
            (5, (bs[20] - bs[13]) / bs[55] if bs[55] else 0),
            (6, bs[69] / bs[40] if bs[40] else 0),
            (7, bs[69] / bs[78] if bs[78] else 0),
            (8, is_[28] / is_[4] if is_[4] else 0),
            (9, is_[28] / bs[40] if bs[40] else 0),
            (10, is_[28] / bs[78] if bs[78] else 0),
            (11, is_[4] / bs[40] if bs[40] else 0)
        ]:
            all_ratios[row].append(calc)

        if i < len(years) - 1:
            prev_col = col_idx + 1
            cxc_avg = (sum(bs_val(r, col_idx) for r in list(range(8, 12)) + list(range(24, 28))) + 
                      sum(bs_val(r, prev_col) for r in list(range(8, 12)) + list(range(24, 28)))) / 2
            all_ratios[12].append(is_[4] / cxc_avg if cxc_avg else 0)
            
            inv_avg = ((bs[13] + bs_val(13, prev_col)) + (bs[30] + bs_val(30, prev_col))) / 2
            all_ratios[13].append((-is_[5]) / inv_avg if inv_avg else 0)
        else:
            all_ratios[12].append(0)
            all_ratios[13].append(0)

    with PdfPages(os.path.join(output_dir, "GRAFICOS_RATIOS_FINANCIEROS.pdf")) as pdf:
        fig = plt.figure(figsize=(11.69, 8.27))
        for y_pos, text, size in [(0.7, "Análisis Detallado de Ratios Financieros", 24), 
                                  (0.5, f"Periodo: {min(years)} - {max(years)}", 18), 
                                  (0.3, "Generado por Analizador Financiero SMV", 12)]:
            fig.text(0.5, y_pos, text, ha='center', fontsize=size)
        pdf.savefig(fig)
        plt.close(fig)
        
        for group in [{'title': 'Análisis de Liquidez y Endeudamiento', 'categories': ['Liquidez', 'Endeudamiento'], 'layout': (2, 2)},
                     {'title': 'Análisis de Rentabilidad y Actividad', 'categories': ['Rentabilidad', 'Actividad'], 'layout': (2, 3)}]:
            group_ratios = [r for r in ratios_info if r['category'] in group['categories']]
            n_rows, n_cols = group['layout']
            fig, axes = plt.subplots(n_rows, n_cols, figsize=(11.69, 8.27), squeeze=False)
            fig.suptitle(group['title'], fontsize=18, y=0.98, fontweight='bold', color='darkslategrey')
            
            for i, r in enumerate(group_ratios):
                ax = axes[i // n_cols, i % n_cols]
                values = all_ratios[r['row']]
                rects = ax.bar(np.arange(len(years)), values, width=0.7, color='#1E90FF', edgecolor='black', linewidth=0.5, alpha=0.95)
                ax.set_xticks(np.arange(len(years)))
                ax.set_xticklabels([str(y) for y in years], fontsize=10, fontweight='medium')
                ax.set_title(r['name'], fontsize=12, fontweight='bold', color='darkslategrey')
                nice_ticks(ax, values, r['type'])
                ax.set_ylabel(r['short'], fontsize=10, color='grey')
                ax.grid(axis='y', linestyle='-', alpha=0.4)
                add_value_labels(ax, rects, values, r['type'])
            
            for i in range(len(group_ratios), n_rows * n_cols):
                fig.delaxes(axes[i // n_cols, i % n_cols])

            plt.tight_layout(rect=[0, 0.05, 1, 0.95])
            pdf.savefig(fig)
            plt.close(fig)

    wb.close()
    return os.path.join(output_dir, "GRAFICOS_RATIOS_FINANCIEROS.pdf")