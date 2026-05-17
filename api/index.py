from flask import Flask, render_template, request, send_file, jsonify
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os
import json
import io
import tempfile

app = Flask(__name__, template_folder='../templates', static_folder='../static')

IDR_COLUMNS = [
    'Item Group', 'Item Name', 'Item Parent', 'Conversion Factor',
    'Projection Units', 'Total KGs', 'Packaging Type', 'Packaging Method',
    'Item Type', 'Customer Group', 'Origin', 'Created By',
    'Month', 'Year', 'Last Modified Date', 'Last Modified Time',
    'BOM Item Type', 'Customer', 'Key Account Manager', 'Item Code'
]

DEPT_DATA = [
    ('Production', 80, 48),
    ('Airlines', 34, 14),
    ('DB Conversion', 48, ''),
    ('Conti', 12, ''),
    ('PFS', 16, 16),
    ('FW', 36, 12),
    ('DB Bulk', 35, ''),
    ('VA Bulk', 17, 12),
    ('Inventory', 15, 5),
    ('Dispatch', 20, 8),
    ('RTV', 10, ''),
    ('Quality', 12, 3),
    ('FFS', 10, 10),
    ('HK', 16, 2),
    ('R&M', 5, ''),
]

SHIFT_CAPACITY = {
    'Conti Mix': 6500,
    'FW': '=90*60*3*7',
    'Jar Sealing': 2500,
    'Jar Sealing - WAD Machine': 2400,
    'LD Pack of 2': 2500,
    'LD Pack of 4': '=1248/0.8',
    'Pace - FFS': '=80*60*7',
    'Perfect - FFS': '=45*60*6',
    'PFS': 24000,
    'Table': 2500,
    'Table - Airlines': 18000,
}


def process_files(this_month_file, lookup_file, cf_value, working_days):
    df_proj = pd.read_excel(this_month_file, sheet_name='Projection', engine='openpyxl')
    df_proj.columns = [c.strip() for c in df_proj.columns]

    lookup_file.seek(0)
    df_lookup = pd.read_excel(lookup_file, engine='openpyxl')
    df_lookup.columns = [c.strip() for c in df_lookup.columns]

    lookup = {}
    for _, row in df_lookup.iterrows():
        name = str(row.get('Item Name', '')).strip()
        if name and name != 'nan':
            if name not in lookup:
                lookup[name] = {
                    'Packaging Type': str(row.get('Packaging Type', '')).strip(),
                    'Packaging Method': str(row.get('Packaging Method', '')).strip()
                }

    idr_rows = []
    for _, row in df_proj.iterrows():
        item_name = str(row.get('Item Name', '')).strip()
        pkg = lookup.get(item_name, {'Packaging Type': '', 'Packaging Method': ''})
        idr_row = {
            'Item Group': row.get('Item Group', ''),
            'Item Name': row.get('Item Name', ''),
            'Item Parent': row.get('Item Parent', ''),
            'Conversion Factor': row.get('Conversion Factor', ''),
            'Projection Units': row.get('Projection Units', ''),
            'Total KGs': row.get('Total KGs', ''),
            'Packaging Type': pkg['Packaging Type'],
            'Packaging Method': pkg['Packaging Method'],
            'Item Type': row.get('Item Type', ''),
            'Customer Group': row.get('Customer Group', ''),
            'Origin': row.get('Origin', row.get('Warehouse', '')),
            'Created By': row.get('Created By', ''),
            'Month': row.get('Month', ''),
            'Year': row.get('Year', ''),
            'Last Modified Date': row.get('Last Modified Date', ''),
            'Last Modified Time': row.get('Last Modified Time', ''),
            'BOM Item Type': row.get('BOM Item Type', ''),
            'Customer': row.get('Customer', ''),
            'Key Account Manager': row.get('Key Account Manager', ''),
            'Item Code': row.get('Item Code', '')
        }
        idr_rows.append(idr_row)

    df_idr = pd.DataFrame(idr_rows, columns=IDR_COLUMNS)
    df_idr = df_idr[df_idr['Origin'] == 'Indore'].reset_index(drop=True)

    pivot_source = df_idr[df_idr['Packaging Method'].notna() & (df_idr['Packaging Method'] != '') & (df_idr['Packaging Method'] != 'nan')].groupby('Packaging Method').agg(
        **{'Projected Units': ('Projection Units', lambda x: pd.to_numeric(x, errors='coerce').sum()),
           'Projected Kg': ('Total KGs', lambda x: pd.to_numeric(x, errors='coerce').sum())}
    ).reset_index().rename(columns={'Packaging Method': 'Source'})

    fw_mask = df_idr['Packaging Method'] == 'FW'
    pivot_fw = df_idr[fw_mask].groupby('Item Name').agg(
        **{'Projected Units': ('Projection Units', lambda x: pd.to_numeric(x, errors='coerce').sum()),
           'Projected Kg': ('Total KGs', lambda x: pd.to_numeric(x, errors='coerce').sum())}
    ).reset_index()

    va_mask = df_idr['Item Type'] == 'Value Added'
    pivot_va = df_idr[va_mask].groupby('Item Parent').agg(
        **{'Projected Kg': ('Total KGs', lambda x: pd.to_numeric(x, errors='coerce').sum())}
    ).reset_index().rename(columns={'Item Parent': 'Item Name'})

    return df_proj, df_idr, pivot_source, pivot_fw, pivot_va


def build_output_excel(df_proj, df_idr, pivot_source, pivot_fw, pivot_va, cf_value, working_days):
    wb = openpyxl.Workbook()
    hf = Font(bold=True, color='FFFFFF', size=11)
    hfill = PatternFill('solid', fgColor='2F5496')
    ha = Alignment(horizontal='center', vertical='center', wrap_text=True)
    tb = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    nk = '#,##0.00'
    nu = '#,##0'
    bf = Font(bold=True)

    # --- Tab 1: Projection (raw input) ---
    ws_raw = wb.active
    ws_raw.title = 'Projection'
    proj_cols = list(df_proj.columns)
    for ci, cn in enumerate(proj_cols, 1):
        c = ws_raw.cell(row=1, column=ci, value=cn)
        c.font = hf; c.fill = hfill; c.alignment = ha; c.border = tb
    for ri, (_, row) in enumerate(df_proj.iterrows(), 2):
        for ci, cn in enumerate(proj_cols, 1):
            v = row[cn]
            if pd.isna(v): v = ''
            c = ws_raw.cell(row=ri, column=ci, value=v)
            c.border = tb
    for ci in range(1, len(proj_cols)+1):
        ws_raw.column_dimensions[get_column_letter(ci)].width = 18
    ws_raw.auto_filter.ref = f"A1:{get_column_letter(len(proj_cols))}{len(df_proj)+1}"
    ws_raw.freeze_panes = 'A2'

    # --- Tab 2: IDR Projection ---
    ws_idr = wb.create_sheet('IDR Projection')
    for ci, cn in enumerate(IDR_COLUMNS, 1):
        c = ws_idr.cell(row=1, column=ci, value=cn)
        c.font = hf; c.fill = hfill; c.alignment = ha; c.border = tb
    for ri, (_, row) in enumerate(df_idr.iterrows(), 2):
        for ci, cn in enumerate(IDR_COLUMNS, 1):
            v = row[cn]
            if pd.isna(v): v = ''
            c = ws_idr.cell(row=ri, column=ci, value=v)
            c.border = tb
            if cn in ('Total KGs', 'Conversion Factor'): c.number_format = nk
            elif cn == 'Projection Units': c.number_format = nu
    for ci in range(1, len(IDR_COLUMNS)+1):
        ws_idr.column_dimensions[get_column_letter(ci)].width = 18
    ws_idr.auto_filter.ref = f"A1:T{len(df_idr)+1}"
    ws_idr.freeze_panes = 'A2'

    # --- Tab 3: Pivot (exact layout matching input) ---
    ws_p = wb.create_sheet('Pivot')
    sfill = PatternFill('solid', fgColor='D6E4F0')
    yfill = PatternFill('solid', fgColor='FFFF00')
    lbl_font = Font(bold=True, size=11, color='2F5496')

    # === SECTION 1: Source Summary (cols A-H) ===
    # Row 2: CF and Days inputs
    ws_p.cell(row=2, column=4, value=cf_value).font = bf
    ws_p.cell(row=2, column=6, value=working_days).font = bf

    # Row 3: Headers
    s1_headers = ['Source', 'Projected Units', 'Projected Kg', '1.3 CF',
                  'Sum of Pro (1.3 CF) Kg', 'Per Day PCS Prod.',
                  'No. of PCS Prod. (per shift)', 'No of Cycles']
    for ci, h in enumerate(s1_headers, 1):
        c = ws_p.cell(row=3, column=ci, value=h)
        c.font = hf; c.fill = hfill; c.alignment = ha; c.border = tb

    s1_start = 4
    for ri, (_, row) in enumerate(pivot_source.iterrows()):
        r = s1_start + ri
        ws_p.cell(row=r, column=1, value=row['Source']).border = tb
        ws_p.cell(row=r, column=2, value=row['Projected Units']).border = tb
        ws_p.cell(row=r, column=2).number_format = nu
        kg_val = row['Projected Kg']
        if pd.notna(kg_val) and kg_val != 0:
            ws_p.cell(row=r, column=3, value=kg_val).border = tb
            ws_p.cell(row=r, column=3).number_format = nk
        ws_p.cell(row=r, column=4, value=f'=B{r}*$D$2').border = tb
        ws_p.cell(row=r, column=4).number_format = nk
        ws_p.cell(row=r, column=5, value=f'=C{r}*$D$2').border = tb
        ws_p.cell(row=r, column=5).number_format = nk
        ws_p.cell(row=r, column=6, value=f'=D{r}/$F$2').border = tb
        ws_p.cell(row=r, column=6).number_format = nk
        cap = SHIFT_CAPACITY.get(row['Source'], '')
        ws_p.cell(row=r, column=7, value=cap).border = tb
        ws_p.cell(row=r, column=8, value=f'=F{r}/G{r}').border = tb
        ws_p.cell(row=r, column=8).number_format = nk

    s1_end = s1_start + len(pivot_source) - 1
    gt_row = s1_end + 1
    ws_p.cell(row=gt_row, column=1, value='Grand Total').font = bf
    ws_p.cell(row=gt_row, column=1).border = tb
    for ci in range(2, 9):
        cl = get_column_letter(ci)
        if ci == 7:
            ws_p.cell(row=gt_row, column=ci, value=f'=E{gt_row}').border = tb
        else:
            ws_p.cell(row=gt_row, column=ci, value=f'=SUM({cl}{s1_start}:{cl}{s1_end})').border = tb
        ws_p.cell(row=gt_row, column=ci).font = bf
        ws_p.cell(row=gt_row, column=ci).number_format = nk

    # === SECTION 2: FW Packaging Method breakdown (cols K-O) ===
    ws_p.cell(row=1, column=11, value='Packaging Method').font = lbl_font
    ws_p.cell(row=1, column=12, value='FW').font = lbl_font
    ws_p.cell(row=1, column=12).fill = sfill

    ws_p.cell(row=2, column=14, value=cf_value).font = bf
    ws_p.cell(row=2, column=15, value=working_days).font = bf

    fw_headers = ['Item Name', 'Projected Units', 'Projected Kg', '1.3 CF', 'Per Day Production Kg']
    for ci, h in enumerate(fw_headers):
        c = ws_p.cell(row=3, column=11+ci, value=h)
        c.font = hf; c.fill = hfill; c.alignment = ha; c.border = tb

    # Also write sub-header row for filtered section (row 14 area)
    # But first write FW data starting from row 4 if there are items,
    # otherwise from row 15 area matching original layout

    # Check if pivot_fw is small enough to fit before grand total at row 8
    # In original, FW section has items from row 4-7, Grand Total row 8, then gap,
    # then Packaging Method filter label at row 12, headers at row 14, items from 15+
    # Let's replicate: FW items with Packaging Method=FW filter in rows 4-7 area (top items)
    # Then a separate multi-item filtered section below

    # Top FW section: aggregate all FW into one line since original shows per-item
    fw_data_start = 4
    for ri, (_, row) in enumerate(pivot_fw.iterrows()):
        r = fw_data_start + ri
        ws_p.cell(row=r, column=11, value=row['Item Name']).border = tb
        ws_p.cell(row=r, column=12, value=row['Projected Units']).border = tb
        ws_p.cell(row=r, column=12).number_format = nu
        ws_p.cell(row=r, column=13, value=row['Projected Kg']).border = tb
        ws_p.cell(row=r, column=13).number_format = nk
        ws_p.cell(row=r, column=14, value=f'=M{r}*$N$2').border = tb
        ws_p.cell(row=r, column=14).number_format = nk
        ws_p.cell(row=r, column=15, value=f'=N{r}/$O$2').border = tb
        ws_p.cell(row=r, column=15).number_format = nk

    fw_end = fw_data_start + len(pivot_fw) - 1
    fw_gt = fw_end + 1
    ws_p.cell(row=fw_gt, column=11, value='Grand Total').font = bf
    ws_p.cell(row=fw_gt, column=11).border = tb
    for ci in range(12, 16):
        cl = get_column_letter(ci)
        ws_p.cell(row=fw_gt, column=ci, value=f'=SUM({cl}{fw_data_start}:{cl}{fw_end})').border = tb
        ws_p.cell(row=fw_gt, column=ci).font = bf
        ws_p.cell(row=fw_gt, column=ci).number_format = nk

    # === SECTION 3: Value Added - Item Parent summary (cols R-Y) ===
    ws_p.cell(row=1, column=18, value='Item Type').font = lbl_font
    ws_p.cell(row=1, column=19, value='Value Added').font = lbl_font
    ws_p.cell(row=1, column=19).fill = sfill

    ws_p.cell(row=2, column=20, value=cf_value).font = bf
    ws_p.cell(row=2, column=21, value=working_days).font = bf

    va_headers = ['Item Name', 'Projected Kg', '1.3 CF', 'Per Day Kg Prod',
                  'Roasting Item (%)', 'Roasting Kg', 'Oven Cycle (SKU Level)', 'No of Oven Cycle']
    for ci, h in enumerate(va_headers):
        c = ws_p.cell(row=3, column=18+ci, value=h)
        c.font = hf; c.fill = hfill; c.alignment = ha; c.border = tb

    va_start = 4
    for ri, (_, row) in enumerate(pivot_va.iterrows()):
        r = va_start + ri
        ws_p.cell(row=r, column=18, value=row['Item Name']).border = tb
        ws_p.cell(row=r, column=19, value=row['Projected Kg']).border = tb
        ws_p.cell(row=r, column=19).number_format = nk
        ws_p.cell(row=r, column=20, value=f'=S{r}*$T$2').border = tb
        ws_p.cell(row=r, column=20).number_format = nk
        ws_p.cell(row=r, column=21, value=f'=T{r}/{working_days}').border = tb
        ws_p.cell(row=r, column=21).number_format = nk
        # Roasting %, Roasting Kg, Oven Cycle, No of Oven Cycle - left empty for user to fill
        for cc in range(22, 26):
            ws_p.cell(row=r, column=cc).border = tb

    va_end = va_start + len(pivot_va) - 1
    va_gt = va_end + 1
    ws_p.cell(row=va_gt, column=18, value='Grand Total').font = bf
    ws_p.cell(row=va_gt, column=18).border = tb
    for ci in [19, 20]:
        cl = get_column_letter(ci)
        ws_p.cell(row=va_gt, column=ci, value=f'=SUM({cl}{va_start}:{cl}{va_end})').border = tb
        ws_p.cell(row=va_gt, column=ci).font = bf
        ws_p.cell(row=va_gt, column=ci).number_format = nk
    for ci in [23, 24, 25]:
        cl = get_column_letter(ci)
        ws_p.cell(row=va_gt, column=ci, value=f'=SUM({cl}{va_start}:{cl}{va_end})').border = tb
        ws_p.cell(row=va_gt, column=ci).font = bf
        ws_p.cell(row=va_gt, column=ci).number_format = nk

    # === SECTION 4: Department table (cols A-C, below source summary) ===
    dept_start = gt_row + 3
    dh_fill = PatternFill('solid', fgColor='2F5496')
    ws_p.cell(row=dept_start, column=1, value='Department').font = hf
    ws_p.cell(row=dept_start, column=1).fill = dh_fill
    ws_p.cell(row=dept_start, column=1).border = tb
    ws_p.cell(row=dept_start, column=2, value='Day').font = hf
    ws_p.cell(row=dept_start, column=2).fill = dh_fill
    ws_p.cell(row=dept_start, column=2).border = tb
    c = ws_p.cell(row=dept_start, column=3, value='Night')
    c.font = hf; c.fill = yfill; c.border = tb

    for di, (dept, day, night) in enumerate(DEPT_DATA):
        r = dept_start + 1 + di
        ws_p.cell(row=r, column=1, value=dept).border = tb
        ws_p.cell(row=r, column=2, value=day).border = tb
        c = ws_p.cell(row=r, column=3, value=night if night != '' else None)
        c.border = tb
        if night != '':
            c.fill = yfill

    total_r = dept_start + 1 + len(DEPT_DATA)
    ws_p.cell(row=total_r, column=1, value='Total').font = bf
    ws_p.cell(row=total_r, column=1).border = tb
    ws_p.cell(row=total_r, column=2, value=f'=SUM(B{dept_start+1}:B{total_r-1})').font = bf
    ws_p.cell(row=total_r, column=2).border = tb
    ws_p.cell(row=total_r, column=3, value=f'=SUM(C{dept_start+1}:C{total_r-1})').font = bf
    ws_p.cell(row=total_r, column=3).border = tb

    # Column widths
    for ci in range(1, 26):
        ws_p.column_dimensions[get_column_letter(ci)].width = 22

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/process', methods=['POST'])
def process():
    if 'this_month' not in request.files or 'lookup_file' not in request.files:
        return jsonify({'error': 'Both files are required'}), 400

    this_month = request.files['this_month']
    lookup_f = request.files['lookup_file']
    cf_value = float(request.form.get('cf_value', 1.3))
    working_days = int(request.form.get('working_days', 25))

    this_buf = io.BytesIO(this_month.read())
    lookup_buf = io.BytesIO(lookup_f.read())

    try:
        df_proj, df_idr, pivot_source, pivot_fw, pivot_va = process_files(this_buf, lookup_buf, cf_value, working_days)
    except Exception as e:
        return jsonify({'error': f'Processing error: {str(e)}'}), 500

    output = build_output_excel(df_proj, df_idr, pivot_source, pivot_fw, pivot_va, cf_value, working_days)

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
    tmp.write(output.read())
    tmp.close()
    app.config['LAST_OUTPUT'] = tmp.name

    summary = {
        'total_items': len(df_idr),
        'total_raw_items': len(df_proj),
        'total_projection_units': int(pd.to_numeric(df_idr['Projection Units'], errors='coerce').sum()),
        'total_kgs': round(float(pd.to_numeric(df_idr['Total KGs'], errors='coerce').sum()), 2),
        'unique_items': int(df_idr['Item Name'].nunique()),
        'unique_customers': int(df_idr['Customer'].nunique()),
        'cf_value': cf_value,
        'working_days': working_days,
        'item_groups': df_idr['Item Group'].value_counts().head(15).to_dict(),
        'customer_groups': df_idr['Customer Group'].value_counts().head(10).to_dict(),
        'packaging_methods': df_idr['Packaging Method'].value_counts().head(10).to_dict(),
        'packaging_types': df_idr['Packaging Type'].value_counts().head(10).to_dict(),
        'item_types': df_idr['Item Type'].value_counts().to_dict(),
        'origins': df_idr['Origin'].value_counts().head(10).to_dict(),
        'top_items_by_kg': df_idr.groupby('Item Name')['Total KGs'].apply(
            lambda x: pd.to_numeric(x, errors='coerce').sum()
        ).nlargest(10).to_dict(),
        'kam_summary': df_idr.groupby('Key Account Manager')['Total KGs'].apply(
            lambda x: pd.to_numeric(x, errors='coerce').sum()
        ).nlargest(10).to_dict(),
        'pivot_source': pivot_source.to_dict(orient='records'),
        'pivot_fw': pivot_fw.head(30).to_dict(orient='records'),
        'pivot_va': pivot_va.head(30).to_dict(orient='records'),
        'idr_preview': json.loads(df_idr.head(50).to_json(orient='records', default_handler=str)),
        'pkg_match_rate': round(
            (df_idr['Packaging Type'].notna() & (df_idr['Packaging Type'] != '') & (df_idr['Packaging Type'] != 'nan')).sum() / len(df_idr) * 100, 1
        ) if len(df_idr) > 0 else 0,
        'dept_data': [{'Department': d, 'Day': dy, 'Night': n if n != '' else ''} for d, dy, n in DEPT_DATA],
    }

    return jsonify(summary)


@app.route('/download')
def download():
    output_path = app.config.get('LAST_OUTPUT')
    if output_path and os.path.exists(output_path):
        return send_file(output_path, as_attachment=True, download_name='IDR_Projection_Output.xlsx',
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    return jsonify({'error': 'No output file found. Process files first.'}), 404


if __name__ == '__main__':
    app.run(debug=True, port=5000)
