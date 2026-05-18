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

ROASTING_LOOKUP = {
    'Apple Pie Date Bites': (0.401, 460),
    'Berry Mix': (0, 0),
    'Classic Date Bites': (0.4555, 262),
    'Classic Date bites': (0.4555, 262),
    'Coffe Rush Date Bites': (0.37, 460),
    'Dark Choco Date Bites': (0.401, 460),
    'Fruit and Nut Mix': (0.20, 300),
    'Mexican Peri Peri Snack Mix': (0.54, 600),
    'Mexican Peri Peri Snack Mix (FFS)': (0.70, 600),
    'Nut Mix': (1.0, 250),
    'Nut Mix - Airlines': (1.0, 250),
    'Premium Panchmewa Superfood': (0.40, 480),
    'Premium Panchmewa Superfood (FFS)': (0.30, 600),
    'Roasted and Salted Almonds': (1.0, 250),
    'Roasted and Salted Cashew': (1.0, 250),
    'Roasted and Salted Cashew (W240)': (1.0, 250),
    'Satva Mix': (0.63, 300),
    'Seed Mix': (0.875, 600),
    'Smokey Almonds': (0.92, 250),
    'Smokey Almonds - Airlines': (1.0, 250),
    'Smokey BBQ Party Mix': (0.94, 600),
    'Sweet and Salty Mix': (0.20, 300),
    'Trail Mix': (0.59, 600),
    'Tropical Savoury Mix': (0.23, 600),
    'Cream N Onion Makhana': (1.0, 100),
    'Mango Flavoured Raisins': (1.0, 200),
    'Roasted Black Pepper Cashew': (1.0, 240),
    'Roasted Chatpata Cashew': (1.0, 240),
    'Roasted Thai Chilli Cashew': (1.0, 240),
    'Roasted & Salted Pistachios': (1.0, 250),
    'Jumbo Iranian Roasted & Salted Pistachios': (1.0, 250),
    'Prasadam Roasted Makhana': (1.0, 0),
    'Paan Mix': (0, 0),
}

FFS_METHODS = {'Pace - FFS', 'Perfect - FFS'}


def process_files(this_month_file, lookup_file, cf_value, working_days):
    xls = pd.ExcelFile(this_month_file, engine='openpyxl')
    proj_sheet = None
    for name in ['Projection', 'Query Report']:
        if name in xls.sheet_names:
            proj_sheet = name
            break
    if proj_sheet is None:
        proj_sheet = xls.sheet_names[0]
    df_proj = pd.read_excel(xls, sheet_name=proj_sheet)
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
        idr_rows.append({
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
        })

    df_idr = pd.DataFrame(idr_rows, columns=IDR_COLUMNS)
    df_idr = df_idr[df_idr['Origin'] == 'Indore'].reset_index(drop=True)
    df_idr['Projection Units'] = pd.to_numeric(df_idr['Projection Units'], errors='coerce').fillna(0)
    df_idr['Total KGs'] = pd.to_numeric(df_idr['Total KGs'], errors='coerce').fillna(0)

    valid_pkg = df_idr['Packaging Method'].notna() & (df_idr['Packaging Method'] != '') & (df_idr['Packaging Method'] != 'nan')
    pivot_source = df_idr[valid_pkg].groupby('Packaging Method').agg(
        **{'Projected Units': ('Projection Units', 'sum'),
           'Projected Kg': ('Total KGs', 'sum')}
    ).reset_index().rename(columns={'Packaging Method': 'Source'})

    fw_mask = df_idr['Packaging Method'] == 'FW'
    pivot_fw = df_idr[fw_mask].groupby('Item Name').agg(
        **{'Projected Units': ('Projection Units', 'sum'),
           'Projected Kg': ('Total KGs', 'sum')}
    ).reset_index()

    ffs_mask = df_idr['Packaging Method'].isin(FFS_METHODS)
    pivot_ffs = df_idr[ffs_mask].groupby('Item Name').agg(
        **{'Projected Units': ('Projection Units', 'sum'),
           'Projected Kg': ('Total KGs', 'sum')}
    ).reset_index()

    va_mask = df_idr['Item Type'] == 'Value Added'
    pivot_va = df_idr[va_mask].groupby('Item Parent').agg(
        **{'Projected Kg': ('Total KGs', 'sum')}
    ).reset_index().rename(columns={'Item Parent': 'Item Name'})

    return df_proj, df_idr, pivot_source, pivot_fw, pivot_ffs, pivot_va


def build_output_excel(df_proj, df_idr, pivot_source, pivot_fw, pivot_ffs, pivot_va, cf_value, working_days):
    wb = openpyxl.Workbook()
    hf = Font(bold=True, color='FFFFFF', size=11)
    hfill = PatternFill('solid', fgColor='2F5496')
    ha = Alignment(horizontal='center', vertical='center', wrap_text=True)
    tb = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    nk = '#,##0.00'
    nu = '#,##0'
    bf = Font(bold=True)
    sfill = PatternFill('solid', fgColor='D6E4F0')
    yfill = PatternFill('solid', fgColor='FFFF00')
    lbl_font = Font(bold=True, size=11, color='2F5496')

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
            ws_raw.cell(row=ri, column=ci, value=v).border = tb
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

    # --- Tab 3: Pivot ---
    ws_p = wb.create_sheet('Pivot')

    # ========== SECTION 1: Source Summary (A-H) ==========
    ws_p.cell(row=2, column=4, value=cf_value).font = bf
    ws_p.cell(row=2, column=6, value=working_days).font = bf

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
        kg = row['Projected Kg']
        if pd.notna(kg) and kg != 0:
            ws_p.cell(row=r, column=3, value=kg).border = tb
            ws_p.cell(row=r, column=3).number_format = nk
        else:
            ws_p.cell(row=r, column=3).border = tb
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
    gt1 = s1_end + 1
    ws_p.cell(row=gt1, column=1, value='Grand Total').font = bf
    ws_p.cell(row=gt1, column=1).border = tb
    for ci in range(2, 9):
        cl = get_column_letter(ci)
        if ci == 7:
            ws_p.cell(row=gt1, column=ci, value=f'=E{gt1}').border = tb
        else:
            ws_p.cell(row=gt1, column=ci, value=f'=SUM({cl}{s1_start}:{cl}{s1_end})').border = tb
        ws_p.cell(row=gt1, column=ci).font = bf
        ws_p.cell(row=gt1, column=ci).number_format = nk

    # ========== SECTION 2a: FW Packaging Method (K-O) ==========
    ws_p.cell(row=1, column=11, value='Packaging Method').font = lbl_font
    ws_p.cell(row=1, column=12, value='FW').font = lbl_font
    ws_p.cell(row=1, column=12).fill = sfill

    ws_p.cell(row=2, column=14, value=cf_value).font = bf
    ws_p.cell(row=2, column=15, value=working_days).font = bf

    fw_headers = ['Item Name', 'Projected Units', 'Projected Kg', '1.3 CF', 'Per Day Production Kg']
    for ci, h in enumerate(fw_headers):
        c = ws_p.cell(row=3, column=11+ci, value=h)
        c.font = hf; c.fill = hfill; c.alignment = ha; c.border = tb

    fw_start = 4
    for ri, (_, row) in enumerate(pivot_fw.iterrows()):
        r = fw_start + ri
        ws_p.cell(row=r, column=11, value=row['Item Name']).border = tb
        ws_p.cell(row=r, column=12, value=row['Projected Units']).border = tb
        ws_p.cell(row=r, column=12).number_format = nu
        ws_p.cell(row=r, column=13, value=row['Projected Kg']).border = tb
        ws_p.cell(row=r, column=13).number_format = nk
        ws_p.cell(row=r, column=14, value=f'=M{r}*$N$2').border = tb
        ws_p.cell(row=r, column=14).number_format = nk
        ws_p.cell(row=r, column=15, value=f'=N{r}/$O$2').border = tb
        ws_p.cell(row=r, column=15).number_format = nk

    fw_end = fw_start + max(len(pivot_fw) - 1, 0)
    fw_gt = fw_end + 1
    ws_p.cell(row=fw_gt, column=11, value='Grand Total').font = bf
    ws_p.cell(row=fw_gt, column=11).border = tb
    for ci in range(12, 16):
        cl = get_column_letter(ci)
        if len(pivot_fw) > 0:
            ws_p.cell(row=fw_gt, column=ci, value=f'=SUM({cl}{fw_start}:{cl}{fw_end})').border = tb
        else:
            ws_p.cell(row=fw_gt, column=ci, value=0).border = tb
        ws_p.cell(row=fw_gt, column=ci).font = bf
        ws_p.cell(row=fw_gt, column=ci).number_format = nk

    # ========== SECTION 2b: (Multiple Items) = Pace-FFS + Perfect-FFS (K-O) ==========
    ffs_label_row = fw_gt + 4
    ws_p.cell(row=ffs_label_row, column=11, value='Packaging Method').font = lbl_font
    ws_p.cell(row=ffs_label_row, column=12, value='(Multiple Items)').font = lbl_font
    ws_p.cell(row=ffs_label_row, column=12).fill = sfill

    ffs_param_row = ffs_label_row + 1
    ws_p.cell(row=ffs_param_row, column=14, value=cf_value).font = bf
    ws_p.cell(row=ffs_param_row, column=15, value=working_days).font = bf

    ffs_hdr_row = ffs_param_row + 1
    for ci, h in enumerate(fw_headers):
        c = ws_p.cell(row=ffs_hdr_row, column=11+ci, value=h)
        c.font = hf; c.fill = hfill; c.alignment = ha; c.border = tb

    ffs_start = ffs_hdr_row + 1
    n2_ref = f'$N${ffs_param_row}'
    o2_ref = f'$O${ffs_param_row}'
    for ri, (_, row) in enumerate(pivot_ffs.iterrows()):
        r = ffs_start + ri
        ws_p.cell(row=r, column=11, value=row['Item Name']).border = tb
        ws_p.cell(row=r, column=12, value=row['Projected Units']).border = tb
        ws_p.cell(row=r, column=12).number_format = nu
        ws_p.cell(row=r, column=13, value=row['Projected Kg']).border = tb
        ws_p.cell(row=r, column=13).number_format = nk
        ws_p.cell(row=r, column=14, value=f'=M{r}*{n2_ref}').border = tb
        ws_p.cell(row=r, column=14).number_format = nk
        ws_p.cell(row=r, column=15, value=f'=N{r}/{o2_ref}').border = tb
        ws_p.cell(row=r, column=15).number_format = nk

    ffs_end = ffs_start + max(len(pivot_ffs) - 1, 0)
    ffs_gt = ffs_end + 1
    ws_p.cell(row=ffs_gt, column=11, value='Grand Total').font = bf
    ws_p.cell(row=ffs_gt, column=11).border = tb
    for ci in range(12, 16):
        cl = get_column_letter(ci)
        if len(pivot_ffs) > 0:
            ws_p.cell(row=ffs_gt, column=ci, value=f'=SUM({cl}{ffs_start}:{cl}{ffs_end})').border = tb
        else:
            ws_p.cell(row=ffs_gt, column=ci, value=0).border = tb
        ws_p.cell(row=ffs_gt, column=ci).font = bf
        ws_p.cell(row=ffs_gt, column=ci).number_format = nk

    # ========== SECTION 3: Value Added (R-Y) ==========
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
        item = row['Item Name']
        ws_p.cell(row=r, column=18, value=item).border = tb
        ws_p.cell(row=r, column=19, value=row['Projected Kg']).border = tb
        ws_p.cell(row=r, column=19).number_format = nk
        ws_p.cell(row=r, column=20, value=f'=S{r}*$T$2').border = tb
        ws_p.cell(row=r, column=20).number_format = nk
        ws_p.cell(row=r, column=21, value=f'=T{r}/$U$2').border = tb
        ws_p.cell(row=r, column=21).number_format = nk

        roast_pct, oven_cap = ROASTING_LOOKUP.get(item, (0, 0))
        ws_p.cell(row=r, column=22, value=roast_pct).border = tb
        ws_p.cell(row=r, column=22).number_format = '0%'
        ws_p.cell(row=r, column=23, value=f'=U{r}*V{r}').border = tb
        ws_p.cell(row=r, column=23).number_format = nk
        ws_p.cell(row=r, column=24, value=oven_cap).border = tb
        ws_p.cell(row=r, column=24).number_format = nu
        if oven_cap and oven_cap > 0:
            ws_p.cell(row=r, column=25, value=f'=W{r}/X{r}').border = tb
        else:
            ws_p.cell(row=r, column=25, value=0).border = tb
        ws_p.cell(row=r, column=25).number_format = '0.00'

    va_end = va_start + max(len(pivot_va) - 1, 0)
    va_gt = va_end + 1
    ws_p.cell(row=va_gt, column=18, value='Grand Total').font = bf
    ws_p.cell(row=va_gt, column=18).border = tb
    for ci in [19, 20]:
        cl = get_column_letter(ci)
        ws_p.cell(row=va_gt, column=ci, value=f'=SUM({cl}{va_start}:{cl}{va_end})').border = tb
        ws_p.cell(row=va_gt, column=ci).font = bf
        ws_p.cell(row=va_gt, column=ci).number_format = nk
    ws_p.cell(row=va_gt, column=22).border = tb
    for ci in [23, 24, 25]:
        cl = get_column_letter(ci)
        ws_p.cell(row=va_gt, column=ci, value=f'=SUM({cl}{va_start}:{cl}{va_end})').border = tb
        ws_p.cell(row=va_gt, column=ci).font = bf
        ws_p.cell(row=va_gt, column=ci).number_format = nk

    # ========== SECTION 4: Department (A-C below source) ==========
    dept_start = gt1 + 3
    ws_p.cell(row=dept_start, column=1, value='Department').font = hf
    ws_p.cell(row=dept_start, column=1).fill = hfill; ws_p.cell(row=dept_start, column=1).border = tb
    ws_p.cell(row=dept_start, column=2, value='Day').font = hf
    ws_p.cell(row=dept_start, column=2).fill = hfill; ws_p.cell(row=dept_start, column=2).border = tb
    c = ws_p.cell(row=dept_start, column=3, value='Night')
    c.font = Font(bold=True); c.fill = yfill; c.border = tb

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
        df_proj, df_idr, pivot_source, pivot_fw, pivot_ffs, pivot_va = process_files(this_buf, lookup_buf, cf_value, working_days)
    except Exception as e:
        return jsonify({'error': f'Processing error: {str(e)}'}), 500

    output = build_output_excel(df_proj, df_idr, pivot_source, pivot_fw, pivot_ffs, pivot_va, cf_value, working_days)

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
    tmp.write(output.read())
    tmp.close()
    app.config['LAST_OUTPUT'] = tmp.name

    # Build VA data with roasting for dashboard
    va_dashboard = []
    for _, row in pivot_va.iterrows():
        item = row['Item Name']
        kg = row['Projected Kg']
        cf_kg = kg * cf_value
        per_day = cf_kg / working_days
        rpct, ocap = ROASTING_LOOKUP.get(item, (0, 0))
        rkg = per_day * rpct
        nocyc = rkg / ocap if ocap > 0 else 0
        va_dashboard.append({
            'Item Name': item, 'Projected Kg': round(kg, 2),
            '1.3 CF': round(cf_kg, 2), 'Per Day Kg Prod': round(per_day, 0),
            'Roasting Item (%)': f'{rpct:.0%}', 'Roasting Kg': round(rkg, 2),
            'Oven Cycle': ocap, 'No of Oven Cycle': round(nocyc, 2)
        })

    # Build FFS dashboard data
    ffs_dashboard = []
    for _, row in pivot_ffs.iterrows():
        cf_val = row['Projected Kg'] * cf_value
        ffs_dashboard.append({
            'Item Name': row['Item Name'],
            'Projected Units': int(row['Projected Units']),
            'Projected Kg': round(row['Projected Kg'], 2),
            '1.3 CF': round(cf_val, 2),
            'Per Day Production Kg': round(cf_val / working_days, 0)
        })

    summary = {
        'total_items': len(df_idr),
        'total_raw_items': len(df_proj),
        'total_projection_units': int(df_idr['Projection Units'].sum()),
        'total_kgs': round(float(df_idr['Total KGs'].sum()), 2),
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
        'top_items_by_kg': df_idr.groupby('Item Name')['Total KGs'].sum().nlargest(10).to_dict(),
        'kam_summary': df_idr.groupby('Key Account Manager')['Total KGs'].sum().nlargest(10).to_dict(),
        'pivot_source': pivot_source.to_dict(orient='records'),
        'pivot_fw': pivot_fw.to_dict(orient='records'),
        'pivot_ffs': ffs_dashboard,
        'pivot_va': va_dashboard,
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
