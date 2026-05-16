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


def process_files(this_month_file, last_month_file):
    df_proj = pd.read_excel(this_month_file, sheet_name='Projection', engine='openpyxl')
    df_proj.columns = [c.strip() for c in df_proj.columns]

    last_month_file.seek(0)
    last_month_sheets = pd.ExcelFile(last_month_file, engine='openpyxl').sheet_names
    idr_sheet = [s for s in last_month_sheets if 'IDR' in s][0]
    last_month_file.seek(0)
    df_last_idr = pd.read_excel(last_month_file, sheet_name=idr_sheet, engine='openpyxl')
    df_last_idr.columns = [c.strip() for c in df_last_idr.columns]

    lookup = {}
    for _, row in df_last_idr.iterrows():
        name = str(row.get('Item Name', '')).strip()
        if name and name != 'nan':
            if name not in lookup:
                lookup[name] = {
                    'Packaging Type': row.get('Packaging Type', ''),
                    'Packaging Method': row.get('Packaging Method', '')
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

    pivot_source = df_idr.groupby('Packaging Method').agg(
        **{'Projected Units': ('Projection Units', 'sum'),
           'Projected Kg': ('Total KGs', 'sum')}
    ).reset_index().rename(columns={'Packaging Method': 'Source'})
    pivot_source = pivot_source[pivot_source['Source'].notna() & (pivot_source['Source'] != '') & (pivot_source['Source'] != 'nan')]

    fw_mask = df_idr['Packaging Method'] == 'FW'
    pivot_fw = df_idr[fw_mask].groupby('Item Name').agg(
        **{'Projected Units': ('Projection Units', 'sum'),
           'Projected Kg': ('Total KGs', 'sum')}
    ).reset_index()

    va_mask = df_idr['Item Type'] == 'Value Added'
    pivot_va = df_idr[va_mask].groupby('Item Parent').agg(
        **{'Projected Kg': ('Total KGs', 'sum')}
    ).reset_index().rename(columns={'Item Parent': 'Item Name'})

    return df_idr, pivot_source, pivot_fw, pivot_va


def build_output_excel(df_idr, pivot_source, pivot_fw, pivot_va):
    wb = openpyxl.Workbook()

    header_font = Font(bold=True, color='FFFFFF', size=11)
    header_fill = PatternFill('solid', fgColor='2F5496')
    header_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    num_fmt_kg = '#,##0.00'
    num_fmt_units = '#,##0'

    ws_idr = wb.active
    ws_idr.title = 'IDR Projection'
    for c_idx, col_name in enumerate(IDR_COLUMNS, 1):
        cell = ws_idr.cell(row=1, column=c_idx, value=col_name)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = thin_border

    for r_idx, (_, row) in enumerate(df_idr.iterrows(), 2):
        for c_idx, col_name in enumerate(IDR_COLUMNS, 1):
            val = row[col_name]
            if pd.isna(val):
                val = ''
            cell = ws_idr.cell(row=r_idx, column=c_idx, value=val)
            cell.border = thin_border
            if col_name in ('Total KGs', 'Conversion Factor'):
                cell.number_format = num_fmt_kg
            elif col_name == 'Projection Units':
                cell.number_format = num_fmt_units

    for c in range(1, len(IDR_COLUMNS) + 1):
        ws_idr.column_dimensions[get_column_letter(c)].width = 18
    ws_idr.auto_filter.ref = f"A1:T{len(df_idr) + 1}"
    ws_idr.freeze_panes = 'A2'

    ws_pivot = wb.create_sheet('Pivot')
    sections = [
        ('Source Summary (by Packaging Method)', pivot_source, ['Source', 'Projected Units', 'Projected Kg']),
        ('FW Packaging - Item Breakdown', pivot_fw, ['Item Name', 'Projected Units', 'Projected Kg']),
        ('Value Added - Item Parent Summary', pivot_va, ['Item Name', 'Projected Kg']),
    ]

    section_fill = PatternFill('solid', fgColor='D6E4F0')
    section_font = Font(bold=True, size=12, color='2F5496')
    current_row = 1

    for title, df, cols in sections:
        cell = ws_pivot.cell(row=current_row, column=1, value=title)
        cell.font = section_font
        cell.fill = section_fill
        for merge_c in range(2, len(cols) + 1):
            ws_pivot.cell(row=current_row, column=merge_c).fill = section_fill
        ws_pivot.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=len(cols))
        current_row += 1

        for c_idx, col_name in enumerate(cols, 1):
            cell = ws_pivot.cell(row=current_row, column=c_idx, value=col_name)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_align
            cell.border = thin_border
        current_row += 1

        data_start = current_row
        for _, row in df.iterrows():
            for c_idx, col_name in enumerate(cols, 1):
                val = row[col_name]
                if pd.isna(val):
                    val = ''
                cell = ws_pivot.cell(row=current_row, column=c_idx, value=val)
                cell.border = thin_border
                if 'Kg' in col_name:
                    cell.number_format = num_fmt_kg
                elif 'Units' in col_name:
                    cell.number_format = num_fmt_units
            current_row += 1

        total_row = current_row
        ws_pivot.cell(row=total_row, column=1, value='Grand Total').font = Font(bold=True)
        ws_pivot.cell(row=total_row, column=1).border = thin_border
        for c_idx, col_name in enumerate(cols, 1):
            if c_idx > 1:
                col_letter = get_column_letter(c_idx)
                formula = f'=SUM({col_letter}{data_start}:{col_letter}{total_row - 1})'
                cell = ws_pivot.cell(row=total_row, column=c_idx, value=formula)
                cell.font = Font(bold=True)
                cell.border = thin_border
                if 'Kg' in col_name:
                    cell.number_format = num_fmt_kg
                else:
                    cell.number_format = num_fmt_units
        current_row += 2

    for c in range(1, 6):
        ws_pivot.column_dimensions[get_column_letter(c)].width = 25

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/process', methods=['POST'])
def process():
    if 'this_month' not in request.files or 'last_month' not in request.files:
        return jsonify({'error': 'Both files are required'}), 400

    this_month = request.files['this_month']
    last_month = request.files['last_month']

    this_buf = io.BytesIO(this_month.read())
    last_buf = io.BytesIO(last_month.read())

    try:
        df_idr, pivot_source, pivot_fw, pivot_va = process_files(this_buf, last_buf)
    except Exception as e:
        return jsonify({'error': f'Processing error: {str(e)}'}), 500

    output = build_output_excel(df_idr, pivot_source, pivot_fw, pivot_va)

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
    tmp.write(output.read())
    tmp.close()
    app.config['LAST_OUTPUT'] = tmp.name

    summary = {
        'total_items': len(df_idr),
        'total_projection_units': int(pd.to_numeric(df_idr['Projection Units'], errors='coerce').sum()),
        'total_kgs': round(float(pd.to_numeric(df_idr['Total KGs'], errors='coerce').sum()), 2),
        'unique_items': int(df_idr['Item Name'].nunique()),
        'unique_customers': int(df_idr['Customer'].nunique()),
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
        'pivot_fw': pivot_fw.head(20).to_dict(orient='records'),
        'pivot_va': pivot_va.head(20).to_dict(orient='records'),
        'idr_preview': json.loads(df_idr.head(50).to_json(orient='records', default_handler=str)),
        'pkg_match_rate': round(
            (df_idr['Packaging Type'].notna() & (df_idr['Packaging Type'] != '') & (df_idr['Packaging Type'] != 'nan')).sum() / len(df_idr) * 100, 1
        ),
    }

    return jsonify(summary)


@app.route('/download')
def download():
    output_path = app.config.get('LAST_OUTPUT')
    if output_path and os.path.exists(output_path):
        return send_file(output_path, as_attachment=True, download_name='IDR_Projection_Output.xlsx',
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    return jsonify({'error': 'No output file found. Process files first.'}), 404
