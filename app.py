import os
import re
import sqlite3
import openpyxl
from flask import Flask, request, jsonify, render_template, send_file
from excel_exporter import build_utility_workbook

app = Flask(__name__, template_folder='templates', static_folder='static')
DB_PATH = os.path.join(os.path.dirname(__file__), 'utility.db')

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/units', methods=['GET'])
def get_units():
    search = request.args.get('q', '').strip()
    tower = request.args.get('tower', '').strip()
    conn = get_db()
    cur = conn.cursor()
    
    query = "SELECT DISTINCT unit, tower, owner_name, kva FROM billing_data WHERE 1=1"
    params = []
    
    if tower:
        query += " AND tower = ?"
        params.append(tower)
    if search:
        query += " AND (unit LIKE ? OR owner_name LIKE ?)"
        params.extend([f"%{search}%", f"%{search}%"])
        
    query += " ORDER BY unit LIMIT 50"
    cur.execute(query, params)
    units = [dict(row) for row in cur.fetchall()]
    conn.close()
    return jsonify(units)

@app.route('/api/periods', methods=['GET'])
def get_periods():
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT DISTINCT period FROM billing_data ORDER BY period DESC")
    periods = [row['period'] for row in cur.fetchall()]
    conn.close()
    return jsonify(periods)

@app.route('/api/towers', methods=['GET'])
def get_towers():
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT DISTINCT tower FROM billing_data WHERE tower IS NOT NULL AND tower != '' ORDER BY tower")
    towers = [row['tower'] for row in cur.fetchall()]
    conn.close()
    return jsonify(towers)

def _fetch_unit_usage(unit, periods):
    conn = get_db()
    cur = conn.cursor()

    query = "SELECT * FROM billing_data WHERE unit = ?"
    params = [unit]

    if periods and len(periods) > 0:
        placeholders = ','.join(['?'] * len(periods))
        query += f" AND period IN ({placeholders})"
        params.extend(periods)

    query += " ORDER BY period ASC"
    cur.execute(query, params)
    rows = [dict(row) for row in cur.fetchall()]
    conn.close()
    return rows

@app.route('/api/usage', methods=['GET'])
def get_usage():
    unit = request.args.get('unit', '').strip()
    periods = request.args.getlist('periods')

    if not unit:
        return jsonify({'error': 'Unit is required'}), 400

    rows = _fetch_unit_usage(unit, periods)
    return jsonify(rows)

@app.route('/api/export-excel', methods=['GET'])
def export_excel():
    unit = request.args.get('unit', '').strip()
    periods = [p.strip() for p in request.args.getlist('periods') if p.strip()]

    if not unit:
        return jsonify({'error': 'Unit is required'}), 400

    rows = _fetch_unit_usage(unit, periods)
    if not rows:
        return jsonify({'error': 'Data tidak ditemukan untuk unit dan periode yang dipilih'}), 404

    output = build_utility_workbook(rows)

    safe_unit = re.sub(r'[^A-Za-z0-9_-]+', '_', unit).strip('_') or 'UNIT'
    row_periods = [r['period'] for r in rows if r.get('period')]
    if len(row_periods) == 1:
        period_slug = re.sub(r'[^A-Za-z0-9_-]+', '_', row_periods[0])
    elif len(row_periods) > 1:
        p_start = re.sub(r'[^A-Za-z0-9_-]+', '_', row_periods[0])
        p_end = re.sub(r'[^A-Za-z0-9_-]+', '_', row_periods[-1])
        period_slug = f"{p_start}_sd_{p_end}"
    else:
        period_slug = "all"

    filename = f"Laporan_Utilitas_{safe_unit}_{period_slug}.xlsx"

    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=filename,
    )

@app.route('/api/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
        
    file = request.files['file']
    period = request.form.get('period', '').strip()
    
    if not period:
        return jsonify({'error': 'Period is required (e.g. 2026-11)'}), 400
    if not file.filename.endswith('.xlsx'):
        return jsonify({'error': 'Only .xlsx files are supported'}), 400
        
    temp_path = os.path.join(os.path.dirname(__file__), 'temp_upload.xlsx')
    file.save(temp_path)
    
    try:
        wb = openpyxl.load_workbook(temp_path, data_only=True)
        sheet_name = 'ME4a' if 'ME4a' in wb.sheetnames else wb.sheetnames[0]
        ws = wb[sheet_name]
        
        rows_to_insert = []
        for r in range(7, ws.max_row + 1):
            unit = ws.cell(r, 3).value
            if not unit:
                continue
            row = (
                period,
                ws.cell(r, 1).value,
                ws.cell(r, 2).value,
                str(unit).strip(),
                ws.cell(r, 4).value,
                ws.cell(r, 5).value,
                ws.cell(r, 6).value,
                ws.cell(r, 7).value,
                ws.cell(r, 8).value,
                ws.cell(r, 9).value,
                ws.cell(r, 10).value,
                ws.cell(r, 11).value,
                ws.cell(r, 12).value,
                ws.cell(r, 13).value,
                ws.cell(r, 14).value,
                ws.cell(r, 15).value,
                ws.cell(r, 16).value,
                ws.cell(r, 17).value,
                ws.cell(r, 18).value,
                ws.cell(r, 19).value,
                ws.cell(r, 20).value,
                ws.cell(r, 21).value,
                ws.cell(r, 22).value,
                ws.cell(r, 23).value,
                ws.cell(r, 24).value,
                ws.cell(r, 25).value,
                ws.cell(r, 26).value
            )
            rows_to_insert.append(row)
            
        conn = get_db()
        cur = conn.cursor()
        cur.execute("DELETE FROM billing_data WHERE period = ?", (period,))
        cur.executemany('''
        INSERT INTO billing_data (
            period, no, tower, unit, owner_name, mcb, kva, note,
            elec_meter_start, elec_meter_end, elec_usage, kwh_min, jam_nyala, elec_subtotal, elec_total,
            water_meter_start, water_meter_end, water_usage, water_m1, water_m2, water_m3,
            water_cost_m1, water_cost_m2, water_cost_m3, water_fixed_cost, water_total, grand_total
        ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
        ''', rows_to_insert)
        conn.commit()
        conn.close()
        
        if os.path.exists(temp_path):
            os.remove(temp_path)
            
        return jsonify({'message': f'Berhasil mengimpor {len(rows_to_insert)} unit untuk periode {period}!'})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5005, debug=True)
