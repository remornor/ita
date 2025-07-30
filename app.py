from flask import Flask, render_template, request, redirect, flash, send_file, jsonify
import sqlite3
from datetime import datetime
import os
import csv
import io
import json
from openpyxl import Workbook

app = Flask(__name__)
app.secret_key = 'your-secret-key'
DB_PATH = 'complaints.db'
ADMIN_PASSWORD = '123456'  # แก้ไขตรงนี้หากต้องการเปลี่ยนรหัส

def init_db():
    if not os.path.exists(DB_PATH):
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS reports (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                id_number TEXT NOT NULL,
                full_name TEXT NOT NULL,
                email TEXT NOT NULL,
                phone TEXT,
                complaint TEXT NOT NULL,
                timestamp TEXT NOT NULL,
                status TEXT DEFAULT 'รอดำเนินการ',
                created TEXT NOT NULL
            )
        """)
        conn.commit()
        conn.close()

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/submit', methods=['POST'])
def submit():
    try:
        data = {
            'idNumber': request.form['idNumber'],
            'fullName': request.form['fullName'],
            'email': request.form['email'],
            'phone': request.form['phone'],
            'complaint': request.form['complaint']
        }

        now = datetime.now()
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO reports (id_number, full_name, email, phone, complaint, timestamp, created)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (
            data['idNumber'],
            data['fullName'],
            data['email'],
            data['phone'],
            data['complaint'],
            now.strftime('%d/%m/%Y %H:%M:%S'),
            now.isoformat()
        ))
        conn.commit()
        conn.close()
        flash("✅ ส่งข้อมูลสำเร็จ", 'success')
    except Exception as e:
        print(e)
        flash("❌ เกิดข้อผิดพลาด", 'danger')
    return redirect('/')

@app.route('/admin/export', methods=['GET', 'POST'])
def admin_export():
    if request.method == 'GET':
        return render_template('admin_export.html')

    # ตรวจสอบรหัสผ่าน
    password = request.form.get('password')
    if password != ADMIN_PASSWORD:
        flash("❌ รหัสผ่านไม่ถูกต้อง", 'danger')
        return redirect('/admin/export')

    format_type = request.form.get('format')
    from_date = request.form.get('from_date')
    to_date = request.form.get('to_date')

    query = "SELECT * FROM reports"
    params = []

    if from_date and to_date:
        query += " WHERE DATE(created) BETWEEN ? AND ?"
        params = [from_date, to_date]

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute(query, params)
    rows = cursor.fetchall()
    columns = [desc[0] for desc in cursor.description]
    conn.close()

    if not rows:
        flash("ℹ️ ไม่พบข้อมูลในช่วงวันที่ระบุ", 'danger')
        return redirect('/admin/export')

    # Export CSV
    if format_type == 'csv':
        output = io.StringIO()
        writer = csv.writer(output)
        writer.writerow(columns)
        writer.writerows(rows)
        output.seek(0)
        return send_file(io.BytesIO(output.getvalue().encode('utf-8-sig')),
                         mimetype='text/csv',
                         as_attachment=True,
                         download_name=f"complaints_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv")

    # Export JSON
    if format_type == 'json':
        data = [dict(zip(columns, row)) for row in rows]
        json_data = json.dumps(data, ensure_ascii=False, indent=2)
        return send_file(io.BytesIO(json_data.encode('utf-8')),
                         mimetype='application/json',
                         as_attachment=True,
                         download_name=f"complaints_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json")

    # Export Excel (.xlsx)
    if format_type == 'excel':
        wb = Workbook()
        ws = wb.active
        ws.title = "Complaints"
        ws.append(columns)
        for row in rows:
            ws.append(row)
        excel_io = io.BytesIO()
        wb.save(excel_io)
        excel_io.seek(0)
        return send_file(excel_io,
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                         as_attachment=True,
                         download_name=f"complaints_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")

    flash("❌ ไม่รองรับรูปแบบไฟล์นี้", 'danger')
    return redirect('/admin/export')

if __name__ == '__main__':
    init_db()
    app.run(debug=True)
