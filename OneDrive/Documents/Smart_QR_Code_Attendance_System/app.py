from flask import Flask, render_template, request, redirect, url_for, session, flash, send_file
from werkzeug.security import generate_password_hash, check_password_hash
import qrcode
import sqlite3
from datetime import datetime
import openpyxl
import os
from io import BytesIO

app = Flask(__name__)
app.secret_key = 'your_secret_key'

DB_FILE = "attendance.db"

# Initialize database
def initialize_db():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()

    # Users table
    cursor.execute('''CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT UNIQUE NOT NULL,
        password TEXT NOT NULL,
        role TEXT NOT NULL
    )''')

    # Attendance table
    cursor.execute('''CREATE TABLE IF NOT EXISTS attendance (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        student_id TEXT NOT NULL,
        name TEXT NOT NULL,
        timestamp TEXT NOT NULL
    )''')

    conn.commit()
    conn.close()

# Save attendance
def save_attendance(student_id, name):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    cursor.execute('''INSERT INTO attendance (student_id, name, timestamp) VALUES (?, ?, ?)''',
                   (student_id, name, timestamp))
    conn.commit()
    conn.close()

# Export attendance as Excel
def export_attendance():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute('SELECT * FROM attendance')
    records = cursor.fetchall()
    conn.close()

    # Create Excel file
    output = BytesIO()
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Attendance Records"

    # Add headers
    sheet.append(["ID", "Student ID", "Name", "Timestamp"])

    # Add data
    for row in records:
        sheet.append(row)

    workbook.save(output)
    output.seek(0)
    return output

@app.route('/')
def home():
    if 'user_id' in session:
        role = session['role']
        return render_template('index.html', role=role)
    return redirect(url_for('login'))

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password']
        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()
        cursor.execute('SELECT id, password, role FROM users WHERE username = ?', (username,))
        user = cursor.fetchone()
        conn.close()

        if user and check_password_hash(user[1], password):
            session['user_id'] = user[0]
            session['role'] = user[2]
            flash("Logged in successfully!", "success")
            return redirect(url_for('home'))
        else:
            flash("Invalid username or password", "danger")
    return render_template('login.html')

@app.route('/signup', methods=['GET', 'POST'])
def signup():
    if request.method == 'POST':
        username = request.form['username']
        password = generate_password_hash(request.form['password'])
        role = request.form['role']

        try:
            conn = sqlite3.connect(DB_FILE)
            cursor = conn.cursor()
            cursor.execute('INSERT INTO users (username, password, role) VALUES (?, ?, ?)',
                           (username, password, role))
            conn.commit()
            conn.close()
            flash("Account created successfully!", "success")
            return redirect(url_for('login'))
        except sqlite3.IntegrityError:
            flash("Username already exists", "danger")
    return render_template('signup.html')

@app.route('/generate_qr', methods=['GET', 'POST'])
def generate_qr():
    if 'user_id' not in session:
        return redirect(url_for('login'))

    if request.method == 'POST':
        student_id = request.form['student_id']
        name = request.form['name']
        data = f"ID: {student_id}, Name: {name}"

        qr = qrcode.make(data)
        qr.save(f"static/{student_id}.png")
        flash("QR Code generated successfully!", "success")
        return render_template('generate_qr.html', qr_path=f"static/{student_id}.png")
    return render_template('generate_qr.html')

@app.route('/download_qr/<filename>', methods=['GET'])
def download_qr(filename):
    if 'user_id' not in session:
        return redirect(url_for('login'))
    
    # Ensure the file exists in the static folder
    qr_path = os.path.join('static', filename)
    if os.path.exists(qr_path):
        return send_file(qr_path, as_attachment=True, download_name=filename)
    else:
        flash("QR code not found", "danger")
        return redirect(url_for('generate_qr'))

@app.route('/search_attendance', methods=['GET', 'POST'])
def search_attendance():
    if 'user_id' not in session:
        return redirect(url_for('login'))

    results = []
    if request.method == 'POST':
        student_id = request.form['student_id']

        # Query the database for attendance records matching the student ID
        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()
        cursor.execute('SELECT * FROM attendance WHERE student_id = ?', (student_id,))
        results = cursor.fetchall()
        conn.close()

        if not results:
            flash(f"No attendance records found for Student ID: {student_id}", "danger")
        else:
            flash(f"Found {len(results)} record(s) for Student ID: {student_id}", "success")

    return render_template('search_attendance.html', results=results)

@app.route('/records', methods=['GET', 'POST'])
def records():
    if 'user_id' not in session:
        return redirect(url_for('login'))

    sort_by = request.args.get('sort_by', 'timestamp')
    order = request.args.get('order', 'asc')
    date_filter = request.args.get('date', None)

    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()

    query = f"SELECT * FROM attendance WHERE 1"
    params = []

    if date_filter:
        query += " AND DATE(timestamp) = ?"
        params.append(date_filter)

    query += f" ORDER BY {sort_by} {order.upper()}"
    cursor.execute(query, params)
    data = cursor.fetchall()
    conn.close()

    next_order = 'desc' if order == 'asc' else 'asc'
    return render_template('records.html', data=data, sort_by=sort_by, order=order, next_order=next_order, date_filter=date_filter)

@app.route('/export', methods=['GET'])
def export():
    if 'user_id' not in session:
        return redirect(url_for('login'))

    output = export_attendance()
    return send_file(output, as_attachment=True, download_name="attendance.xlsx")

@app.route('/logout')
def logout():
    session.clear()
    flash("Logged out successfully", "success")
    return redirect(url_for('login'))

if __name__ == '__main__':
    initialize_db()
    app.run(debug=True)
