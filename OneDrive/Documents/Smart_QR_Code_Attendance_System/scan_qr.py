import cv2
from pyzbar.pyzbar import decode
import openpyxl
import os
import sqlite3
from datetime import datetime

# File paths
excel_file = 'attendance.xlsx'
db_file = 'attendance.db'

# Initialize Excel
if not os.path.exists(excel_file):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = 'Attendance'
    sheet.append(["Order", "Student ID", "Name", "Timestamp"])
    workbook.save(excel_file)

# Initialize SQLite database
def initialize_db():
    connection = sqlite3.connect(db_file)
    cursor = connection.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS Attendance (
            order_id INTEGER PRIMARY KEY AUTOINCREMENT,
            student_id TEXT,
            name TEXT,
            timestamp TEXT,
            UNIQUE(student_id, timestamp)
        )
    ''')
    connection.commit()
    connection.close()

# Function to check and store data in SQLite database
def store_in_db(student_id, name):
    connection = sqlite3.connect(db_file)
    cursor = connection.cursor()
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # Check if the student has already marked attendance today
    cursor.execute('''
        SELECT * FROM Attendance
        WHERE student_id = ? AND DATE(timestamp) = DATE(?)
    ''', (student_id, timestamp))
    if cursor.fetchone():
        print(f"Attendance already marked for {name} today.")
        connection.close()
        return None

    # Insert record and fetch order_id
    cursor.execute('''
        INSERT INTO Attendance (student_id, name, timestamp)
        VALUES (?, ?, ?)
    ''', (student_id, name, timestamp))
    order_id = cursor.lastrowid
    connection.commit()
    connection.close()
    print(f"Attendance marked for {name}.")
    return order_id, timestamp

def mark_attendance(data):
    # Open the workbook and select the active sheet
    workbook = openpyxl.load_workbook(excel_file)
    sheet = workbook.active

    # Extract student info
    student_info = data.split(', ')
    student_id = student_info[0].split(': ')[1]
    name = student_info[1].split(': ')[1]

    # Check and store attendance in the database
    result = store_in_db(student_id, name)
    if not result:
        return
    order_id, timestamp = result

    # Append new data to Excel
    sheet.append([order_id, student_id, name, timestamp])
    workbook.save(excel_file)

def scan_qr_code():
    cap = cv2.VideoCapture(0)

    while True:
        _, frame = cap.read()

        # Flip the frame horizontally
        frame = cv2.flip(frame, 1)  # Change 1 to 0 for vertical flip, or -1 for both axes

        decoded_objects = decode(frame)
        for obj in decoded_objects:
            data = obj.data.decode('utf-8')
            mark_attendance(data)
            print(f"Scanned Data: {data}")

        cv2.imshow("QR Code Scanner", frame)

        if cv2.waitKey(1) & 0xFF == ord('q'):  # Press 'q' to quit
            break

    cap.release()
    cv2.destroyAllWindows()


if __name__ == "__main__":
    # Initialize database
    initialize_db()

    while True:
        print("\n--- Attendance System ---")
        print("1. Start QR Code Scanning")
        print("2. Exit")
        choice = input("Enter your choice: ")

        if choice == '1':
            scan_qr_code()
        elif choice == '2':
            print("Exiting...")
            break
        else:
            print("Invalid choice. Please try again.")
