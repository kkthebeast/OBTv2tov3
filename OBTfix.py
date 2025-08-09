import pyodbc
import json
import shutil
import sys
import os
from tkinter import filedialog
import tkinter as tk

def update_slide_records(mdb_path):
    # Step 1: Copy the original MDB to a new file with _WSv3
    new_mdb_path = mdb_path.replace('.mdb', '_WSv3.mdb')
    shutil.copy(mdb_path, new_mdb_path)

    # Step 2: Set up the Access connection string
    conn_str = (
        r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
        f'DBQ={new_mdb_path};'
    )

    # Step 3: Connect to the copied database
    conn = pyodbc.connect(conn_str)
    cursor = conn.cursor()

    print(f"Connected to: {new_mdb_path}")

    # Step 4: Ensure JSON_FIELDS exists on SLIDE_RECORDS (add if missing)
    try:
        cursor.execute("SELECT JSON_FIELDS FROM SLIDE_RECORDS WHERE 1=0")
    except pyodbc.Error:
        print("Adding JSON_FIELDS column to SLIDE_RECORDS...")
        cursor.execute("ALTER TABLE SLIDE_RECORDS ADD COLUMN JSON_FIELDS TEXT")

        # Only initialize rows where SEQUENCE_NO is not empty AND some identifying fields exist
        cursor.execute("SELECT UID, SEQUENCE_NO, ACTUALWELL, DAY, MONTH, YEAR FROM SLIDE_RECORDS")
        rows = cursor.fetchall()
        for row in rows:
            if (
                row.SEQUENCE_NO and str(row.SEQUENCE_NO).strip() != ""
                and (row.ACTUALWELL or row.DAY or row.MONTH or row.YEAR)
            ):
                default_json = json.dumps({"offBtmTq": 0.000000, "formation": ""})
                cursor.execute("UPDATE SLIDE_RECORDS SET JSON_FIELDS = ? WHERE UID = ?", (default_json, row.UID))
        conn.commit()
        print("JSON_FIELDS column added and initialized (only for valid rows).")

    # Step 5: Build TQ_OFF lookup dictionary from DAILY_REPORTS
    tq_lookup = {}
    try:
        cursor.execute("""
            SELECT ACTUALWELL, DAY, MONTH, YEAR, TQ_OFF 
            FROM DAILY_REPORTS 
            WHERE TQ_OFF IS NOT NULL AND TRIM(TQ_OFF) <> ''
        """)
        for row in cursor.fetchall():
            try:
                key = (row.ACTUALWELL, row.DAY, row.MONTH, row.YEAR)
                tq_value = float(str(row.TQ_OFF).strip() or 0)
                if tq_value != 0:
                    tq_lookup[key] = tq_value
            except (ValueError, TypeError):
                print(f"Warning: Skipping invalid TQ_OFF value for {row.ACTUALWELL} {row.DAY}/{row.MONTH}/{row.YEAR}")
        print(f"Loaded {len(tq_lookup)} valid TQ_OFF records.")
    except pyodbc.Error as e:
        print(f"Error reading DAILY_REPORTS: {e}")
        raise

    # Step 6: Loop through SLIDE_RECORDS and update JSON_FIELDS with matched TQ_OFF
    try:
        cursor.execute("""
            SELECT UID, ACTUALWELL, DAY, MONTH, YEAR, JSON_FIELDS, SEQUENCE_NO
            FROM SLIDE_RECORDS
        """)
        rows = cursor.fetchall()
        update_count = 0
        skip_count = 0

        for row in rows:
            # Skip headers or placeholder rows
            if (
                not row.SEQUENCE_NO or str(row.SEQUENCE_NO).strip() == ""
                or not (row.ACTUALWELL or row.DAY or row.MONTH or row.YEAR)
            ):
                skip_count += 1
                continue

            key = (row.ACTUALWELL, row.DAY, row.MONTH, row.YEAR)
            tq_off = tq_lookup.get(key)

            try:
                json_data = json.loads(row.JSON_FIELDS) if row.JSON_FIELDS else {"offBtmTq": 0.000000, "formation": ""}
            except Exception:
                print(f"Failed to parse JSON_FIELDS for UID {row.UID}, using default.")
                json_data = {"offBtmTq": 0.000000, "formation": ""}

            if tq_off is not None:
                json_data["offBtmTq"] = round(float(tq_off), 6)

            new_json = json.dumps(json_data)
            cursor.execute("UPDATE SLIDE_RECORDS SET JSON_FIELDS = ? WHERE UID = ?", (new_json, row.UID))
            update_count += 1

        conn.commit()
        print(f"Updated {update_count} SLIDE_RECORDS entries.")
        print(f"Skipped {skip_count} rows with empty SEQUENCE_NO and no data.")

    except pyodbc.Error as e:
        print(f"Error updating SLIDE_RECORDS: {e}")
        raise
    finally:
        conn.close()
        print("Database connection closed.")
        print("✅ Done!")

def select_mdb_file():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title="Select MDB File",
        filetypes=[("Access Database Files", "*.mdb")],
        initialdir=os.getcwd()
    )
    return file_path

if __name__ == "__main__":
    try:
        if len(sys.argv) > 1:
            mdb_file = sys.argv[1]
        else:
            mdb_file = select_mdb_file()
            if not mdb_file:
                print("No file selected.")
                input("Press Enter to exit...")
                sys.exit(0)

        if not os.path.isfile(mdb_file) or not mdb_file.lower().endswith(".mdb"):
            print("Error: Please select a valid .mdb file.")
            input("Press Enter to exit...")
            sys.exit(0)

        update_slide_records(mdb_file)
        print("\n✅ Operation completed successfully!")
        input("Press Enter to exit...")

    except pyodbc.Error as e:
        print(f"Database error: {e}")
        input("Press Enter to exit...")
        sys.exit(1)
    except Exception as e:
        print(f"An error occurred: {e}")
        input("Press Enter to exit...")
        sys.exit(1)
