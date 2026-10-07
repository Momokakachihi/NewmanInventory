import json
import sqlite3
from datetime import datetime
from pathlib import Path


DATABASE_FILE = "inventory.db"
IMPORT_COLUMNS = (
	"branch_name",
	"destination",
	"dr_remarks",
	"dr_number",
	"dr_date",
	"area",
	"group_model",
	"color",
	"engine_number",
	"frame_number",
	"warehouse_location",
	"date_forwarded",
)


def connect_database():
	connection = sqlite3.connect(DATABASE_FILE)
	connection.execute("PRAGMA foreign_keys = ON")
	return connection


def setup_database(connection):
	connection.executescript("""
		CREATE TABLE IF NOT EXISTS motorcycles (
			id INTEGER PRIMARY KEY AUTOINCREMENT,
			vin TEXT NOT NULL UNIQUE,
			brand TEXT NOT NULL DEFAULT '',
			model TEXT NOT NULL DEFAULT '',
			color TEXT,
			status TEXT NOT NULL,
			location TEXT,
			created_at TEXT NOT NULL,
			branch_name TEXT,
			destination TEXT,
			dr_remarks TEXT,
			dr_number TEXT,
			dr_date TEXT,
			area TEXT,
			group_model TEXT,
			engine_number TEXT,
			frame_number TEXT,
			warehouse_location TEXT,
			date_forwarded TEXT,
			source_row TEXT
		);

		CREATE TABLE IF NOT EXISTS movements (
			id INTEGER PRIMARY KEY AUTOINCREMENT,
			motorcycle_id INTEGER NOT NULL,
			movement_type TEXT NOT NULL,
			from_location TEXT,
			to_location TEXT,
			movement_date TEXT NOT NULL,
			notes TEXT,
			FOREIGN KEY (motorcycle_id) REFERENCES motorcycles(id)
		);

		CREATE TABLE IF NOT EXISTS inventory_scans (
			id INTEGER PRIMARY KEY AUTOINCREMENT,
			motorcycle_id INTEGER NOT NULL,
			scan_mode TEXT NOT NULL CHECK (scan_mode IN ('single', 'bulk')),
			truck_number TEXT NOT NULL,
			scanned_at TEXT NOT NULL,
			completed_at TEXT,
			FOREIGN KEY (motorcycle_id) REFERENCES motorcycles(id)
		);
	""")
	for column in IMPORT_COLUMNS:
		try:
			connection.execute(f"ALTER TABLE motorcycles ADD COLUMN {column} TEXT")
		except sqlite3.OperationalError as error:
			if "duplicate column name" not in str(error).lower():
				raise
	try:
		connection.execute("ALTER TABLE inventory_scans ADD COLUMN completed_at TEXT")
	except sqlite3.OperationalError as error:
		if "duplicate column name" not in str(error).lower():
			raise
	connection.commit()


def clean_value(value):
	if value is None:
		return ""
	if hasattr(value, "isoformat"):
		return value.isoformat()
	return str(value).strip()


def normalize_header(header):
	return " ".join(str(header).strip().lower().split())


def find_column(row, *names):
	normalized = {normalize_header(key): key for key in row}
	for name in names:
		if normalize_header(name) in normalized:
			return clean_value(row[normalized[normalize_header(name)]])
	return ""


def import_excel(connection):
	print("\nImport Excel workbook")
	path = Path(input("Path to .xlsb or .xlsx file: ").strip().strip('"'))
	if not path.is_file():
		print("File not found.")
		return

	try:
		import pandas as pd
	except ImportError:
		print("Excel import needs pandas and pyxlsb.")
		print("Install them with: pip install pandas pyxlsb openpyxl")
		return

	try:
		engine = "pyxlsb" if path.suffix.lower() == ".xlsb" else "openpyxl"
		excel_file = pd.ExcelFile(path, engine=engine)
		print("Sheets: " + ", ".join(excel_file.sheet_names))
		sheet = input("Sheet name (press Enter for the first sheet): ").strip()
		sheet = sheet or excel_file.sheet_names[0]
		dataframe = pd.read_excel(excel_file, sheet_name=sheet)
	except Exception as error:
		print(f"Could not read workbook: {error}")
		return

	imported = 0
	updated = 0
	now = datetime.now().isoformat(timespec="seconds")
	for row_number, values in enumerate(dataframe.to_dict(orient="records"), start=2):
		row = {str(key): clean_value(value) for key, value in values.items()}
		engine_number = find_column(row, "Engine Number")
		frame_number = find_column(row, "Frame Number")
		dr_number = find_column(row, "DR Number")
		identifier = frame_number or engine_number or dr_number
		if not identifier:
			print(f"Skipped row {row_number}: no Engine Number, Frame Number, or DR Number.")
			continue

		branch_name = find_column(row, "Branch Name")
		group_model = find_column(row, "Group Model")
		warehouse_location = find_column(row, "Warehouse Location")
		status = "IN_WAREHOUSE"
		location = warehouse_location
		fields = {
			"vin": identifier.upper(),
			"brand": branch_name,
			"model": group_model,
			"color": find_column(row, "Color"),
			"status": status,
			"location": location,
			"created_at": now,
			"branch_name": branch_name,
			"destination": find_column(row, "Destination"),
			"dr_remarks": find_column(row, "DR Remarks"),
			"dr_number": dr_number,
			"dr_date": find_column(row, "DR Date"),
			"area": find_column(row, "Area"),
			"group_model": group_model,
			"engine_number": engine_number.upper(),
			"frame_number": frame_number.upper(),
			"warehouse_location": warehouse_location,
			"date_forwarded": find_column(row, "DATE FORWARDED"),
			"source_row": json.dumps(row, ensure_ascii=True),
		}

		existing = connection.execute(
			"SELECT id FROM motorcycles WHERE vin = ?", (fields["vin"],)
		).fetchone()
		if existing:
			set_clause = ", ".join(f"{key} = ?" for key in fields if key != "vin")
			connection.execute(
				f"UPDATE motorcycles SET {set_clause} WHERE id = ?",
				[fields[key] for key in fields if key != "vin"] + [existing[0]],
			)
			updated += 1
		else:
			columns = ", ".join(fields)
			placeholders = ", ".join("?" for _ in fields)
			cursor = connection.execute(
				f"INSERT INTO motorcycles ({columns}) VALUES ({placeholders})",
				list(fields.values()),
			)
			connection.execute("""
				INSERT INTO movements
					(motorcycle_id, movement_type, to_location, movement_date, notes)
				VALUES (?, 'IMPORTED', ?, ?, ?)
			""", (cursor.lastrowid, location, now, f"Imported row {row_number}"))
			imported += 1

	connection.commit()
	print(f"Import complete: {imported} new rows, {updated} existing rows updated.")


def find_by_scan(connection, scan_value):
	scan_value = scan_value.strip().upper()
	return connection.execute("""
		SELECT id, vin, branch_name, group_model, color, status, location,
			engine_number, frame_number, dr_number
		FROM motorcycles
		WHERE UPPER(vin) = ? OR UPPER(engine_number) = ? OR UPPER(frame_number) = ?
	""", (scan_value, scan_value, scan_value)).fetchone()


def scan_motorcycle(connection):
	print("\nScan motorcycle")
	print("Scan the engine number or frame number. Press Enter when the scanner finishes.")
	motorcycle = find_by_scan(connection, input("Scan: "))
	if motorcycle is None:
		print("Motorcycle not found in the imported inventory.")
		return

	print(f"Found: {motorcycle[2]} | {motorcycle[3]} | {motorcycle[4]}")
	print(f"Status: {motorcycle[5]} | Location: {motorcycle[6] or '-'}")
	print(f"DR Number: {motorcycle[9] or '-'}")
	action = input("Action (dispatch/receive/cancel): ").strip().lower()
	if action not in {"dispatch", "receive"}:
		return

	new_location = input("New location or destination: ").strip()
	if not new_location:
		print("A location is required.")
		return

	now = datetime.now().isoformat(timespec="seconds")
	movement_type = "DISPATCHED" if action == "dispatch" else "RECEIVED"
	new_status = "DISPATCHED" if action == "dispatch" else "IN_WAREHOUSE"
	connection.execute(
		"UPDATE motorcycles SET status = ?, location = ?, warehouse_location = ? WHERE id = ?",
		(new_status, new_location, new_location, motorcycle[0]),
	)
	connection.execute("""
		INSERT INTO movements
			(motorcycle_id, movement_type, from_location, to_location, movement_date)
		VALUES (?, ?, ?, ?, ?)
	""", (motorcycle[0], movement_type, motorcycle[6], new_location, now))
	connection.commit()
	print(f"Motorcycle marked as {new_status}.")


def list_motorcycles(connection):
	motorcycles = connection.execute("""
		SELECT vin, engine_number, frame_number, group_model, status, location
		FROM motorcycles ORDER BY id
	""").fetchall()
	if not motorcycles:
		print("No motorcycles have been recorded.")
		return

	print("\nIdentifier | Engine | Frame | Model | Status | Location")
	print("-" * 100)
	for motorcycle in motorcycles:
		print(" | ".join(value or "-" for value in motorcycle))


def show_history(connection):
	motorcycle = find_by_scan(connection, input("Scan VIN, engine, or frame number: "))
	if motorcycle is None:
		print("Motorcycle not found.")
		return
	movements = connection.execute("""
		SELECT movement_type, from_location, to_location, movement_date
		FROM movements WHERE motorcycle_id = ? ORDER BY movement_date
	""", (motorcycle[0],)).fetchall()
	for movement in movements:
		print(f"{movement[3]} | {movement[0]} | {movement[1] or '-'} -> {movement[2] or '-'}")


def main():
	connection = connect_database()
	setup_database(connection)
	try:
		while True:
			print("""
Motorcycle Inventory
1. Import .xlsb/.xlsx rows
2. Scan motorcycle and record movement
3. List motorcycles
4. View movement history by scan
5. Exit
""")
			choice = input("Choose an option: ").strip()
			if choice == "1":
				import_excel(connection)
			elif choice == "2":
				scan_motorcycle(connection)
			elif choice == "3":
				list_motorcycles(connection)
			elif choice == "4":
				show_history(connection)
			elif choice == "5":
				print("Goodbye.")
				break
			else:
				print("Please choose a number from 1 to 5.")
	finally:
		connection.close()


if __name__ == "__main__":
	main()