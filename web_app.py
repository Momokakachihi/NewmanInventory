import json
import os
import tempfile
from datetime import datetime
from pathlib import Path

import pandas as pd
from flask import Flask, jsonify, redirect, render_template, request, session, url_for

from index import (
    clean_value,
    connect_database,
    find_by_scan,
    find_column,
    setup_database,
)


app = Flask(__name__)
app.secret_key = os.environ.get("NEWMAN_SECRET_KEY", "newman-development-secret")
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024
UPLOADS = Path(tempfile.gettempdir()) / "newman_inventory_uploads"
UPLOADS.mkdir(exist_ok=True)


@app.before_request
def require_login():
    if request.endpoint in {"login", "logout", "static"}:
        return None
    if session.get("authenticated"):
        return None
    if request.path.startswith("/api/"):
        return jsonify(error="Please sign in to continue."), 401
    return redirect(url_for("login", next=request.full_path))


@app.route("/login", methods=["GET", "POST"])
def login():
    error = None
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        if username == "admin" and password == "1234":
            session.clear()
            session["authenticated"] = True
            next_url = request.form.get("next", "")
            if not next_url.startswith("/") or next_url.startswith("//"):
                next_url = url_for("home")
            return redirect(next_url)
        error = "Invalid username or password."
    return render_template("login.html", error=error, next=request.args.get("next", ""))


@app.get("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


def clean_excel_value(value):
    if value is None or pd.isna(value):
        return ""
    return clean_value(value)


def read_inventory_sheet(path, sheet_name):
    engine = "pyxlsb" if path.suffix.lower() == ".xlsb" else "openpyxl"
    raw = pd.read_excel(path, sheet_name=sheet_name, engine=engine, header=None)
    known_headers = {
        "branch name", "destination", "dr remarks", "dr number", "dr date",
        "area", "group model", "color", "engine number", "frame number",
        "warehouse location", "date forwarded",
    }
    best_row = None
    best_score = 0
    for row_number in range(min(len(raw.index), 50)):
        values = {" ".join(str(value).strip().lower().split()) for value in raw.iloc[row_number]}
        score = len(values & known_headers)
        if score > best_score:
            best_row = row_number
            best_score = score
    if best_row is None or best_score == 0:
        raise ValueError("Could not find the motorcycle column headers in this worksheet.")

    headers = []
    used_headers = set()
    for column_number, value in enumerate(raw.iloc[best_row]):
        header = clean_excel_value(value) or f"Source Column {column_number + 1}"
        original_header = header
        suffix = 2
        while header in used_headers:
            header = f"{original_header} {suffix}"
            suffix += 1
        used_headers.add(header)
        headers.append(header)

    dataframe = raw.iloc[best_row + 1:].copy()
    dataframe.columns = headers
    return dataframe.dropna(how="all"), best_row + 1


def import_workbook(connection, path, sheet_name):
    dataframe, header_row = read_inventory_sheet(path, sheet_name)
    imported = 0
    updated = 0
    now = datetime.now().isoformat(timespec="seconds")

    for row_number, values in enumerate(dataframe.to_dict(orient="records"), start=2):
        row = {str(key): clean_excel_value(value) for key, value in values.items()}
        engine_number = find_column(row, "Engine Number")
        frame_number = find_column(row, "Frame Number")
        dr_number = find_column(row, "DR Number")
        identifier = frame_number or engine_number or dr_number
        if not identifier:
            continue

        branch_name = find_column(row, "Branch Name")
        group_model = find_column(row, "Group Model")
        warehouse_location = find_column(row, "Warehouse Location")
        fields = {
            "vin": identifier.upper(),
            "brand": branch_name,
            "model": group_model,
            "color": find_column(row, "Color"),
            "status": "IN_WAREHOUSE",
            "location": warehouse_location,
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
            "sts_number": find_column(row, "STS Number"),
            "source_row": json.dumps(row, ensure_ascii=True),
        }

        existing = connection.execute(
            "SELECT id FROM motorcycles WHERE vin = ?", (fields["vin"],)
        ).fetchone()
        if existing:
            motorcycle_id = existing[0]
            update_fields = {
                key: value for key, value in fields.items()
                if key != "vin" and (key != "sts_number" or value)
            }
            set_clause = ", ".join(f"{key} = ?" for key in update_fields)
            connection.execute(
                f"UPDATE motorcycles SET {set_clause} WHERE id = ?",
                list(update_fields.values()) + [motorcycle_id],
            )
            updated += 1
        else:
            columns = ", ".join(fields)
            placeholders = ", ".join("?" for _ in fields)
            cursor = connection.execute(
                f"INSERT INTO motorcycles ({columns}) VALUES ({placeholders})",
                list(fields.values()),
            )
            connection.execute(
                """
                INSERT INTO movements
                    (motorcycle_id, movement_type, to_location, movement_date, notes)
                VALUES (?, 'IMPORTED', ?, ?, ?)
                """,
                (cursor.lastrowid, fields["location"], now, f"Imported row {row_number + header_row}"),
            )
            motorcycle_id = cursor.lastrowid
            imported += 1

        no_file_rows = connection.execute(
            """
            SELECT id, truck_number, destination, scanned_at
            FROM no_file_motorcycles
            WHERE UPPER(COALESCE(frame_number, '')) IN (?, ?)
               OR UPPER(COALESCE(engine_number, '')) IN (?, ?)
            """,
            (
                fields["frame_number"].upper(),
                fields["engine_number"].upper(),
                fields["frame_number"].upper(),
                fields["engine_number"].upper(),
            ),
        ).fetchall()
        for no_file_id, truck_number, destination, scanned_at in no_file_rows:
            connection.execute(
                """
                INSERT INTO movements
                    (motorcycle_id, movement_type, from_location, to_location, movement_date, notes)
                VALUES (?, 'NO_FILE_RECONCILED', ?, ?, ?, ?)
                """,
                (
                    motorcycle_id,
                    destination,
                    destination,
                    now,
                    f"Previously unloaded without a file from {truck_number or 'unknown truck'} "
                    f"at {scanned_at}",
                ),
            )
            connection.execute("DELETE FROM no_file_motorcycles WHERE id = ?", (no_file_id,))

    connection.commit()
    if imported == 0 and updated == 0:
        raise ValueError(
            "The worksheet was read, but no rows contained an Engine Number, "
            "Frame Number, or DR Number. Choose the detailed motorcycle worksheet."
        )
    return {"imported": imported, "updated": updated, "rows": len(dataframe.index)}


def get_connection():
    connection = connect_database()
    setup_database(connection)
    return connection


@app.get("/")
def home():
    return render_template("index.html")


@app.get("/inventory")
def inventory():
    return render_template("inventory.html")


@app.get("/import")
def import_inventory():
    return render_template("import.html")


@app.get("/report")
def report():
    return render_template("report.html")


@app.get("/delivery-report")
def delivery_report():
    return render_template("delivery_report.html")


@app.get("/sts")
def sts():
    return render_template("sts.html")


@app.get("/storage")
def storage():
    return render_template("storage.html")


@app.get("/api/storage")
def storage_data():
    connection = get_connection()
    rows = connection.execute(
        """
                SELECT vin, engine_number, frame_number, group_model, color,
                             status, warehouse_location
        FROM motorcycles
        WHERE status = 'IN_WAREHOUSE'
                    AND (
                            UPPER(warehouse_location) LIKE 'NMT1%'
                            OR UPPER(warehouse_location) LIKE 'NMT2%'
                            OR UPPER(warehouse_location) LIKE 'NMT3%'
                    )
        ORDER BY UPPER(warehouse_location), id
        """
    ).fetchall()
    connection.close()
    return jsonify(
        motorcycles=[
            {
                "identifier": row[0],
                "engine": row[1],
                "frame": row[2],
                "model": row[3],
                "color": row[4],
                "status": row[5],
                "location": row[6],
            }
            for row in rows
        ]
    )


@app.get("/api/reports/trucks")
def truck_reports():
    connection = get_connection()
    rows = connection.execute(
        """
        SELECT inventory_scans.truck_number, inventory_scans.scanned_at,
               motorcycles.vin, motorcycles.engine_number, motorcycles.frame_number,
               motorcycles.dr_number, motorcycles.sts_number, motorcycles.dr_date,
               motorcycles.group_model, inventory_scans.completed_at,
               dispatched.movement_date
        FROM inventory_scans
        JOIN motorcycles ON motorcycles.id = inventory_scans.motorcycle_id
        LEFT JOIN (
            SELECT motorcycle_id, MAX(movement_date) AS movement_date
            FROM movements
            WHERE movement_type = 'DISPATCHED'
            GROUP BY motorcycle_id
        ) AS dispatched ON dispatched.motorcycle_id = motorcycles.id
        ORDER BY inventory_scans.truck_number, inventory_scans.scanned_at
        """
    ).fetchall()
    connection.close()
    return jsonify(
        records=[
            {
                "truck_number": row[0],
                "scanned_at": row[1],
                "identifier": row[2],
                "engine": row[3],
                "frame": row[4],
                "dr_number": row[5],
                "sts_number": row[6],
                "dr_date": row[7],
                "model": row[8],
                "completed_at": row[9],
                "dispatch_date": row[10],
            }
            for row in rows
        ]
    )


@app.get("/api/delivery-report")
def delivery_report_data():
    connection = get_connection()
    rows = connection.execute(
        """
        SELECT motorcycles.dr_number, motorcycles.vin, motorcycles.engine_number,
               motorcycles.frame_number, motorcycles.group_model,
               motorcycles.destination, dispatched.movement_date
        FROM motorcycles
        LEFT JOIN (
            SELECT motorcycle_id, MAX(movement_date) AS movement_date
            FROM movements
            WHERE movement_type = 'DISPATCHED'
            GROUP BY motorcycle_id
        ) AS dispatched ON dispatched.motorcycle_id = motorcycles.id
        WHERE TRIM(COALESCE(motorcycles.dr_number, '')) <> ''
        ORDER BY motorcycles.dr_number, motorcycles.id
        """
    ).fetchall()
    connection.close()
    return jsonify(
        records=[
            {
                "dr_number": row[0],
                "identifier": row[1],
                "engine": row[2],
                "frame": row[3],
                "model": row[4],
                "destination": row[5],
                "dispatch_date": row[6],
            }
            for row in rows
        ]
    )


@app.get("/api/sts")
def sts_data():
    connection = get_connection()
    rows = connection.execute(
        """
        SELECT id, vin, engine_number, frame_number, group_model, dr_number,
               status, location
        FROM motorcycles
        ORDER BY CASE WHEN TRIM(COALESCE(dr_number, '')) = '' THEN 0 ELSE 1 END, id
        """
    ).fetchall()
    connection.close()
    return jsonify(
        records=[
            {
                "id": row[0],
                "identifier": row[1],
                "engine": row[2],
                "frame": row[3],
                "model": row[4],
                "dr_number": row[5] or "",
                "status": row[6],
                "location": row[7] or "",
            }
            for row in rows
        ]
    )


@app.patch("/api/sts/<int:motorcycle_id>")
def update_sts(motorcycle_id):
    dr_number = request.form.get("dr_number", "").strip()
    connection = get_connection()
    cursor = connection.execute(
        "UPDATE motorcycles SET dr_number = ? WHERE id = ?",
        (dr_number, motorcycle_id),
    )
    connection.commit()
    connection.close()
    if cursor.rowcount == 0:
        return jsonify(error="Motorcycle was not found."), 404
    return jsonify(message="DR number updated.", dr_number=dr_number)


@app.post("/api/upload")
def upload_workbook():
    uploaded = request.files.get("file")
    if not uploaded or not uploaded.filename:
        return jsonify(error="Choose an Excel file first."), 400
    suffix = Path(uploaded.filename).suffix.lower()
    if suffix not in {".xlsb", ".xlsx"}:
        return jsonify(error="Only .xlsb and .xlsx files are supported."), 400

    token = next(tempfile._get_candidate_names()) + suffix
    path = UPLOADS / token
    uploaded.save(path)
    try:
        engine = "pyxlsb" if suffix == ".xlsb" else "openpyxl"
        sheets = pd.ExcelFile(path, engine=engine).sheet_names
    except Exception as error:
        path.unlink(missing_ok=True)
        return jsonify(error=f"Could not read workbook: {error}"), 400
    return jsonify(token=token, sheets=sheets, filename=uploaded.filename)


@app.post("/api/import")
def import_uploaded_sheet():
    token = request.form.get("token", "")
    sheet = request.form.get("sheet", "")
    path = UPLOADS / Path(token).name
    if not path.is_file() or path.suffix.lower() not in {".xlsb", ".xlsx"}:
        return jsonify(error="Upload session expired. Please choose the file again."), 400
    try:
        connection = get_connection()
        result = import_workbook(connection, path, sheet)
        connection.close()
        path.unlink(missing_ok=True)
        return jsonify(result=result)
    except Exception as error:
        return jsonify(error=f"Import failed: {error}"), 400


@app.get("/api/stats")
def stats():
    warehouse = request.args.get("warehouse", "ALL").strip().upper()
    connection = get_connection()
    location_filter = "" if warehouse == "ALL" else "AND UPPER({}) LIKE ?"
    location_param = () if warehouse == "ALL" else (f"{warehouse}%",)
    stock_count = connection.execute(
        f"SELECT COUNT(*) FROM motorcycles WHERE status = 'IN_WAREHOUSE' {location_filter.format('location')}",
        location_param,
    ).fetchone()[0]
    departure_rows = connection.execute(
        f"""
        SELECT CAST(strftime('%m', movement_date) AS INTEGER) AS month,
               COUNT(*) AS quantity
        FROM movements
        WHERE movement_type = 'DISPATCHED' {location_filter.format('to_location')}
        GROUP BY month
        ORDER BY month
        """,
        location_param,
    ).fetchall()
    result = {
        "total": connection.execute("SELECT COUNT(*) FROM motorcycles").fetchone()[0],
        "in_warehouse": stock_count,
        "dispatched": connection.execute(
            "SELECT COUNT(*) FROM motorcycles WHERE status = 'DISPATCHED'"
        ).fetchone()[0],
        "warehouse": warehouse,
        "departures": [
            {"month": month, "quantity": quantity}
            for month, quantity in departure_rows
        ],
    }
    connection.close()
    return jsonify(result)


@app.post("/api/scan")
def scan():
    value = request.form.get("value", "")
    connection = get_connection()
    motorcycle = find_by_scan(connection, value)
    if motorcycle is None:
        connection.close()
        return jsonify(error="Motorcycle was not found."), 404
    result = {
        "id": motorcycle[0],
        "identifier": motorcycle[1],
        "branch": motorcycle[2],
        "model": motorcycle[3],
        "color": motorcycle[4],
        "status": motorcycle[5],
        "location": motorcycle[6],
        "engine": motorcycle[7],
        "frame": motorcycle[8],
        "dr_number": motorcycle[9],
    }
    connection.close()
    return jsonify(result)


@app.get("/api/inventory-scans")
def inventory_scans():
    connection = get_connection()
    rows = connection.execute(
        """
        SELECT inventory_scans.id, motorcycles.vin, motorcycles.engine_number,
               motorcycles.frame_number, motorcycles.group_model,
               inventory_scans.truck_number, inventory_scans.scanned_at,
               motorcycles.status, motorcycles.location,
               inventory_scans.completed_at
        FROM inventory_scans
        JOIN motorcycles ON motorcycles.id = inventory_scans.motorcycle_id
         WHERE inventory_scans.completed_at IS NULL
        ORDER BY inventory_scans.id DESC
        """
    ).fetchall()
    connection.close()
    return jsonify(
        scans=[
            {
                "id": row[0],
                "identifier": row[1],
                "engine": row[2],
                "frame": row[3],
                "model": row[4],
                "truck_number": row[5],
                "scanned_at": row[6],
                "status": row[7],
                "location": row[8],
                "completed_at": row[9],
            }
            for row in rows
        ]
    )


@app.post("/api/inventory-scan")
def inventory_scan():
    value = request.form.get("value", "").strip()
    truck_number = request.form.get("truck_number", "").strip()
    if not value or not truck_number:
        return jsonify(error="Enter a scan and select a truck number."), 400

    connection = get_connection()
    motorcycle = find_by_scan(connection, value)
    if motorcycle is None:
        connection.close()
        return jsonify(error="Motorcycle was not found."), 404

    active_load = connection.execute(
        """
        SELECT truck_number FROM inventory_scans
        WHERE motorcycle_id = ? AND completed_at IS NULL
        ORDER BY id DESC
        LIMIT 1
        """,
        (motorcycle[0],),
    ).fetchone()
    if active_load:
        connection.close()
        if active_load[0] == truck_number:
            return jsonify(error=f"This motorcycle is already loaded on {truck_number}."), 409
        return jsonify(
            error=(
                f"This motorcycle is already loaded on {active_load[0]}. "
                "It must be received or dispatched before loading it onto another truck."
            )
        ), 409

    now = datetime.now().isoformat(timespec="seconds")
    connection.execute(
        """
        INSERT INTO inventory_scans
            (motorcycle_id, scan_mode, truck_number, scanned_at)
        VALUES (?, ?, ?, ?)
        """,
        (motorcycle[0], "bulk", truck_number, now),
    )
    connection.execute(
        "UPDATE motorcycles SET status = 'DISPATCHED' WHERE id = ?",
        (motorcycle[0],),
    )
    connection.commit()
    connection.close()
    return jsonify(
        message="Motorcycle added to the load list.",
        identifier=motorcycle[1],
        model=motorcycle[3],
        truck_number=truck_number,
    )


@app.post("/api/unload")
def unload():
    value = request.form.get("value", "").strip()
    destination = request.form.get("destination", "").strip().upper()
    if not value or destination not in {"NMT1", "NMT2", "NMT3"}:
        return jsonify(error="Enter a scan and choose a destination."), 400

    connection = get_connection()
    motorcycle = find_by_scan(connection, value)
    now = datetime.now().isoformat(timespec="seconds")
    if motorcycle is None:
        connection.execute(
            """
            INSERT INTO no_file_motorcycles
                (frame_number, engine_number, destination, scanned_at, notes)
            VALUES (?, ?, ?, ?, ?)
            """,
            (value.upper(), value.upper(), destination, now, "Unloaded without a matching file"),
        )
        connection.commit()
        connection.close()
        return jsonify(
            message="No file found. The scan was saved for later reconciliation.",
            identifier=value.upper(),
            no_file=True,
        )

    connection.execute(
        """
        UPDATE motorcycles
        SET status = 'IN_WAREHOUSE', location = ?, warehouse_location = ?
        WHERE id = ?
        """,
        (destination, destination, motorcycle[0]),
    )
    connection.execute(
        """
        INSERT INTO movements
            (motorcycle_id, movement_type, from_location, to_location, movement_date)
        VALUES (?, 'RECEIVED', ?, ?, ?)
        """,
        (motorcycle[0], motorcycle[6], destination, now),
    )
    connection.execute(
        """
        UPDATE inventory_scans SET completed_at = ?
        WHERE motorcycle_id = ? AND completed_at IS NULL
        """,
        (now, motorcycle[0]),
    )
    connection.commit()
    connection.close()
    return jsonify(
        message="Motorcycle unloaded and added to the warehouse.",
        identifier=motorcycle[1],
        no_file=False,
    )


@app.get("/api/no-file")
def no_file_data():
    connection = get_connection()
    rows = connection.execute(
        """
        SELECT id, frame_number, engine_number, truck_number, destination, scanned_at
        FROM no_file_motorcycles
        ORDER BY id DESC
        """
    ).fetchall()
    connection.close()
    return jsonify(
        records=[
            {
                "id": row[0],
                "frame": row[1],
                "engine": row[2],
                "truck_number": row[3] or "",
                "destination": row[4] or "",
                "scanned_at": row[5],
            }
            for row in rows
        ]
    )


@app.post("/api/inventory-scan/<int:scan_id>/movement")
def inventory_scan_movement(scan_id):
    action = request.form.get("action", "").strip().lower()
    destination = request.form.get("destination", "").strip().upper()
    if action not in {"receive", "dispatch"} or destination not in {"NMT1", "NMT2", "NMT3"}:
        return jsonify(error="Choose Receive or Dispatch and a location from NMT1, NMT2, or NMT3."), 400

    connection = get_connection()
    row = connection.execute(
        """
        SELECT inventory_scans.motorcycle_id, motorcycles.location
        FROM inventory_scans
        JOIN motorcycles ON motorcycles.id = inventory_scans.motorcycle_id
        WHERE inventory_scans.id = ?
        """,
        (scan_id,),
    ).fetchone()
    if row is None:
        connection.close()
        return jsonify(error="Loaded motorcycle was not found."), 404

    now = datetime.now().isoformat(timespec="seconds")
    new_status = "IN_WAREHOUSE" if action == "receive" else "DISPATCHED"
    movement_type = "RECEIVED" if action == "receive" else "DISPATCHED"
    connection.execute(
        "UPDATE motorcycles SET status = ?, location = ?, warehouse_location = ? WHERE id = ?",
        (new_status, destination, destination, row[0]),
    )
    connection.execute(
        """
        INSERT INTO movements
            (motorcycle_id, movement_type, from_location, to_location, movement_date)
        VALUES (?, ?, ?, ?, ?)
        """,
        (row[0], movement_type, row[1], destination, now),
    )
    connection.execute(
        "UPDATE inventory_scans SET completed_at = ? WHERE id = ?",
        (now, scan_id),
    )
    connection.commit()
    connection.close()
    return jsonify(status=new_status, location=destination)


@app.post("/api/move")
def move():
    value = request.form.get("value", "")
    action = request.form.get("action", "")
    destination = request.form.get("destination", "").strip()
    if action not in {"receive", "dispatch"} or destination.upper() not in {"NMT1", "NMT2", "NMT3"}:
        return jsonify(error="Choose an action and a location from NMT1, NMT2, or NMT3."), 400
    destination = destination.upper()

    connection = get_connection()
    motorcycle = find_by_scan(connection, value)
    if motorcycle is None:
        connection.close()
        return jsonify(error="Motorcycle was not found."), 404
    now = datetime.now().isoformat(timespec="seconds")
    new_status = "IN_WAREHOUSE" if action == "receive" else "DISPATCHED"
    movement_type = "RECEIVED" if action == "receive" else "DISPATCHED"
    connection.execute(
        "UPDATE motorcycles SET status = ?, location = ?, warehouse_location = ? WHERE id = ?",
        (new_status, destination, destination, motorcycle[0]),
    )
    connection.execute(
        """
        INSERT INTO movements
            (motorcycle_id, movement_type, from_location, to_location, movement_date)
        VALUES (?, ?, ?, ?, ?)
        """,
        (motorcycle[0], movement_type, motorcycle[6], destination, now),
    )
    connection.execute(
        """
        UPDATE inventory_scans
        SET completed_at = ?
        WHERE motorcycle_id = ? AND completed_at IS NULL
        """,
        (now, motorcycle[0]),
    )
    connection.commit()
    connection.close()
    return jsonify(status=new_status, location=destination)


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=False)
