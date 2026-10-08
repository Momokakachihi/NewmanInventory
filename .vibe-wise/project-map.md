# Project Map

## Purpose
NewmanInventory is a Flask web application for importing motorcycle inventory, viewing warehouse statistics, scanning motorcycles, and recording movements.

## Requirements
`web_app.py` defines the Flask app and HTTP routes. `index.py` owns SQLite setup, shared database helpers, scan lookup, and a command-line interface. `templates/index.html` renders the overview page; `templates/inventory.html` renders loading and unloading controls; `templates/import.html` handles workbook imports; `templates/report.html`, `templates/delivery_report.html`, and `templates/sts.html` provide reporting and DR assignment views; `templates/storage.html` shows warehouse locations. `static/images/` contains interface images.

## Components
Browser page -> Flask route in `web_app.py` -> shared database/helper functions in `index.py` or route-local SQL -> SQLite `inventory.db` -> JSON response -> browser JavaScript updates the page. Workbook imports additionally pass through the temporary upload directory and pandas/openpyxl or pyxlsb. Unknown unloading scans go to `no_file_motorcycles` and are reconciled by later imports.

## Main Flow
SQLite stores `motorcycles`, `movements`, `inventory_scans`, and `no_file_motorcycles`. Flask receives browser requests and uploaded files. Uploaded workbooks are saved temporarily before being read and removed after import. Authentication is not implemented in the inspected routes.

## Data and Trust Boundaries
The Flask app is started from `web_app.py` using the project's Python environment; exact command and deployment configuration still need verification.

## Build and Deployment
Not mapped yet.

## Unknowns
Run/deployment commands and the remaining client-side flow need verification. The command-line flow in `index.py` and the web flow in `web_app.py` share database logic but are separate entry points.
