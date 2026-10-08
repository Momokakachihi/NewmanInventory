# Learning Progress

No learning events recorded yet.

## Onboarding

- Learner is improving and extending the NewmanInventory application.
- Learner began working with the codebase about three days ago.
- Learner is a beginner in Python and Flask and has a database background.
- Learner's goals are to understand code flow, debug confidently, and add features with understanding.
- Learner understands at a high level that SQLite stores data, Flask serves HTML, and Python files provide application functionality.

## Inventory expansion design

- Confirmed navigation areas: Overview, Inventory, Storage, Import, Report, Delivery Report, and STS.
- Loading remains the existing truck scan flow; Inventory now lets the user choose Load truck or Unload truck.
- Unloading checks frame number or engine number against the master inventory. Unknown scans are stored in `no_file_motorcycles`.
- A later workbook import reconciles matching No File records into the master inventory while preserving a reconciliation movement history.
- Report groups scans by truck; the user selects a truck and the printable report uses the current print date.
- Delivery Report groups all motorcycles sharing a DR number and derives dispatch date from the `DISPATCHED` movement.
- STS has Without DR and Assigned DR sections with editable DR numbers.
- A dedicated Import page is the chosen location for workbook uploads.

## STS report redesign

- The report follows the supplied STS form layout for Newmann 478 Trucking.
- The user selects a loaded truck and manually enters driver, helper, destination, truck plate number, and date prepared.
- Rows group motorcycles by DR number, or by a stored STS number when available; quantity is calculated per group.
- STS/DR Date comes from the imported DR date. Dispatched Date comes from the DISPATCHED movement. ETD and ETA remain blank.
- A future separate STS workbook import is still pending because its exact headers and matching fields are not yet available.
