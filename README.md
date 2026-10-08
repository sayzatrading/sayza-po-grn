# PO + GRN Control Center — PO/GRN First Build

This build intentionally excludes the Invoice Generator. Invoice generation will be merged later after the PO, GRN and reconciliation workflow is finalized.

## Included
- Upload one PO or multiple POs in one upload.
- Upload one GRN or multiple GRNs in one upload.
- PDF/XLSX/XLSM parsing using the existing working PO and GRN parsers.
- Purchase Orders tab with:
  - One main search bar across PO fields.
  - Individual column filters for every PO column.
  - Select individual rows or all visible filtered rows.
  - Live totals for selected rows: Qty, Taxable Value, GST, Total Value.
  - Booked Appointment Date selector/save column.
  - Original / Excel / PDF download for each uploaded PO.
- GRNs tab with:
  - One main search bar across GRN fields.
  - Individual column filters for every GRN column.
  - Select individual rows or all visible filtered rows.
  - Live totals for selected rows: Expected Qty, Received Qty, Taxable Value, GST, Total Value.
  - Original / Excel / PDF download for each uploaded GRN.
- PO ↔ GRN reconciliation remains based on PO Number + item code/SKU and aggregates multiple GRNs for the same PO/item.
- Existing reports/export functionality retained.

## Local run
```text
py -m pip install -r requirements.txt
py server.py
```
Open `http://127.0.0.1:3002`.

Do not use Live Server. This application needs the Python/Flask server.

## Existing database
If you already have a working PO/GRN database from the previous PO + GRN Control Center, copy its `data\\control_center.db` into this build's `data` folder before starting.

## Invoice Generator
Not included intentionally. It will be developed and merged separately after this PO/GRN build is approved.

## V3 Excel-style filters

PO and GRN tables now use Excel-style filter dropdowns directly in every column header.
- Click the ▼ in any column header.
- Search values inside that column.
- Select multiple values with checkboxes.
- Select All / Clear.
- Apply multiple column filters together (AND between columns, OR within a column).
- The global search bar still searches across the row fields.
- Select filtered rows and the totals at the bottom update for the selected rows only.

## V7 update: PO ↔ GRN column filters

The PO ↔ GRN Reconciliation table now uses the same Excel-style per-column multi-select filters as the PO and GRN tabs. Every reconciliation column has a filter button, including PO Number, Item Code, quantities, taxable values, totals, GRN Number(s), appointment date, and Status. The filter popup supports searching values, Select All for the currently searched values, Clear, Apply, and Clear Filter. Filters combine with the global reconciliation search and status selector.
