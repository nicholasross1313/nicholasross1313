# Multi-File Reconciliation

A public Excel/VBA portfolio demo that compares **multiple operational datasets**, not just two files, and surfaces only the records that need attention.

## Business problem

A common back-office process looks like this:

1. Sales produces an order report.
2. Warehouse produces a shipment report.
3. Finance produces a billing report.
4. Someone manually compares the files row by row.

That works at small scale, but it becomes slow and error-prone when the files contain hundreds or thousands of rows.

## Demo solution

The VBA module reconciles three sheets:

- **Orders**
- **Warehouse**
- **Billing**

It builds a union of all Order IDs and checks:

- whether the order exists in every source
- SKU consistency
- ordered, shipped and billed quantity
- expected value versus invoice value
- unexpected downstream records
- records requiring manual review

The result is written to a single **Reconciliation Results** sheet.

## Example statuses

- Matched
- Missing warehouse record
- Missing billing record
- Quantity mismatch
- SKU mismatch
- Amount mismatch
- Extra downstream record

## Files

- `src/MultiFileReconciliation.bas` — VBA source
- `sample-data/orders.csv` — synthetic order data
- `sample-data/warehouse.csv` — synthetic warehouse data
- `sample-data/billing.csv` — synthetic billing data
- `sample-data/expected-results.csv` — expected output for the sample

## Try it

1. Create a new Excel workbook.
2. Import the three CSV files into sheets named `Orders`, `Warehouse` and `Billing`.
3. Open the VBA editor with `Alt+F11`.
4. Import `MultiFileReconciliation.bas`.
5. Run `RunMultiFileReconciliation`.
6. Review the generated `Reconciliation Results` sheet.

## Design choices

The demo uses late-bound dictionaries, so it does not require a manual reference to Microsoft Scripting Runtime.

It intentionally keeps business rules visible and simple. In a real implementation, the validation rules would be agreed with the process owner first.

## Data safety

All data in this project is synthetic and created for demonstration purposes only. No employer, client or production data is included.
