# Local Document Batch Processor

A browser-based JavaScript portfolio demo for high-volume document handling.

The application processes **file names and metadata locally in the browser**. It does not upload documents anywhere.

## Business problem

Teams often receive hundreds or thousands of files per month. Common manual steps include:

- checking whether names follow a standard
- identifying duplicates
- checking whether document pairs are complete
- preparing a consistent rename list
- creating an exception report

## Demo solution

The app accepts local files or built-in sample data and:

- parses a simple document naming convention
- validates document type, reference and date
- standardises the proposed file name
- flags duplicate documents
- flags missing INV / PO pairs
- exports a CSV rename and exception report

## Demo naming convention

`TYPE_REFERENCE_YYYY-MM-DD.ext`

Supported demo types:

- `INV` — invoice
- `PO` — purchase order
- `DN` — delivery note

Example:

`INV_ORD-1001_2026-10-01.pdf`

## Try it

Open `index.html` in a browser and click **Load sample data**.

You can also choose your own local files. Only the file names are read by the demo.

## Privacy

This demo intentionally uses browser-side processing. No external API, AI service or server is required.

All built-in data is synthetic and created for this portfolio.
