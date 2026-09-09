# Contact List Cleaner

A browser-only contact cleaning app built with React and Vite. Contact files are processed on the device and are not uploaded to a server.

## Run

- `npm ci`
- `npm run dev`
- `npm test`
- `npm run lint`
- `npm run build`

## Workflow

1. Upload CSV, XLSX, or XLS (up to 25 MB).
2. Select the worksheet and header row, then confirm column mappings.
3. Choose contact requirements, country, name handling, multi-value selection, duplicate matching and additional columns.
4. Search and filter all results, compare original and export values, and restore removed rows as needed.
5. Export cleaned XLSX/CSV, removed rows with reasons, or a complete original-data audit.

Phone validation checks numbering rules, not reachability. Extensions are exported separately. Email validation checks basic syntax, not delivery. Duplicate matching keeps the most complete record (number of populated output fields) or the first record, with source order breaking ties. Different phone extensions remain separate. Shared phone numbers can be preserved by selecting email-only matching or disabling deduplication.

Extra columns retain their original text. Automatic name splitting can be disabled for business or complex names. Original values remain available in comparison and audit exports. Restoring a row overrides removal but does not restore invalid phone/email values into the cleaned output.

CSV exports prefix formula-like values (including + phones) with an apostrophe for spreadsheet safety. Use XLSX for text cells without this prefix.

## Deployment

Personal GitHub repository: `averyricco25/dataset-cleaner`.
The existing Vercel hosting project belongs to the work account. For automatic production deployments, connect that Vercel project to this GitHub repository and select `main` as the production branch. Build command: `npm run build`; output directory: `dist`.
