# LIC Policy Tracker

A simple offline-first LIC policy tracker built with React, Vite, Tailwind CSS,
and SheetJS. Policy data is saved locally in the browser and can be exported
as an updated Excel file.

## Easiest option

Double-click the finished `LIC-Policy-Tracker.html` file. It is fully
self-contained and works without a server or internet connection.

## Work with the source code

On Windows, double-click `START_LOCALHOST.bat`. It starts the tracker at
`http://localhost:5173` and opens it in your browser.

Or start it manually:

1. Install Node.js.
2. Open a terminal in the source folder.
3. Run `npm install`.
4. Run `npm run dev`.
5. Open the local address shown in the terminal.

Do not double-click the source `index.html`; Vite source files must be run with
the command above.

## Publish with GitHub Pages

Run `npm run build`, then upload the generated `dist/index.html` to the root of
the GitHub repository and replace the source `index.html`. The build is a single
self-contained file, so GitHub Pages does not need an Actions workflow or an
assets folder.

## Excel columns

The first worksheet can use the LIC format shown below:

`SNo`, `Name`, `DOB`, `Policy`, `Agcode`, `Com.Date`, `P/T/PP`, `SumAssd`,
`Mode`, `Premium`

Only `.xlsx` files are accepted. You can select or drop multiple workbooks at
once, and later uploads are added to the records already saved locally. Duplicate
policy numbers are skipped automatically. Use **Export Updated Excel** to save
the combined data.
If the `Name` cell contains the name, phone number, and address on separate
lines, the tracker preserves and displays all of them. A missing `Status`
column defaults to **Unpaid**.

Policies, premium dues, critical unpaid records, and paid history are rendered
in pages of 40 records to keep the tracker responsive with large workbooks.

The Dashboard includes separate **Old Agency** and **New Agency** workspaces.
Their fixed, non-editable agency codes are `0035713F` for Old Agency and
`0212313F` for New Agency. New Excel uploads are automatically separated by the
`Agcode` column, and existing matching policies are moved to the correct agency
without removing unmatched records. Each agency keeps its own policies, dues,
paid history, uploads, and exports. Records saved by older tracker versions are
initially placed in Old Agency. The selected agency and separated records are
restored automatically whenever the same tracker file is reopened on this
device.
