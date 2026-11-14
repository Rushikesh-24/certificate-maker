# Certify

A static, browser-side certificate generator. Choose a PDF template, upload a CSV/XLS/XLSX spreadsheet, select the name column, style and position the name, choose whether to preserve names or convert them to title case, then download personalized PDFs in a ZIP file. The searchable font catalog contains 1,957 open-source Google Fonts.

## Run locally

Serve this folder with any static file server (for example, `python3 -m http.server 8000`). No build step or backend is needed.

The PDF and spreadsheet are processed locally in the browser. The app loads PDF.js, pdf-lib, SheetJS, and JSZip from cdnjs; Google Fonts are loaded from Google Fonts when selected. An internet connection is needed to load these libraries and fonts.

The first worksheet is used, and the spreadsheet's first row must contain column headers. The name is placed on the selected PDF page. Output ZIPs contain one full PDF copy per non-empty selected-column cell, named from that cell's value. Duplicate names receive a numbered suffix such as `Rushi Gaonkar (2).pdf`.
