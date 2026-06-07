# doc_extractor

A starter project for extracting engineering drawing information from PDF and scanned PDF drawings into Excel.

## Features
- Reads PDF text directly when available.
- Falls back to OCR for scanned pages when a Tesseract installation is available.
- Builds a simple cable schedule-style table from the drawing text based on the user's prompt.
- Exports the result to an Excel workbook.

## Quick start
1. cd c:\\Workspace\\doc_extractor
2. python -m pip install -r requirements.txt
3. python -m doc_extractor.main --input <drawing.pdf> --prompt "Prepare a cable schedule..." --output cable_schedule.xlsx

## Example prompt
Prepare a cable schedule from the MBSRE system drawing with the columns Sl.no, Cable tag, From Equipment, Location, To Equipment, Location, Cable Size, Cable Type, Remarks.
