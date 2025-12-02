# Excel Notes Splitter

A JavaScript script that processes Excel files and splits merged NOTES column values across all rows that the merge encompasses.

## Features

- Reads Excel files (.xlsx, .xls)
- Detects merged cells in the NOTES column
- Splits merged NOTES values across all rows within the merge range
- **Preserves all styles, formatting, and cell properties**
- **Preserves embedded images** (using exceljs library)
- Preserves column widths and row heights
- Outputs processed data to a new Excel file
- Provides JSON preview of processed data

## Installation

```bash
npm install
```

## Usage

```bash
node splitExcelNotes.js <input-file> [output-file]
```

### Examples

```bash
# Process input.xlsx and save to output.xlsx
node splitExcelNotes.js input.xlsx output.xlsx

# Process input.xlsx and save to input_processed.xlsx (default)
node splitExcelNotes.js input.xlsx
```

## Expected Excel Format

The script expects an Excel file with the following column headers:

- PHOTO
- EV STYLE #
- MR STYLE #
- UNITS
- COLOR
- FINDINGS
- COST EA.
- COST
- NOTES (this column may contain merged cells)

## How It Works

1. Reads the Excel file using exceljs (preserves images, styles, formatting)
2. Identifies merged cells in the worksheet
3. Finds all merged cells specifically in the NOTES column
4. Extracts the NOTES value from the top-left cell of each merge
5. Applies that value to all rows within the merge range while preserving cell styles
6. Removes the merge definition for NOTES column (keeps other merges intact)
7. Outputs the processed data to a new Excel file with all original formatting preserved

## Technical Details

- Uses **exceljs** library for full support of Excel features including embedded images
- All cell styles, fonts, colors, borders, and formatting are preserved
- Images embedded in cells (especially PHOTO column) are automatically preserved
- Column widths and row heights are maintained

## Output

The script generates:

- A new Excel file with NOTES values split across all rows
- Console output showing:
  - Number of processed rows
  - Column headers
  - Preview of first 3 processed rows
