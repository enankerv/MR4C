import XLSX from "xlsx";
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";

/**
 * Processes an Excel file and splits merged NOTES column across all rows
 * @param {string} inputFilePath - Path to the input Excel file
 * @param {string} outputFilePath - Path to the output Excel file (optional)
 * @returns {Array} Processed data array
 */
function splitExcelNotes(inputFilePath, outputFilePath = null) {
  // Read the Excel file
  const workbook = XLSX.readFile(inputFilePath);
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];

  // Get the range of the sheet
  const range = XLSX.utils.decode_range(worksheet["!ref"]);

  // Convert sheet to JSON to get the data
  const data = XLSX.utils.sheet_to_json(worksheet, {
    header: 1,
    defval: null,
    raw: false,
  });

  // Get headers (first row)
  const headers = data[0] || [];
  const notesColumnIndex = headers.indexOf("NOTES");

  if (notesColumnIndex === -1) {
    throw new Error("NOTES column not found in the Excel file");
  }

  // Get merged cells information
  const mergedCells = worksheet["!merges"] || [];

  // Create a map of merged cells for the NOTES column
  // Key: row index, Value: merged cell info
  const notesMerges = new Map();

  mergedCells.forEach((merge) => {
    const startCol = merge.s.c; // Start column
    const endCol = merge.e.c; // End column
    const startRow = merge.s.r; // Start row (0-indexed, but header is row 0)
    const endRow = merge.e.r; // End row

    // Check if this merge is in the NOTES column
    if (startCol === notesColumnIndex && endCol === notesColumnIndex) {
      // Store the merge info for all rows in the range
      for (let row = startRow; row <= endRow; row++) {
        notesMerges.set(row, {
          startRow: startRow,
          endRow: endRow,
          value: null, // Will be filled from the start row
        });
      }
    }
  });

  // Extract the NOTES value from merged cells
  // The value is typically stored in the top-left cell of the merge
  notesMerges.forEach((mergeInfo, rowIndex) => {
    if (rowIndex === mergeInfo.startRow) {
      // Get the value from the start row
      const cellAddress = XLSX.utils.encode_cell({
        r: rowIndex,
        c: notesColumnIndex,
      });
      const cell = worksheet[cellAddress];
      mergeInfo.value = cell ? cell.v || cell.w || "" : "";

      // Propagate this value to all rows in the merge
      for (let r = mergeInfo.startRow; r <= mergeInfo.endRow; r++) {
        const existingInfo = notesMerges.get(r);
        if (existingInfo) {
          existingInfo.value = mergeInfo.value;
        }
      }
    }
  });

  // Process the data rows (skip header row)
  const processedData = [];

  // Add header row
  processedData.push(headers);

  // Process each data row
  for (let rowIndex = 1; rowIndex < data.length; rowIndex++) {
    const row = [...data[rowIndex]]; // Copy the row

    // If this row is part of a merged NOTES cell, apply the merged value
    if (notesMerges.has(rowIndex)) {
      const mergeInfo = notesMerges.get(rowIndex);
      row[notesColumnIndex] = mergeInfo.value || "";
    } else {
      // If not merged, keep the existing value or set to empty string
      row[notesColumnIndex] = row[notesColumnIndex] || "";
    }

    // Ensure row has the correct length (pad with empty strings if needed)
    while (row.length < headers.length) {
      row.push("");
    }

    processedData.push(row);
  }

  // If output path is provided, write to a new Excel file
  if (outputFilePath) {
    // Create a new worksheet from the processed data
    const newWorksheet = XLSX.utils.aoa_to_sheet(processedData);

    // Create a new workbook
    const newWorkbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, sheetName);

    // Write to file
    XLSX.writeFile(newWorkbook, outputFilePath);
    console.log(`Processed file saved to: ${outputFilePath}`);
  }

  // Convert to JSON format for easier inspection
  const jsonData = [];
  for (let i = 1; i < processedData.length; i++) {
    const row = processedData[i];
    const rowObj = {};
    headers.forEach((header, index) => {
      rowObj[header] = row[index] || "";
    });
    jsonData.push(rowObj);
  }

  return {
    headers,
    data: jsonData,
    rawData: processedData,
  };
}

// Main execution
function main() {
  const args = process.argv.slice(2);

  if (args.length === 0) {
    console.log("Usage: node splitExcelNotes.js <input-file> [output-file]");
    console.log("Example: node splitExcelNotes.js input.xlsx output.xlsx");
    process.exit(1);
  }

  const inputFile = args[0];
  const outputFile =
    args[1] || inputFile.replace(/\.(xlsx|xls)$/i, "_processed.$1");

  if (!fs.existsSync(inputFile)) {
    console.error(`Error: Input file "${inputFile}" not found`);
    process.exit(1);
  }

  try {
    console.log(`Reading Excel file: ${inputFile}`);
    const result = splitExcelNotes(inputFile, outputFile);

    console.log(`\nProcessed ${result.data.length} rows`);
    console.log(`Headers: ${result.headers.join(", ")}`);

    // Display first few rows as preview
    console.log("\nPreview of processed data (first 3 rows):");
    result.data.slice(0, 3).forEach((row, index) => {
      console.log(`Row ${index + 1}:`, JSON.stringify(row, null, 2));
    });

    console.log(
      `\n✓ Successfully processed file. Output saved to: ${outputFile}`
    );
  } catch (error) {
    console.error("Error processing file:", error.message);
    process.exit(1);
  }
}

// Run if executed directly (not imported as module)
const __filename = fileURLToPath(import.meta.url);
const isMainModule =
  process.argv[1] && path.resolve(process.argv[1]) === path.resolve(__filename);

if (isMainModule) {
  main();
}

export { splitExcelNotes };
