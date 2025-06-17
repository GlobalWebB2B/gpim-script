function main(workbook: ExcelScript.Workbook) {
  // Get the active sheet
  const sheet = workbook.getActiveWorksheet();

  // Get the range of data
  const range = sheet.getUsedRange();
  const values: string[][] = range.getValues() as string[][];

  // Remove columns A and B (index 0 and 1)
  const trimmedValues = values.map(row => row.slice(2));

  // Transpose the trimmed data
  let transposedValues: string[][] = transpose(trimmedValues);

  // Find the first row where column A (index 0) contains a non-empty string
  let cutoffIndex = transposedValues.findIndex(row => row[0].toString().trim() !== "");

  if (cutoffIndex !== -1) {
    // Remove all rows from cutoffIndex onwards
    transposedValues = transposedValues.slice(0, cutoffIndex);
  }

  // Remove column A (index 0) from transposedValues
  transposedValues = transposedValues.map(row => row.slice(1));

  // Replace newline characters with <br>
  transposedValues = transposedValues.map(row =>
    row.map(cell =>
      typeof cell === "string"
        ? cell.replace(/\r\n|\n|\r/g, "<br>")
        : cell
    )
  );

  // Create a new sheet
  const newSheet = workbook.addWorksheet("Transposed Data");

  // Write the cleaned transposed data into the new sheet
  const newRange = newSheet.getRangeByIndexes(0, 0, transposedValues.length, transposedValues[0].length);
  newRange.setValues(transposedValues);
}

// Function to transpose
function transpose(array: string[][]): string[][] {
  return array[0].map((_, colIndex) => array.map(row => row[colIndex]));
}
