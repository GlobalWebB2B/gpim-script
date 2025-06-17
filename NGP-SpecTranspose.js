function main(workbook: ExcelScript.Workbook) {
    // Get the active sheet
    const sheet = workbook.getActiveWorksheet();

    // Get the range of data
    const range = sheet.getUsedRange();
    const values: string[][] = range.getValues() as string[][];

    // Create transposed Value
    const transposedValues: string[][] = transpose(values);

    // Create a new sheet
    const newSheet = workbook.addWorksheet("Transposed Data");

    // Write the transposed data into the new sheet
    const newRange = newSheet.getRangeByIndexes(0, 0, transposedValues.length, transposedValues[0].length);
    newRange.setValues(transposedValues);
}

// Function to transpose
function transpose(array: string[][]): string[][] {
    return array[0].map((_, colIndex) => array.map(row => row[colIndex]));
}