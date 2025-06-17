function main(workbook: ExcelScript.Workbook) {
    // Get all worksheets in the workbook
    const sheets: ExcelScript.Worksheet[] = workbook.getWorksheets();

    // Check if at least 2 sheets exist
    if (sheets.length < 2) {
        throw new Error("At least 2 sheets are required in the workbook.");
    }

    // Get the first two sheets
    const sheet1: ExcelScript.Worksheet = sheets[0];
    const sheet2: ExcelScript.Worksheet = sheets[1];

    // Get the used ranges
    const range1: ExcelScript.Range = sheet1.getUsedRange();
    const range2: ExcelScript.Range = sheet2.getUsedRange();

    const values1: string[][] = range1.getValues() as string[][];
    const values2: string[][] = range2.getValues() as string[][];

    // Get the maximum number of rows and columns
    const maxRows: number = Math.max(values1.length, values2.length);
    const maxCols: number = Math.max(values1[0].length, values2[0].length);

    // Set the background color to yellow for differing cells
    for (let row: number = 0; row < maxRows; row++) {
        for (let col: number = 0; col < maxCols; col++) {
            const value1: string | undefined = values1[row] ? values1[row][col] : undefined;
            const value2: string | undefined = values2[row] ? values2[row][col] : undefined;

            if (value1 !== value2) {
                // Set the background color to yellow for differing cells in sheet1
                if (row < values1.length && col < values1[0].length) {
                    sheet1.getRangeByIndexes(row, col, 1, 1).getFormat().getFill().setColor("yellow");
                }
                // Set the background color to yellow for differing cells in sheet2
                if (row < values2.length && col < values2[0].length) {
                    sheet2.getRangeByIndexes(row, col, 1, 1).getFormat().getFill().setColor("yellow");
                }
            }
        }
    }
}