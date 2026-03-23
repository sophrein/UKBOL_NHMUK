function main(workbook: ExcelScript.Workbook) {
    // get the active worksheet
    let selectedSheet = workbook.getActiveWorksheet();

    // get the used range to determine how many rows have data
    let usedRange = selectedSheet.getUsedRange();

    // error if the worksheet has no data
    if (!usedRange) {
        console.log("Error: No data found in worksheet.");
        return;
    }

    // get the number of rows with data
    let rowCount = usedRange.getRowCount();

    // get only column A for the used range
    let columnA = selectedSheet.getRange("A1:A" + rowCount);
    // get all values from column A
    let values = columnA.getValues();

    // check if there are values
    if (!values || values.length === 0) {
        console.log("Error: No data found in column A.");
        return;
    }

    // loop through each cell in column A and add "NHMUK" prefix
    for (let i = 0; i < values.length; i++) {
        let cellValue = values[i][0];  // column A is index 0
        // Only add prefix if cell is not empty
        if (cellValue !== null && cellValue !== undefined && cellValue !== "") {
            values[i][0] = "NHMUK0" + cellValue.toString();
        }
    }

    // write the modified values back to column A
    columnA.setValues(values);
}
