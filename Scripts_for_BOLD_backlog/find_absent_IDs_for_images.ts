function main(workbook: ExcelScript.Workbook) {
    // get the active cell and worksheet
    let selectedCell = workbook.getActiveCell();
    let selectedSheet = workbook.getActiveWorksheet();

    // assigning the two sheets to variable names
    let targetSheet = workbook.getWorksheet("target");
    let boldDownloadSheet = workbook.getWorksheet("BOLD_DOWNLOAD_05.01.26");

    // define the columns in both sheets int variables
    let targetIdColumn = 0;
    let boldIdColumn = 2;

    // prepares and retrives all of the filled cells
    let targetSheetRange = targetSheet.getUsedRange();
    let boldDownloadRange = boldDownloadSheet.getUsedRange();
    // returns all of the data within the cells outlines
    let targetSheetValues = targetSheetRange.getColumn(targetIdColumn).getValues();
    let boldDownloadValues = boldDownloadRange.getColumn(boldIdColumn).getValues();

    // create an empty set with no duplicates to store the unique IDs
    let boldIds = new Set<string>();

    // looping through each row and if there is data it converts it to a string and creates an array of IDs from BOLD database download
    for (let i = 1; i < boldDownloadValues.length; i++) {
        // make sure that the cell is not empty
        if (boldDownloadValues[i][0] !== null && boldDownloadValues[i][0] !== "") {
            let prefixIds = boldDownloadValues[i][0].toString().trim();
            boldIds.add(prefixIds);
        }
    }

    // find the IDs in target sheet that are not in the bold download
    let missingIds: (string | number)[][] = [];
    for (let j = 1; j < targetSheetValues.length; j++) {
        let id = targetSheetValues[j][0];
        if (id !== null && id != "" && typeof id !== "boolean") {
            let idAsString = id.toString().trim();

            // checking if the id is not in the BOLD database worksheet
            if (!boldIds.has(idAsString)) {
                missingIds.push([id]); // push the missing IDs
            }
        }
    }

    // create a sheet that stores the results
    let resultSheet = workbook.getWorksheet("Missing IDs");
    if (resultSheet) {
        resultSheet.delete();
    }
    resultSheet = workbook.addWorksheet("Missing IDs"); 

    if (missingIds.length > 0) {
        resultSheet.getRange(`A2:A${missingIds.length + 1}`).setValues(missingIds);

        // make sure there is no duplicates of IDs
        let resultRange = resultSheet.getUsedRange();
        const result = resultRange.removeDuplicates([0], true);
    }
}
