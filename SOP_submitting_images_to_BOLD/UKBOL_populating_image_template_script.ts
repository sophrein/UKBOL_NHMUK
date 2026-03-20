function main(workbook: ExcelScript.Workbook) {
    let selectedSheet = workbook.getActiveWorksheet();

    // get the used range which is set to how many rows that data
    let usedRange = selectedSheet.getUsedRange();

    // get the number of rows with data
    let rowCount = usedRange.getRowCount();

    // text to columns on rows in column A - split by backslash
    let sourceRange = selectedSheet.getRange("A1:A" + rowCount);
    let destinationRange = selectedSheet.getRange("A1");

    // get all values from column A (all rows which have data in them)
    for (let row = 0; row < rowCount; row++) {
        let cellValue = sourceRange.getRow(row).getValues()[0][0].toString();

        // remove the extra .jpg if the filename ends with .jpg.jpg
        cellValue = cellValue.replace(/\.jpg\.jpg$/i, '.jpg');

        // split on backslash
        let sourceRangeValues = cellValue.split('\\');
        destinationRange.getOffsetRange(row, 0).getResizedRange(0, sourceRangeValues.length - 1).setValues([sourceRangeValues]);
    }

    // delete columns A:H
    selectedSheet.getRange("A:H").delete(ExcelScript.DeleteShiftDirection.left);

    // paste column A to column C
    selectedSheet.getRange("C:C").copyFrom(selectedSheet.getRange("A:A"), ExcelScript.RangeCopyType.all, false, false);

    // text to columns on column A split by underscore
    sourceRange = selectedSheet.getRange("A1:A" + rowCount);
    destinationRange = selectedSheet.getRange("A1");

    for (let row = 0; row < rowCount; row++) {
        let sourceRangeValues = sourceRange.getRow(row).getValues()[0][0].toString().split('_');
        destinationRange.getOffsetRange(row, 0).getResizedRange(0, sourceRangeValues.length - 1).setValues([sourceRangeValues]);
    }

    // remove all .jpg from column B
    selectedSheet.getRange("B:B").replaceAll(".*", "", { completeMatch: false, matchCase: false });

    let columnA = selectedSheet.getRange("A1:A" + rowCount);

    // get all values from column A
    let values = columnA.getValues();

    // add NHMUK0 prefix to the beginning of barcodes
    for (let i = 0; i < values.length; i++) {
        let cellValue = values[i][0];
        // Only add prefix if cell is not empty
        if (cellValue !== null && cellValue !== undefined && cellValue !== "") {
            let cellString = cellValue.toString();
            // Check if the value already starts with "0"
            if (cellString.startsWith("0")) {
                values[i][0] = "NHMUK" + cellString;
            } else {
                values[i][0] = "NHMUK0" + cellString;
            }
        }
    }

    // paste values back to column A
    columnA.setValues(values);

    // removing .jpg from all text in column A (if its there)
    selectedSheet.getRange("A:A").replaceAll(".jpg", "", { completeMatch: false, matchCase: false });

    // find the bold worksheet with the specimen data in it (for sample ID)
    let boldSheet = workbook.getWorksheet("bold");

    if (!boldSheet) {
        console.log("bold worksheet not found.");
        return;
    }

    // get values from MUSEUMID and SAMPLEID columns from BOLD sheet
    let boldUsedRange = boldSheet.getUsedRange();
    let boldValues = boldUsedRange.getValues();
    let boldHeaders = boldValues[0];

    // find column indices for MUSEUMID and SAMPLEID
    let museumIdColIndex = boldHeaders.indexOf("MUSEUMID");
    let sampleIdColIndex = boldHeaders.indexOf("SAMPLEID");

    // vlookup MUSEUMID -> SAMPLEID
    let lookupMap = new Map<string, string>();
    for (let i = 1; i < boldValues.length; i++) {  
        let museumId = boldValues[i][museumIdColIndex];
        let sampleId = boldValues[i][sampleIdColIndex];

        if (museumId !== null && museumId !== undefined && museumId !== "") {
            lookupMap.set(museumId.toString(), sampleId ? sampleId.toString() : "");
        }
    }

    // get barcodes from column A in target sheet and find matching SAMPLEIDs in BOLD sheet
    let targetBarcodes = selectedSheet.getRange("A1:A" + rowCount).getValues();
    let sampleIdResults: string[][] = [];

    for (let i = 0; i < targetBarcodes.length; i++) {
        let barcode = targetBarcodes[i][0];
        let matchingSampleId = "";

        if (barcode !== null && barcode !== undefined && barcode !== "") {
            matchingSampleId = lookupMap.get(barcode.toString()) || "";
        }

        sampleIdResults.push([matchingSampleId]);
    }

    // write SAMPLEID results to column D in target
    selectedSheet.getRange("D1").getResizedRange(sampleIdResults.length - 1, 0).setValues(sampleIdResults);
}

