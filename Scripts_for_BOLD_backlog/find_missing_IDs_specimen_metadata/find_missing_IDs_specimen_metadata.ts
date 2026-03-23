function main(workbook: ExcelScript.Workbook) {
  
  // get the active cell and worksheet.
  let selectedCell = workbook.getActiveCell();
  let selectedSheet = workbook.getActiveWorksheet();

  // assigning the two sheets to variable names
  let targetSheet = workbook.getWorksheet("target");
  let boldDownloadSheet = workbook.getWorksheet("BOLD_DOWNLOAD");

  // define the columns in both sheets int variables
  let targetIdColumn = 0;
  let boldIdColumn = 2;

  // prepares and retrives all of the filled cells
  let targetSheetRange = targetSheet.getUsedRange();
  let boldDownloadRange = boldDownloadSheet.getUsedRange();
  // returns all of the data within the cells outlines
  let targetSheetValues = targetSheetRange.getColumn(targetIdColumn).getValues();
  let boldDownloadValues = boldDownloadRange.getColumn(boldIdColumn).getValues();

  // create an empty set (array with no duplicates) to store the unique IDs
  let boldIds = new Set<string>();

  // looping through each row and if there is data it converts it to a string and creates an array of IDs from BOLD database download and removes the NHMUK
  for (let i = 1; i < boldDownloadValues.length; i++) {
    // make sure that the cells are not empty
    if (boldDownloadValues[i][0] !== null) {
      let prefixIds = boldDownloadValues[i][0].toString(); // [i] = which row,[0] = which column in that row (0 since you only extracted one column)
      // remove NHMUK from BOLD IDs by replacing NHMUK with an empty string
      let numericPart = prefixIds.replace(/NHMUK/i, "").replace(/^0+/, "").trim();
      boldIds.add(numericPart);
    }
  }

  // find the missing IDs in target sheet that are not in the bold download
  let missingIds: (string | number)[][] = []; // creates an empty array that can handle both string and number incase there is a letter in the ID
  for (let j = 1; j < targetSheetValues.length; j++) {
    let id = targetSheetValues[j][0];
    if (id !== null && id != "" && typeof id !== "boolean") {
      let idAsString = id.toString().trim(); // have to convert it to string as its only possible to trim any spaces around the ID as a string
      // remove NHMUK from target IDs too so both sets match, works whether prefix is present or not
      let cleanedId = idAsString.replace(/NHMUK/i, "").replace(/^0+/, "").trim();

      // checking if the id is in the BOLD database worksheet
      if (!boldIds.has(cleanedId)) {
        missingIds.push([id]); // push the missing IDs
      }
    }
  }

  // create a sheet that store the results
  let resultSheet = workbook.getWorksheet("Missing IDs");
  if (resultSheet) {
    resultSheet.delete();
  }
  resultSheet = workbook.addWorksheet("Missing IDs");

  if (missingIds.length > 0) {
    resultSheet.getRange(`A2:A${missingIds.length + 1}`).setValues(missingIds);

    // make sure there is no duplicates of IDs
    let resultRange = resultSheet.getUsedRange();
    const result = resultRange.removeDuplicates([0], false);
  }

  // track how many times each ID appears in the target sheet
  let idCount: { [key: string]: number } = {};
  for (let j = 1; j < targetSheetValues.length; j++) {
    let id = targetSheetValues[j][0];
    if (id !== null && id != "" && typeof id !== "boolean") {
      // clean the ID the same way as above so comparisons are consistent
      let cleanedId = id.toString().trim().replace(/NHMUK/i, "").replace(/^0+/, "").trim();
      idCount[cleanedId] = (idCount[cleanedId] || 0) + 1;
    }
  }

  // collect IDs that appear more than once
  let duplicateIds: (string | number)[][] = [["Duplicate ID", "Count"]];
  for (let id in idCount) {
    if (idCount[id] > 1) {
      duplicateIds.push([id, idCount[id]]);
    }
  }

  // create a sheet to store the duplicate results
  let duplicateSheet = workbook.getWorksheet("Duplicate IDs");
  if (duplicateSheet) {
    duplicateSheet.delete();
  }
  duplicateSheet = workbook.addWorksheet("Duplicate IDs");

  if (duplicateIds.length > 1) {
    duplicateSheet.getRange(`A1:B${duplicateIds.length}`).setValues(duplicateIds);
  } else {
    // if no duplicates found, write a message
    duplicateSheet.getRange("A1").setValue("No duplicate IDs found in target sheet.");
  }

}
