function main(workbook: ExcelScript.Workbook) {
  // assign relevant sheets involved to variable names
  let targetSheet = workbook.getWorksheet("target");
  let template = workbook.getWorksheet("ImageData");
  let foundIdsSheet = workbook.getWorksheet("Found IDs");

  // check if sheets exist and have data
  if (!targetSheet || !template || !foundIdsSheet) {
    console.log("ERROR: One or more sheets not found");
    return;
  }

  let targetRange = targetSheet.getUsedRange();
  let foundIdsRange = foundIdsSheet.getUsedRange();
  
  if (!targetRange || !foundIdsRange) {
    console.log("ERROR: One or more sheets are empty");
    return;
  }

  let targetData = targetRange.getValues();
  let foundIdsData = foundIdsRange.getValues();

  // create a set of unique IDs from Found IDs sheet
  let uniqueIds = new Set<string>();
  
  for (let i = 1; i < foundIdsData.length; i++) {
    let id = foundIdsData[i][0];
    // checking if the cell is empty, and converts the number to string in order to trim any spaces around the text etc.
    if (id !== null && id !== "") {
      uniqueIds.add(id.toString().trim());
      // uniqueIds contains now each ID without duplicates
    }
  }

  let templateRow = 2;

  // loop through each unique ID
  uniqueIds.forEach(uniqueId => {
    // search through all of the row in the target from column A (0)
    for (let j = 1; j < targetData.length; j++) {
      let targetId = targetData[j][0];
      
      // convert targetID to string in order to trim around and or skip if the cell is empty
      if (targetId !== null && targetId !== "") {
        let targetIdString = targetId.toString().trim();
        
        // if the unique ID is the same as the target ID from the target sheet
        if (uniqueId == targetIdString) {
          // define variables of metadata
          let secondaryText = targetData[j][1]; // dorsal/ventral
          let imageFilename = targetData[j][2]; // image filename
          
          // Write to template
          template.getRange(`A${templateRow}`).setValue(imageFilename);
          template.getRange(`C${templateRow}`).setValue(secondaryText);
          template.getRange(`G${templateRow}`).setValue(targetId);
          
          templateRow++;
          // Don't break: continue to find all images for this ID until there is no assoicated metadata for this unique ID
        }
      }
    }
  });
}
