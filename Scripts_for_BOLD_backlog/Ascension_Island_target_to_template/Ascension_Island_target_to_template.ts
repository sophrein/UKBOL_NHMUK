function main(workbook: ExcelScript.Workbook) {
    // get sheets
    const targetSheet = workbook.getWorksheet("target");
    const metadataSheet = workbook.getWorksheet("ascension_island_metadata");
    const templateSheet = workbook.getWorksheet("template");

    // get all data from target sheet (columns A and B)
    const targetData = targetSheet.getUsedRange().getValues();

    // Get all data from metadata sheet (columns A and B)
    const metadataData = metadataSheet.getUsedRange().getValues();

    // lookup map between metadata of imageID and sample ID
    const metadataMap = new Map<string, string>();

    for (let i = 1; i < metadataData.length; i++) {
        const sampleId = metadataData[i][0];
        const metaImageId = metadataData[i][1]; 

        if (metaImageId !== null && metaImageId !== "") {
            metadataMap.set(metaImageId.toString().trim(), sampleId.toString().trim());
        }
    }

    // start writing data from row 2
    let templateRow = 2;

    // loop through target sheet rows
    for (let j = 1; j < targetData.length; j++) {
        const imageId = targetData[j][0];   // column A: image ID
        const imageName = targetData[j][1]; // column B: image name
        const viewData = targetData[j][2]; // column C: view metadata

        if (imageId === null || imageId === "") continue;

        const imageIdStr = imageId.toString().trim();

        // check if this image ID exists in specimen metadata sheet
        if (metadataMap.has(imageIdStr)) {
            const sampleId = metadataMap.get(imageIdStr)!;

            // write image name to template column A
            templateSheet.getRange(`A${templateRow}`).setValue(imageName);

            // write sample_id to template column G
            templateSheet.getRange(`G${templateRow}`).setValue(sampleId);

            // write View metadata to column C
            templateSheet.getRange(`C${templateRow}`).setValue(viewData);


            templateRow++;
        }
    }
}
