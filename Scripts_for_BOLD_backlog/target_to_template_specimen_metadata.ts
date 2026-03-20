function main(workbook: ExcelScript.Workbook) {
    // get the active cell and worksheet
    let selectedCell = workbook.getActiveCell();
    let selectedSheet = workbook.getActiveWorksheet();

    // assign relevant sheets involved to variable names
    let targetSheet = workbook.getWorksheet("target");
    let template = workbook.getWorksheet("BOLD_template");

    // to get all the data from the target worksheet (with metadata we want to extract) and also the data currently in the column A of the template
    let targetData = targetSheet.getUsedRange().getValues();
    let templateData = template.getUsedRange().getValues();

    for (let h = 1; h < templateData.length; h++) {
        let templateId = templateData[h][0];

        // remove 'NHMUK0' prefix from template ID for comparison (UNCOMMENT IF NECESSARY (if the target barcodes do not have 'NHMUK0' prefix))
        //let templateIdStripped = typeof templateId === 'string'
            //? templateId.replace(/^NHMUK0/, '')
            //: templateId;

        // loop through each row of target sheet starting from 1
        for (let j = 1; j < targetData.length; j++) {
            let targetId = targetData[j][0];

            if (templateId == targetId) {
                // assign column numbers to variable names col(letter of column)
                let colA = targetData[j][0];
                let colC = targetData[j][2];
                let colF = targetData[j][5];
                let colI = targetData[j][8];
                let colJ = targetData[j][9];
                let colK = targetData[j][10];
                let colX = targetData[j][23];
                let colY = targetData[j][24];
                let colZ = targetData[j][25];
                let colAA = targetData[j][26];
                let colAB = targetData[j][27];
                let colAG = targetData[j][32];
                let colAH = targetData[j][33];
                let colAI = targetData[j][34];
                let colAJ = targetData[j][35];
                let colAK = targetData[j][36];
                let colAL = targetData[j][37];
                let colAO = targetData[j][40];
                let colAQ = targetData[j][42];
                let colAS = targetData[j][44];
                let colAT = targetData[j][45];
                let colAV = targetData[j][47];
                let colE = targetData[j][4];
                let colD = targetData[j][3];
                let colU = targetData[j][20];
                let colV = targetData[j][21];
                let colW = targetData[j][22];
                let colAC = targetData[j][28];
                let colAE = targetData[j][30];
                let colL = targetData[j][11];
                let colT = targetData[j][19];
                let colG = targetData[j][6];
                let colQ = targetData[j][16];
                let colH = targetData[j][7];
                let colM = targetData[j][12];
                let colN = targetData[j][13];
                let colO = targetData[j][14];
                let colAN = targetData[j][39];
                let colAU = targetData[j][46];
                let colAD = targetData[j][29];
                let colAF = targetData[j][31];
                let colAM = targetData[j][38];
                let colP = targetData[j][15];
                let colAR = targetData[j][43];
                let colS = targetData[j][18];

                let templateRow = h + 1;
               
                // Write to template:

                // Collection date from
                template.getRange(`AD${templateRow}`).setValue(colAD);

                // Collection date to
                template.getRange(`AE${templateRow}`).setValue(colAE);

                // Collectors
                template.getRange(`AC${templateRow}`).setValue(colAC);

                // Country/Ocean
                template.getRange(`AF${templateRow}`).setValue(colAF);

                // Province/State
                template.getRange(`AG${templateRow}`).setValue(colAG);

                // Region
                template.getRange(`AH${templateRow}`).setValue(colAH);

                // Site
                template.getRange(`AJ${templateRow}`).setValue(colAJ);

                // Lat
                template.getRange(`AK${templateRow}`).setValue(colAK);

                // Long
                template.getRange(`AL${templateRow}`).setValue(colAL);

                // Phylum
                template.getRange(`G${templateRow}`).setValue(colF);

                // Class
                template.getRange(`H${templateRow}`).setValue(colG);

                // Order
                template.getRange(`I${templateRow}`).setValue(colH);

                // Family
                template.getRange(`J${templateRow}`).setValue(colI);

                // Genus
                template.getRange(`M${templateRow}`).setValue(colL);

                // Identified by
                template.getRange(`P${templateRow}`).setValue(colO);

                // Identification method
                template.getRange(`Q${templateRow}`).setValue(colQ);

                // Taxonomy notes
                template.getRange(`R${templateRow}`).setValue(colR);

                // Life stage
                template.getRange(`U${templateRow}`).setValue(colU);

                // Sex 
                template.getRange(`S${templateRow}`).setValue(colS);

                // Sample ID
                template.getRange(`A${templateRow}`).setValue("NHMUK" + colE);

                // Collection code
                template.getRange(`D${templateRow}`).setValue(colD);

                // Museum ID
                template.getRange(`C${templateRow}`).setValue(colC); 

                // Species
                template.getRange(`N${templateRow}`).setValue(colM);

                // extract genus from genus + species
                 let genusOnly: string | number | boolean;
                 if (colM && typeof colM === 'string') {
                  genusOnly = colM.trim().split(' ')[0] || colM;
                  } else {
                 genusOnly = colM;
                 }

                template.getRange(`M${templateRow}`).setValue(genusOnly);
            }
        }
    }

    // get all of the cells and their data
    let finalTemplateData = template.getUsedRange().getValues();

    // going down the list of column 2
    for (let k = 1; k < finalTemplateData.length; k++) {
        let targetId = finalTemplateData[k][2];

    }
}
