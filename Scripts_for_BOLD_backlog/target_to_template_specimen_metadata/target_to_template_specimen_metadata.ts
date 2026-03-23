function main(workbook: ExcelScript.Workbook) {

    let targetSheet = workbook.getWorksheet("target");
    let template = workbook.getWorksheet("template");

    let targetData = targetSheet.getUsedRange().getValues();
    let templateData = template.getUsedRange().getValues();

    // get the number of columns in the template
    let templateColCount = template.getUsedRange().getColumnCount();
    let outputRows: (string | number | boolean)[][] = [];

    for (let h = 1; h < templateData.length; h++) {
        let templateId = templateData[h][0];

        // copy existing row data as the base so unchanged columns stay intact
        let newRow: (string | number | boolean)[] = [...templateData[h] as (string | number | boolean)[]];

        for (let j = 1; j < targetData.length; j++) {
            let targetId = targetData[j][0];

            if (templateId == targetId) {

                // pull all source columns
                let colC  = targetData[j][2];
                let colD  = targetData[j][3];
                let colF  = targetData[j][5];
                let colE  = targetData[j][4];
                let colG  = targetData[j][6];
                let colH  = targetData[j][7];
                let colI  = targetData[j][8];
                let colM  = targetData[j][12];
                let colO  = targetData[j][14];
                let colS  = targetData[j][18];
                let colU  = targetData[j][20];
                let colAC = targetData[j][28];
                let colAD = targetData[j][29];
                let colAE = targetData[j][30];
                let colAF = targetData[j][31];
                let colAG = targetData[j][32];
                let colAH = targetData[j][33];
                let colAJ = targetData[j][35];
                let colAK = targetData[j][36];
                let colAL = targetData[j][37];
                let colB  = targetData[j][1];
                let colJ  = targetData[j][9];
                let colN  = targetData[j][13];
                let colP  = targetData[j][15];

                // extract genus from species string
                let genusOnly: string | number | boolean;
                if (colM && typeof colM === 'string') {
                    genusOnly = colM.trim().split(' ')[0] || colM;
                } else {
                    genusOnly = colM;
                }

                // allocate template columns to target
                newRow[2]  = colC;   // C  - Museum ID
                newRow[1]  = colB;   // B  - FieldID
                newRow[3]  = colD;   // D  - Collection code
                newRow[4]  = colE;   // E  - Institution Storing
                newRow[6]  = colG;   // G  - Phylum
                newRow[7]  = colH;   // H  - Class
                newRow[8]  = colI;   // I  - Order
                newRow[9]  = colJ;   // J  - Family
                newRow[12] = genusOnly; // M - Genus
                newRow[13] = colN;   // N  - Species
                newRow[15] = colP;   // P  - Identified by
                newRow[18] = colS;   // S  - Sex
                newRow[20] = colU;   // U  - Life stage
                newRow[28] = colAC;  // AC - Collectors
                newRow[29] = colAD;  // AD - Collection date from
                newRow[30] = colAE;  // AE - Collection date to
                newRow[31] = colAF;  // AF - Country/Ocean
                newRow[32] = colAG;  // AG - Province/State
                newRow[33] = colAH;  // AH - Region
                newRow[35] = colAJ;  // AJ - Site
                newRow[36] = colAK;  // AK - Lat
                newRow[37] = colAL;  // AL - Long

                break;
            }
        }

        outputRows.push(newRow);
    }

    // writes into all columns and rows in template at once
    let writeRange = template.getRangeByIndexes(
        1,                    
        0,                    
        outputRows.length,    
        templateColCount   
    );
    writeRange.setValues(outputRows);
}
