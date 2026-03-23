# Add NHMUK Prefix — Excel Office Script

Adds the `NHMUK0` prefix to every filled cell in Column A of the active worksheet.

## Requirements

- Run the script on the worksheet you want to modify — it operates on the **currently active sheet**
- Sample IDs must be in **Column A**

## What It Does

1. Reads all values from Column A
2. For each non-empty cell, prepends `NHMUK0` to the existing value
3. Writes the updated values back to Column A

**Example:** `12345` → `NHMUK012345`

## Usage

1. Open your Excel workbook and navigate to the sheet you want to modify
2. Go to **Automate → New Script**
3. Paste the script and click **Run**
4. Column A will be updated in place with the `NHMUK0` prefix added

> **Note:** This script modifies the sheet directly — there is no undo once run. Consider making a copy of your data beforehand.