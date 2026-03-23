# Find Missing IDs — Excel Office Script

Compares a target spreadsheet against a BOLD database download and identifies which sample IDs are missing from BOLD. Also flags any duplicate IDs within the target sheet.

## Requirements

- Both sheets must be in the **same Excel workbook** as separate tabs
- Name the target tab: `target`
- Name the BOLD database download tab: `BOLD_DOWNLOAD`
- The target sheet must have sample IDs in **Column A**
- The BOLD download must have sample IDs in **Column C**

## What It Does

1. Reads all sample IDs from the `target` sheet (Column A)
2. Reads all sample IDs from the `BOLD_DOWNLOAD` sheet (Column C)
3. Strips the `NHMUK` prefix and any leading zeros from IDs in both sheets so they can be compared consistently
4. Identifies IDs present in `target` but **not** in `BOLD_DOWNLOAD`
5. Identifies any IDs that appear **more than once** in the `target` sheet

## Output

Two new sheets are created in the workbook:

| Sheet | Description |
|---|---|
| `Missing IDs` | IDs in the target sheet that are not in BOLD |
| `Duplicate IDs` | IDs that appear more than once in the target sheet, with a count |

## Usage

1. Open your Excel workbook containing the `target` and `BOLD_DOWNLOAD` sheets
2. Go to **Automate → New Script**
3. Paste the script and click **Run**
4. Check the `Missing IDs` and `Duplicate IDs` sheets for results

## Next Step

Copy the `Missing IDs` list into **Column A (`SAMPLEID`)** of the BOLD template, then run to import the associated metadata.