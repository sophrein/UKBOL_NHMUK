# Ascension Island Image Metadata Script

An Office Scripts automation that matches image IDs from a target sheet against specimen metadata, then writes the results to a template sheet.

## Sheets Required

| Sheet Name | Description |
|---|---|
| `target` | Source data with image ID (col A), image name (col B), and view metadata (col C) |
| `ascension_island_metadata` | Lookup table with sample ID (col A) and image ID (col B) |
| `template` | Output sheet where matched results are written |

## What It Does

1. Reads all rows from `target` and `ascension_island_metadata`
2. Builds a lookup map of image ID → sample ID from the metadata sheet
3. For each row in `target`, checks if the image ID exists in the metadata
4. If matched, writes to `template` starting at row 2:
   - **Column A** — Image name
   - **Column C** — View metadata
   - **Column G** — Sample ID

Rows with no matching image ID are skipped.

## Usage

1. Open the workbook in Excel (desktop or web)
2. Go to **Automate → New Script**, paste the script, and click **Run**