# Populate Specimen Metadata Template — Excel Office Script

Matches sample IDs from a `template` sheet against a `target` sheet and auto-populates the template with the associated specimen metadata.

## Requirements

- Both sheets must be in the **same Excel workbook** as separate tabs
- Sheets must be named exactly as follows:

| Sheet Name | Description |
|---|---|
| `target` | Source sheet containing all specimen metadata |
| `template` | BOLD specimen metadata template to populate |

- The `template` sheet must already have **Sample IDs filled in Column A** before running

## Column Mapping

The script maps the following columns from `target` into `template`:

| Template Column | Field |
|---|---|
| B | Field ID |
| C | Museum ID |
| D | Collection Code |
| E | Institution Storing |
| G | Phylum |
| H | Class |
| I | Order |
| J | Family |
| M | Genus (extracted from the species string) |
| N | Species |
| P | Identified By |
| S | Sex |
| U | Life Stage |
| AC | Collectors |
| AD | Collection Date From |
| AE | Collection Date To |
| AF | Country/Ocean |
| AG | Province/State |
| AH | Region |
| AJ | Site |
| AK | Latitude |
| AL | Longitude |

> **Note:** The script extracts the genus automatically from the species name string — it takes the first word only.

## What It Does

1. Reads all rows from both the `target` and `template` sheets
2. For each row in the template, looks for a matching Sample ID in the target
3. If a match is found, copies the metadata into the correct template columns
4. If no match is found, the existing row data is left unchanged
5. Writes all updated rows back to the template in one operation

## Usage

1. Ensure the `template` sheet has Sample IDs populated in Column A
2. Open your Excel workbook containing the `target` and `template` sheets
3. Go to **Automate → New Script**
4. Paste the script and click **Run**
5. **Validate:** Cross-check at least 10 sample IDs against the target sheet to confirm the metadata has been imported into the correct columns