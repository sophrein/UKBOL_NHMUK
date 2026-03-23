# Populate Image Metadata Template — Excel Office Script

Looks up image metadata from a target sheet using a list of found IDs, and writes the results into the `ImageData` template ready for BOLD image upload.

## Requirements

- All three sheets must be in the **same Excel workbook** as separate tabs
- Sheets must be named exactly as follows:

| Sheet Name | Description |
|---|---|
| `target` | Source sheet containing all image metadata |
| `ImageData` | BOLD image metadata template to populate |
| `Found IDs` | Sample IDs which are in BOLD |

### Target Sheet Column Layout

| Column | Contents |
|---|---|
| A | Image ID |
| B | View Metadata (e.g. dorsal/ventral) |
| C | Image filename |

## What It Does

1. Reads all IDs from the `Found IDs` sheet and deduplicates them
2. For each unique ID, searches the `target` sheet for all matching rows
3. For every match, writes the following to the `ImageData` template starting at row 2:

| Template Column | Value Written |
|---|---|
| A | Image filename |
| C | View Metadata (e.g. dorsal/ventral) |
| G | Image ID |

> **Note:** If an ID has multiple associated images, all of them are written — the script does not stop at the first match.

## Usage

1. Open your Excel workbook containing the `target`, `ImageData`, and `Found IDs` sheets
2. Go to **Automate → New Script**
3. Paste the script and click **Run**
4. Check the `ImageData` sheet for the populated results