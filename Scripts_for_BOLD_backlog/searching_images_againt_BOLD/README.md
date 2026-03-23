# Image ID Lookup Scripts — Excel Office Scripts

Two scripts that compare a target sheet against a BOLD database download to identify which image IDs are missing from BOLD, and which are already in BOLD.

## Requirements

- Both sheets must be in the **same Excel workbook** as separate tabs
- Sheets must be named exactly as follows:

| Sheet Name | Description |
|---|---|
| `target` | Source sheet with image IDs in **Column A** |
| `BOLD_DOWNLOAD` | BOLD database download with image IDs in **Column C** |

---

## Script 1 — find_absent_IDs_for_images

Finds image IDs that are in the `target` sheet but **not** in the BOLD database.

### What It Does

1. Reads all IDs from Column C of `BOLD_DOWNLOAD`
2. Compares them against Column A of `target`
3. Writes any IDs not found in BOLD to a new sheet, removing duplicates

### Output

| Sheet | Description |
|---|---|
| `Missing IDs` | IDs in `target` that are not in BOLD |

---

## Script 2 — find_present_IDs_for_images

Finds image IDs that are in the `target` sheet and already exist in the BOLD database.

### What It Does

1. Reads all IDs from Column C of `BOLD_DOWNLOAD`
2. Compares them against Column A of `target`
3. Writes any IDs that match BOLD to a new sheet, removing duplicates

### Output

| Sheet | Description |
|---|---|
| `Found IDs` | IDs in `target` that are already in BOLD |

---

## Usage

1. Open your Excel workbook containing the `target` and `BOLD_DOWNLOAD` sheets
2. Go to **Automate → New Script**
3. Paste the relevant script and click **Run**
4. Check the output sheet for results

> **Note:** Each time a script is run, it deletes and recreates the output sheet. Any manual edits to `Missing IDs` or `Found IDs` will be overwritten.