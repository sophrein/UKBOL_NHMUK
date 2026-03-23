# UKBOL Accelerated — Sophia's Handover README

This is the handover README for all of the leftover folders as part of my role in UKBOL Accelerated. Included are documents and folders that were used in handling the data backlog uploads to BOLD and EMu.

---

## Folder Structure

### BGE
Spreadsheets for uploading and editing BGE data into BOLD Systems.

---

### Image BOLD Upload
All image folders and image metadata templates for BOLD image upload.

#### Submitted to BOLD
Image folders uploaded to BOLD during Sophia's contract for UKBOL Accelerated.
- Each folder is labelled after its original backlog folder
- Contains compressed images and an `ImageData.xls` metadata sheet

---

### In BOLD
All specimen metadata spreadsheets successfully submitted to the BOLD database.

---

### Submitted to BOLD
Spreadsheets submitted to BOLD that have **not yet been reviewed** by the BOLD data team.


---

### Unfinished BOLD Template Uploads
Templates that need editing before submission to BOLD.

- "Unfinished" typically means a **required field is missing metadata**
- The missing metadata needs to be located before the spreadsheet can be submitted


---

### Templates
| Template | Source |
|---|---|
| Specimen metadata template | Provided by BOLD |
| Image metadata template | Provided by BOLD |
| EMu submission template | Provided by Hillery |

---

## Documents

### `all_meeting_notes.docx`
Notes from meetings between Sophia and Ben.

---

### `image_checking_spreadsheet.xlsx`
Spreadsheet for processing image metadata for BOLD image submissions.

| Sheet | Description |
|---|---|
| `BOLD_DOWNLOAD` | All current BOLD records |
| `ImageData` | Template for filling in image metadata |
| `Target` | Paste image file names here to format data for template population |

---

### `SOP submitting images to BOLD.docx`
Standard Operating Procedure for submitting images to BOLD for UKBOL Accelerated.

---

### `SOP submitting specimen data to BOLD.docx`
Standard Operating Procedure for submitting specimen data to BOLD for UKBOL Accelerated.

---

### `specimen_metadata_checking_spreadsheet.xlsx`
Spreadsheet for processing specimen metadata for BOLD submission.

| Sheet | Description |
|---|---|
| `BOLD_DOWNLOAD` | All current BOLD records |
| `template` | Template for specimen metadata |
| `target` | Paste target sheet here; scripts identify missing sample IDs and copy their metadata into the template |

---

### `Tracking_Sheet.xlsx`
Tracks the submission of specimen metadata and images into BOLD Systems v.5.

---

## Workflows

### 1. Finding Missing IDs from a Target Spreadsheet

**What it does:** Compares a target spreadsheet against the BOLD database download and returns a list of IDs present in the target but missing from BOLD.

**Requirements:**
- Both the target worksheet and BOLD database download must be in the **same Excel file as separate tabs**
- Name the target tab: `target`
- Name the BOLD download tab: `BOLD_DOWNLOAD_{date}` (e.g. `BOLD_DOWNLOAD_25.11.25`)

**Steps:**
1. Go to the [UKBOL GitHub repository](https://github.com/sophrein/UKBOL_NHMUK/tree/main)
2. Copy the script: [`find_missing_IDs_specimen_metadata.ts`](https://github.com/sophrein/UKBOL_NHMUK/blob/main/Scripts_for_BOLD_backlog/find_missing_IDs_specimen_metadata.ts)
3. Open a new Script in Excel's **Automate** tab, paste the code, and click **Run**

**Output:** Two new sheets are created:
- `Missing IDs` — IDs to be submitted to BOLD
- `Duplicated IDs` — IDs already present, to avoid duplicate submissions

4. Copy the Missing IDs list into Column A (`SAMPLEID`) of the BOLD template

---

### 2. Auto-Populating the BOLD Template with Missing IDs

**What it does:** Imports all metadata associated with the missing IDs into the BOLD template ready for submission.

**Requirements:**
- The BOLD template must already have the Missing IDs filled in Column A (`SAMPLEID`)
- The target sheet must contain all metadata (including records already in BOLD — the script will only extract the missing ones)

**Steps:**
1. Copy the script: [`target_to_template_specimen_metadata.ts`](https://github.com/sophrein/UKBOL_NHMUK/blob/main/Scripts_for_BOLD_backlog/target_to_template_specimen_metadata.ts)
2. Open a new Script in Excel's **Automate** tab, paste the code, and click **Run**
3. **Validate:** Cross-check at least **10 sample IDs** against the target sheet to confirm metadata has been imported into the correct columns

---

## Typescripts that may be useful for excel automation

[Scripts_for_BOLD_backlog](https://github.com/sophrein/UKBOL_NHMUK/tree/main/Scripts_for_BOLD_backlog)

## Contacts
sophiareinisch@gmail.com

ben.price@nhm.ac.uk