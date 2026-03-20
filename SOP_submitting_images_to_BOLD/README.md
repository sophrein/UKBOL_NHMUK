# Populating the Image Metadata template for BOLD Systems v.5 upload

Script that can be applied in Excel MS Office Script.

## Set Up

Go into your excel and click on the 'Automate' tab, then click on 'New Script' and paste the Type Script code into a script template.

Then you can Run and should work as long as you follow instructions in 'SOP submitting images to BOLD.pdf'.

## What this script does
### Overview
This script uses image file paths in Excel to extract museum barcodes, formats them, and looks up matching Sample IDs from a `bold` worksheet in order to subsequenlty be pasted into a BOLD systems v.5 image metadata template.

## Steps
 **Split File Paths**
   - Splits text in Column A using `\`
   - Fixes `.jpg.jpg` → `.jpg` (if needed)

 **Splits Barcode and View Metadata**
   - Splits the barcode from the view metadata using `_` (this value can be changed in line 37 if you want to split based on another character)

 **Cleans Text**
   - Removes `.jpg` from columns

 **Formats Barcodes**
   - Adds prefix:
     - Starts with `0` → `NHMUK`
     - Otherwise → `NHMUK0`

 **Lookup Sample IDs**
   - Uses `bold` sheet
   - Matches:
     - `MUSEUMID` → `SAMPLEID`

 **Output**
   - Writes Sample IDs to Column D

## Requirements
- Image file path in Column A
- Sheet named `bold` with specimen metadata
    - Columns:
        - `MUSEUMID`
        - `SAMPLEID`
