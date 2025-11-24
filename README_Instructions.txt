================================================================
RESERVATION DATA PROCESSING MACRO - INSTALLATION & USAGE GUIDE
================================================================

OVERVIEW:
---------
This VBA macro processes the CSV file "resenteredon102243710-lpo.csv" and creates
a formatted "Entered On" sheet in the Excel workbook with TDF calculations.

TDF CALCULATION LOGIC:
----------------------
- 1BA Room: 20 AED per room per night
- 2BA Room: 40 AED per room per night
- Reservations > 30 nights:
  * 1BA: Fixed 600 AED (after 30 nights, no additional TDF)
  * 2BA: Fixed 1200 AED (after 30 nights, no additional TDF)

DATA MAPPING:
-------------
- Share column (AI in output) = SHARE_AMOUNT_PER_STAY from CSV column 35
- ADR = Amount ÷ Number of Nights
- NET (Column I) = Amount without TDF
- TOTAL (Column J) = Amount with TDF
- Column P in output = SHARE_AMOUNT from CSV column 32

SPILLOVER ROW HANDLING:
-----------------------
The macro automatically fixes rows that spill over to the next line.
Pattern identified: Rows starting with "T- ", "S- ", "C- " etc.
These are merged with the previous row's Travel Agent field (column 28).

INSTALLATION INSTRUCTIONS:
--------------------------

METHOD 1: Direct Import into Excel
1. Open "22-11-2025 Entered On.xlsm"
2. Press ALT + F11 to open VBA Editor
3. Go to File > Import File
4. Select "ProcessReservations.vba"
5. Close VBA Editor (ALT + Q)

METHOD 2: Copy-Paste Method
1. Open "ProcessReservations.vba" in Notepad
2. Copy all the code (CTRL + A, then CTRL + C)
3. Open "22-11-2025 Entered On.xlsm"
4. Press ALT + F11 to open VBA Editor
5. In VBA Editor, go to Insert > Module
6. Paste the code (CTRL + V)
7. Close VBA Editor (ALT + Q)

USAGE:
------
1. Ensure both files are in the same folder:
   - 22-11-2025 Entered On.xlsm
   - resenteredon102243710-lpo.csv

2. Open "22-11-2025 Entered On.xlsm"

3. Enable Macros if prompted

4. Run the macro:
   - Press ALT + F8
   - Select "ProcessEnteredOnReport"
   - Click "Run"

5. Wait for processing to complete
   - A message box will show the number of records processed

6. Review the "Entered On" sheet
   - Data will be formatted with borders and frozen headers
   - Number columns will have proper formatting

IMPORTANT NOTES:
----------------
1. The macro will clear existing data in "Entered On" sheet before processing
2. CSV file must be in the same folder as the Excel file
3. Backup your data before running the macro
4. The macro processes from the CSV file each time it runs

TROUBLESHOOTING:
----------------
Error: "File not found"
  → Ensure resenteredon102243710-lpo.csv is in the same folder as the Excel file

Error: "Macro not found"
  → Verify the macro was imported correctly (ALT + F11 to check)

Spillover rows still appearing:
  → Check that rows match the pattern "[Letter]- [Text]" in column A
  → Manually review the CSV for unusual formatting

Wrong TDF calculations:
  → Verify room categories are exactly "1BA" or "2BA"
  → Check that nights column (30) has valid numbers

COLUMN LAYOUT IN OUTPUT:
------------------------
A: ARRIVAL              N: ROOM_CATEGORY
B: DATE                 O: RATE_CODE
C: RESV_NAME_ID         P: SHARE (from col 32)
D: GUARANTEE_CODE       Q: INSERT_USER
E: RESV_STATUS          R: INSERT_DATE
F: ROOM                 S: GUARANTEE_DESC
G: FULL_NAME            T: COMPANY_NAME
H: DEPARTURE            U: TRAVEL_AGENT
I: NET (no TDF)         V: ARRIVAL_DATE
J: TOTAL (with TDF)     W: NIGHTS
K: PERSONS              X: COMP_HOUSE
L: GROUP_NAME           Y: TDF
M: NO_OF_ROOMS          Z: ADR
                        AA: SOURCE
                        AB: STATUS

CONTACT & SUPPORT:
------------------
If you encounter issues:
1. Check that the CSV file format matches the expected structure
2. Verify Excel macro security settings allow macros to run
3. Review the VBA code comments for additional details

================================================================
