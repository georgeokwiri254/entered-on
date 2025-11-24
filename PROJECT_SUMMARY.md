# Reservation Data Processing - Project Summary

## Overview
This project processes hotel reservation data from a CSV export into a formatted Excel workbook with Tourism Dirham Fee (TDF) calculations.

## Files Created
1. **ProcessReservations.vba** - Main VBA macro code
2. **InstallMacro.vbs** - Automated installer script
3. **README_Instructions.txt** - Detailed usage instructions
4. **PROJECT_SUMMARY.md** - This file

## Data Analysis Results

### CSV Structure
- **Format**: Tab-delimited (TSV) file
- **Columns**: 35 columns total
- **Records**: ~10,000+ reservation records
- **Encoding**: UTF-8 with international characters

### Key Columns Identified
| Column | Name | Description |
|--------|------|-------------|
| 1 | RESORT | Arrival/Entry date |
| 2 | GRPBY_DISP1 | Display date |
| 13 | RESV_NAME_ID | Reservation ID |
| 16 | ROOM | Room number |
| 17 | FULL_NAME | Guest name |
| 22 | ROOM_CATEGORY_LABEL | Room type (1BA, 2BA, etc.) |
| 29 | ARRIVAL | Check-in date |
| 30 | NIGHTS | Number of nights |
| 32 | SHARE_AMOUNT | Share amount (→ Output column P) |
| 33 | C_T_S_NAME | Company/Travel/Source name |
| 34 | SHORT_RESV_STATUS | Status code |
| 35 | SHARE_AMOUNT_PER_STAY | Total amount (→ Output column AI/Share) |

### Spillover Row Problem
**Pattern Detected:**
- Some rows have data that spills into the next row
- Spillover rows start with: `[Letter]- [Text]` (e.g., "T- TA Connections", "S- VCC", "C- ASL Airlines")
- Common prefixes: T-, S-, C-
- These represent continuation of the Travel Agent/Source field

**Examples Found:**
```
Line 182: ...	ASL Airlines	TA Connections	24.11.25	...
Line 183: T- TA Connections	CKIN	800
```

**Solution Implemented:**
- `FixSpilloverRows()` function processes from bottom to top
- Merges spillover row Column A with previous row Column 28 (Travel Agent)
- Moves spillover Columns B and C to previous row Columns 34 and 35
- Deletes the spillover row

## TDF Calculation Logic

### Rules
1. **1BA Rooms (1 Bedroom Apartment)**
   - Rate: 20 AED per room per night
   - For stays ≤ 30 nights: `TDF = Nights × 20`
   - For stays > 30 nights: `TDF = 600 AED` (capped)

2. **2BA Rooms (2 Bedroom Apartment)**
   - Rate: 40 AED per room per night
   - For stays ≤ 30 nights: `TDF = Nights × 40`
   - For stays > 30 nights: `TDF = 1200 AED` (capped)

3. **Other Room Types**
   - TDF = 0 AED

### Examples
| Room Type | Nights | TDF Calculation | TDF Amount |
|-----------|--------|-----------------|------------|
| 1BA | 7 | 7 × 20 | 140 AED |
| 1BA | 30 | 30 × 20 | 600 AED |
| 1BA | 45 | Capped | 600 AED |
| 2BA | 10 | 10 × 40 | 400 AED |
| 2BA | 30 | 30 × 40 | 1200 AED |
| 2BA | 60 | Capped | 1200 AED |
| ST | 10 | N/A | 0 AED |

## Output Sheet Columns

### Column Mapping
| Output Col | Column Name | Source | Calculation/Notes |
|------------|-------------|--------|-------------------|
| A | ARRIVAL | CSV Col 1 | Entry date |
| B | DATE | CSV Col 2 | Display date |
| C | RESV_NAME_ID | CSV Col 13 | Reservation ID |
| D | GUARANTEE_CODE | CSV Col 14 | Guarantee code |
| E | RESV_STATUS | CSV Col 15 | Reservation status |
| F | ROOM | CSV Col 16 | Room number |
| G | FULL_NAME | CSV Col 17 | Guest name |
| H | DEPARTURE | CSV Col 18 | Checkout date |
| I | NET | CSV Col 35 | Amount WITHOUT TDF |
| J | TOTAL | Calculated | Amount WITH TDF (I + Y) |
| K | PERSONS | CSV Col 19 | Number of persons |
| L | GROUP_NAME | CSV Col 20 | Group name |
| M | NO_OF_ROOMS | CSV Col 21 | Number of rooms |
| N | ROOM_CATEGORY | CSV Col 22 | Room type |
| O | RATE_CODE | CSV Col 23 | Rate code |
| P | SHARE | CSV Col 32 | Share amount |
| Q | INSERT_USER | CSV Col 24 | User who created |
| R | INSERT_DATE | CSV Col 25 | Creation date |
| S | GUARANTEE_DESC | CSV Col 26 | Guarantee description |
| T | COMPANY_NAME | CSV Col 27 | Company name |
| U | TRAVEL_AGENT | CSV Col 28 + 33 | Travel agent (merged) |
| V | ARRIVAL_DATE | CSV Col 29 | Arrival date |
| W | NIGHTS | CSV Col 30 | Number of nights |
| X | COMP_HOUSE | CSV Col 31 | Complimentary flag |
| Y | TDF | Calculated | Tourism Dirham Fee |
| Z | ADR | Calculated | P ÷ W (Average Daily Rate) |
| AA | SOURCE | CSV Col 33 | Source name |
| AB | STATUS | CSV Col 34 | Short status |

## Macro Functions

### Main Functions

#### `ProcessEnteredOnReport()`
- **Purpose**: Main processing routine
- **Actions**:
  1. Opens CSV file
  2. Fixes spillover rows
  3. Creates/clears "Entered On" sheet
  4. Processes each reservation record
  5. Calculates TDF and derived fields
  6. Formats output sheet
  7. Reports completion

#### `CalculateTDF(roomCategory, nights)`
- **Purpose**: Calculate Tourism Dirham Fee
- **Parameters**:
  - `roomCategory`: String (e.g., "1BA", "2BA")
  - `nights`: Long (number of nights)
- **Returns**: Double (TDF amount in AED)
- **Logic**: Implements the TDF rules above

#### `FixSpilloverRows(ws)`
- **Purpose**: Merge spillover rows with parent rows
- **Parameters**: `ws` - Worksheet object
- **Algorithm**:
  1. Loop from last row to first (bottom-up)
  2. Check if row starts with `[Letter]-` pattern
  3. Merge data with previous row
  4. Delete spillover row

#### `SetupHeaders(ws)`
- **Purpose**: Create column headers
- **Parameters**: `ws` - Worksheet object
- **Actions**: Sets 28 column headers with formatting

#### `FormatSheet(ws, lastRow)`
- **Purpose**: Apply formatting to output
- **Parameters**:
  - `ws` - Worksheet object
  - `lastRow` - Last row with data
- **Formatting**:
  - Auto-fit columns
  - Number format for currency columns
  - Borders on all cells
  - Freeze header row
  - Bold header with gray background

## Installation Methods

### Method 1: Automatic Installation (Recommended)
1. Double-click `InstallMacro.vbs`
2. Script automatically:
   - Reads VBA code
   - Opens Excel file
   - Adds macro module
   - Saves and closes

**Requirements**:
- Enable "Trust access to the VBA project object model"
- Path: Excel Options → Trust Center → Macro Settings

### Method 2: Manual Import
1. Open Excel file
2. Press `ALT + F11`
3. File → Import File
4. Select `ProcessReservations.vba`
5. Close VBA Editor

### Method 3: Copy-Paste
1. Open `ProcessReservations.vba` in text editor
2. Copy all code
3. Open Excel file
4. Press `ALT + F11`
5. Insert → Module
6. Paste code
7. Close VBA Editor

## Usage Instructions

### Prerequisites
- Excel 2010 or later (with macro support)
- Both files in same folder:
  - `22-11-2025 Entered On.xlsm`
  - `resenteredon102243710-lpo.csv`
- Macros enabled

### Running the Macro
1. Open `22-11-2025 Entered On.xlsm`
2. Enable macros when prompted
3. Press `ALT + F8`
4. Select `ProcessEnteredOnReport`
5. Click "Run"
6. Wait for completion message
7. Review "Entered On" sheet

### Expected Results
- New/refreshed "Entered On" sheet
- ~10,000+ records processed
- All calculations completed
- Professional formatting applied
- Processing time: ~30-60 seconds (depends on system)

## Validation Checklist

After running the macro, verify:

1. **Row Count**
   - [ ] Total rows matches expected count (CSV rows - spillover rows)

2. **TDF Calculations**
   - [ ] 1BA with ≤30 nights: TDF = Nights × 20
   - [ ] 1BA with >30 nights: TDF = 600
   - [ ] 2BA with ≤30 nights: TDF = Nights × 40
   - [ ] 2BA with >30 nights: TDF = 1200

3. **Column Mappings**
   - [ ] Column P (SHARE) = CSV Column 32
   - [ ] Column AI equivalent (would be around column 35 in manual entry) = CSV Column 35
   - [ ] NET (Column I) = Amount without TDF
   - [ ] TOTAL (Column J) = NET + TDF

4. **ADR Calculation**
   - [ ] ADR = SHARE ÷ NIGHTS
   - [ ] Check random samples

5. **Spillover Rows**
   - [ ] No rows starting with "T- ", "S- ", "C- " in Column A
   - [ ] Travel agent names properly concatenated

6. **Data Integrity**
   - [ ] No blank rows in middle of data
   - [ ] All dates properly formatted
   - [ ] Guest names intact
   - [ ] Room numbers preserved

## Troubleshooting Guide

### Common Errors

#### "File not found"
**Cause**: CSV file not in same folder as Excel file
**Solution**: Move CSV to same folder or update path in code line:
```vba
csvPath = wb.Path & "\resenteredon102243710-lpo.csv"
```

#### "Macro not found"
**Cause**: Macro not properly installed
**Solution**: Re-import using Method 2 or 3

#### "Permission denied" / "Access denied"
**Cause**: File is open elsewhere or read-only
**Solution**: Close all Excel instances, remove read-only attribute

#### "Trust access to VBA object model"
**Cause**: Security setting preventing VBA automation
**Solution**:
1. Excel Options
2. Trust Center
3. Trust Center Settings
4. Macro Settings
5. Enable "Trust access to the VBA project object model"

#### Wrong TDF amounts
**Cause**: Room categories don't match "1BA" or "2BA" exactly
**Solution**: Check CSV for variations (spaces, lowercase, etc.)

#### Spillover rows still appearing
**Cause**: Different spillover pattern than expected
**Solution**:
1. Open CSV in text editor
2. Find spillover examples
3. Update pattern in `FixSpilloverRows()` function

### Performance Issues

If processing is slow:
- Disable screen updating: Add `Application.ScreenUpdating = False` at start
- Disable automatic calculation: Add `Application.Calculation = xlCalculationManual`
- Re-enable at end of processing

## Technical Notes

### Character Encoding
- CSV uses UTF-8 encoding
- Contains international characters (Arabic, Cyrillic, Chinese, etc.)
- Excel handles automatically with proper import

### Data Types
- Dates: Various formats (DD/MM/YYYY, DD.MM.YY)
- Numbers: Decimal separator is period (.)
- Currency: AED (no symbol in data)

### Memory Considerations
- Large file (~10K rows, 35 columns)
- Use array processing if experiencing memory issues
- Current implementation uses cell-by-cell for clarity

## Future Enhancements

Potential improvements:
1. **Array Processing**: Load data to array for faster processing
2. **Error Logging**: Create error log sheet for problematic records
3. **Progress Bar**: Show processing progress
4. **Validation Report**: Automatic data validation summary
5. **Configurable TDF**: Make TDF rates configurable
6. **Multiple CSV Support**: Process multiple CSV files in batch
7. **Data Comparison**: Compare with previous reports
8. **Export Options**: Export to PDF or other formats

## Version History

**Version 1.0** (Current)
- Initial release
- Core processing functionality
- TDF calculation
- Spillover row fixing
- Basic formatting

## Contact & Support

For issues or questions:
1. Review this documentation
2. Check README_Instructions.txt
3. Verify CSV format matches expected structure
4. Test with small subset of data first

---

**Generated**: 2025-11-24
**Status**: Ready for testing
**Files**: 4 files created in project folder
