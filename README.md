# Reservation Data Processor

A complete VBA macro solution for processing hotel reservation data from CSV files into formatted Excel reports with automatic TDF (Tourism Dirham Fee) calculations.

## 🎯 Overview

This project automates the processing of reservation data from a delimited data source into a standardized "Entered On" report format. It includes duplicate detection, automatic calculations, spillover row handling, and dynamic formula generation.

## ✨ Features

- ✅ **Automated Data Processing**: Reads from `DELIMITED DATA` sheet, writes to `ENTERED ON` sheet
- ✅ **Duplicate Prevention**: Checks existing RESV_ID records before adding
- ✅ **TDF Calculation**:
  - 1BA: 20 AED per night (capped at 600 AED for 30+ nights)
  - 2BA: 40 AED per night (capped at 1200 AED for 30+ nights)
- ✅ **Spillover Row Handling**: Automatically detects and merges rows with "T-", "S-", "C-" patterns
- ✅ **Name Splitting**: Separates full names into Last Name and First Name
- ✅ **Date Format Preservation**: Maintains dd/mm/yyyy format
- ✅ **ADR Calculation**: Automatic Average Daily Rate calculation
- ✅ **Dynamic Formulas**: Adds Season, Booking Lead Time, and Events Dates formulas
- ✅ **Performance Optimized**: Fast processing with screen updating and calculation optimization

## 📋 Prerequisites

- Microsoft Excel 2010 or later (with macro support)
- Windows operating system
- Macro security set to allow VBA execution

## 🚀 Quick Start

### Installation

**Option A: Automatic (Recommended)**
1. Double-click `InstallMacro_Updated.vbs`
2. Follow the prompts

**Option B: Manual**
1. Open your Excel workbook (`22-11-2025 Entered On.xlsm`)
2. Press `ALT + F11` to open VBA Editor
3. Go to File → Import File
4. Select `ReservationProcessor_Final.vba`
5. Close VBA Editor

### Usage

1. **Import CSV Data**
   - Open the Excel workbook
   - Navigate to `DELIMITED DATA` sheet
   - Data → From Text/CSV
   - Select your CSV file (Tab-delimited, UTF-8)
   - Load the data

2. **Run the Macro**
   - Press `ALT + F8`
   - Select `ProcessReservations`
   - Click "Run"
   - Wait for completion message

3. **Review Results**
   - Check the `ENTERED ON` sheet
   - Verify calculations and formulas

## 📊 Column Mapping

| Source (DELIMITED DATA) | Target (ENTERED ON) | Description |
|------------------------|---------------------|-------------|
| Q: FULL_NAME | A: Last Name + B: First Name | Name splitting |
| AC: ARRIVAL | C: ARRIVAL | Date format: dd/mm/yyyy |
| R: DEPARTURE | D: DEPARTURE | Date format: dd/mm/yyyy |
| AD: NIGHTS | E: NIGHTS | Number of nights |
| V: ROOM_CATEGORY | G: ROOM | Room type |
| Calculated | H: TDF | Tourism Dirham Fee |
| AI: SHARE_AMOUNT_PER_STAY | I: NET | Amount without TDF |
| NET + TDF | J: TOTAL | Total amount |
| AF: SHARE_AMOUNT | P: AMOUNT | Share amount |
| AMOUNT / NIGHTS | O: ADR | Average Daily Rate |
| M+Y: RESV_ID+DATE | S: RESV ID | Unique identifier |
| Formula | T: Season | Winter/Summer |
| Formula | U: Booking Lead Time | Days to arrival |
| Formula | V: Events Dates | 12 major UAE events |

## 🧮 TDF Calculation Examples

| Room Type | Nights | Calculation | TDF Amount |
|-----------|--------|-------------|------------|
| 1BA | 7 | 7 × 20 | 140 AED |
| 1BA | 30 | 30 × 20 | 600 AED |
| 1BA | 45 | Capped | 600 AED |
| 2BA | 10 | 10 × 40 | 400 AED |
| 2BA | 35 | Capped | 1200 AED |
| Other | Any | N/A | 0 AED |

## 📅 Events Tracked (2025)

The macro automatically identifies if arrival dates fall within these major UAE events:

- **Arab Health**: Jan 26-31
- **Gulf Food**: Feb 16-21
- **Ramadan**: Mar 1-29
- **Eid Al Fitr**: Mar 30 - Apr 2
- **Eid Al Adha**: Jun 6-9
- **GITEX**: Oct 12-17
- **Gulf Food Manufacturing**: Nov 3-7
- **Air Show**: Nov 16-21
- **Big 5**: Nov 23-28
- **National Day Holidays**: Nov 29 - Dec 2
- **F1 Yas Island**: Dec 4-7
- **New Year's Eve**: Dec 26-31

## 📁 Project Structure

```
├── ReservationProcessor_Final.vba    # Main macro code (USE THIS)
├── InstallMacro_Updated.vbs          # Automatic installer
├── START_HERE.txt                     # Quick start guide
├── QUICK_REFERENCE.txt                # Quick reference cheat sheet
├── INSTALLATION_GUIDE.txt             # Complete installation guide
├── DATA_FLOW_DIAGRAM.txt              # Visual data flow diagrams
├── README_FINAL.txt                   # Complete documentation
├── README.md                          # This file
└── .gitignore                         # Git ignore rules
```

### Legacy Files (Not Needed)
- `ProcessReservations.vba` - Old version
- `InstallMacro.vbs` - Old installer
- `README_Instructions.txt` - Old documentation
- `PROJECT_SUMMARY.md` - Old summary

## 🔧 Troubleshooting

### Common Issues

**Problem: "DELIMITED DATA sheet not found"**
- **Solution**: Ensure sheet name is exactly "DELIMITED DATA" (all caps)

**Problem: "ENTERED ON sheet not found"**
- **Solution**: Ensure sheet name is exactly "ENTERED ON" (all caps)

**Problem: No records processed**
- **Solution**: Check that DELIMITED DATA has data from row 2 onwards

**Problem: All records skipped as duplicates**
- **Solution**: RESV IDs already exist. Clear ENTERED ON sheet (keep headers) to reprocess

**Problem: TDF always 0**
- **Solution**: Room category must be exactly "1BA" or "2BA" (case-sensitive)

**Problem: Can't install with VBS script**
- **Solution**: Enable "Trust access to VBA project object model"
  - File → Options → Trust Center → Macro Settings
  - Or use manual installation method

## ⚡ Performance

Expected processing times:
- ~50 records: 2-5 seconds
- ~200 records: 10-15 seconds
- ~500 records: 30-45 seconds

The macro uses several optimizations:
- Screen updating disabled during processing
- Manual calculation mode during processing
- Dictionary object for O(1) duplicate lookups
- Bottom-up processing for spillover fixing

## 📖 Documentation

For detailed information, refer to:

- **START_HERE.txt** - Begin here for quick overview
- **QUICK_REFERENCE.txt** - 5-minute quick guide
- **INSTALLATION_GUIDE.txt** - Complete step-by-step guide
- **DATA_FLOW_DIAGRAM.txt** - Visual process flows
- **README_FINAL.txt** - Comprehensive documentation

## 🧪 Testing

Before processing production data:

1. ✅ Backup your Excel file
2. ✅ Clear ENTERED ON sheet (keep headers)
3. ✅ Add 5-10 test rows to DELIMITED DATA
4. ✅ Run the macro
5. ✅ Verify results:
   - Names split correctly
   - Dates in dd/mm/yyyy format
   - TDF calculated correctly
   - NET + TDF = TOTAL
   - Formulas working in columns T, U, V

## 🛡️ Data Privacy

This project processes hotel reservation data. Ensure compliance with:
- Data protection regulations (GDPR, local laws)
- Company data handling policies
- Guest privacy requirements

**Note**: The `.gitignore` file excludes actual data files (`.xlsm`, `.csv`) from version control.

## 📝 Version History

### Version 1.0 (Current)
- Initial release
- Complete data processing functionality
- TDF calculation with proper rules
- Duplicate detection
- Spillover row handling
- Name splitting
- Date format preservation
- ADR calculation
- Season, Lead Time, and Events formulas
- Performance optimizations

## 🤝 Contributing

To contribute or modify:

1. Clone this repository
2. Make changes to the VBA code
3. Test thoroughly with sample data
4. Update documentation as needed
5. Submit pull request with description

## 📄 License

This project is provided as-is for internal use. Modify as needed for your requirements.

## 🆘 Support

For help:

1. Read the documentation files (especially QUICK_REFERENCE.txt)
2. Check the TROUBLESHOOTING sections
3. Review the DATA_FLOW_DIAGRAM.txt for understanding
4. Verify sheet names and column headers match expected format
5. Test with a small dataset first

## 🔑 Key Functions

### Main Functions in VBA Code

- **`ProcessReservations()`** - Main entry point
- **`CalculateTDF()`** - Tourism Dirham Fee calculation
- **`FixSpilloverRows()`** - Handles data spillover
- **`SplitName()`** - Splits full names
- **`AddSeasonFormula()`** - Adds season formula
- **`AddLeadTimeFormula()`** - Adds lead time formula
- **`AddEventsFormula()`** - Adds events formula

## 🎓 Learning Resources

To understand the code:
1. Read the inline comments in `ReservationProcessor_Final.vba`
2. Review `DATA_FLOW_DIAGRAM.txt` for visual understanding
3. Check `INSTALLATION_GUIDE.txt` for detailed explanations

---

**Created**: 2025-11-24
**Status**: Production Ready
**Macro File**: ReservationProcessor_Final.vba

---

Made with ❤️ for efficient reservation data processing
