# Spreadsheet Wrangler

[![Version](https://img.shields.io/badge/version-1.6.0-blue.svg)](https://github.com/BryantWelch/Spreadsheet-Wrangler/releases/tag/v1.6.0)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1+-5391FE.svg)](https://github.com/PowerShell/PowerShell)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Platform](https://img.shields.io/badge/platform-Windows-lightgrey.svg)](https://www.microsoft.com/en-us/windows)

> A feature packed PowerShell GUI tool for spreadsheet operations with no Excel dependency

<img src="https://github.com/user-attachments/assets/f0af31ce-4c7d-42b9-b361-fd763294b308" width=75% height=75%>

## Features

### Backup Functionality
- Create timestamped backups of selected folders
- Support for multiple backup locations
- Automatic ".backup" folder creation
- Option to skip backup process

### Spreadsheet Combining
- Combine spreadsheets with similar numbering across folders
- Support for multiple spreadsheet formats (.xlsx, .xls, .csv)
- Maintain headers from the first spreadsheet (optional)
- Save combined spreadsheets to a user-selected destination

### SKU List Processing
- Match TCGplayer IDs from combined spreadsheets with TIDs in a SKU list CSV file
- Create GSxx spreadsheets with matched data from the SKU list
- Generate barcodes based on SKU and price information
- Extract specific fields like SKU, Card Name, Condition, Cost, and Price
- Create GS_Missing spreadsheet with all unmatched rows for review
- Optimized performance with hashtable-based lookups (4-5x faster)
- Process spreadsheets in natural numerical order

### Label Creation
- Create printer-ready label files from Excel spreadsheets
- Support for both standard printer labels (.tsk files) and Dymo labels (.dymo files)
- Customizable templates for different label formats
- Process all GSxx.xlsx files in a folder automatically
- Proper handling of special characters in card names and other data
- User-friendly dialog with tooltips for all options

### Advanced Spreadsheet Options
- **No Headers**: Exclude headers when combining spreadsheets
- **Duplicate Qty=2**: Duplicate rows with '2' in the "Add to Quantity" column
- **Normalize Qty to 1**: Change all values in "Add to Quantity" column to '1'
- **Log to File**: Save terminal output to a log file for future reference
- **BLANK**: Insert a row with "BLANK" text between data from different spreadsheets
- **Reverse, Reverse**: Reverse the order of data rows in the final combined spreadsheet

### Configuration Management
- Save/load application settings to/from XML files
- Menu system with keyboard shortcuts
- Persistent settings across sessions

## Requirements

- Windows operating system
- PowerShell 5.1 or higher
- ImportExcel PowerShell module (automatically installed if missing)

## Installation

1. Click the green code button at the top of this repo and select download zip
2. Alternatively, you can clone this repository or download the latest release
3. Extract the files to your preferred location

## Running the Application

### Option 1: No Console Window (Recommended)
Double-click on `Launch-SpreadsheetWrangler.vbs` to run the application without showing a PowerShell console window.

### Option 2: Right-click method
Right-click on `SpreadsheetWrangler.ps1` and select "Run with PowerShell"
(Note: This will show a PowerShell console window alongside the application)

### Option 3: Command line
```powershell
powershell -ExecutionPolicy Bypass -File .\SpreadsheetWrangler.ps1
```
## Usage

### Backup Process
1. Add folder locations to back up using the "+" button
2. Select "Skip Backup" option if you only want to combine spreadsheets

### Spreadsheet Combining
1. Add folder locations containing spreadsheets using the "+" button
2. Set the destination folder for combined spreadsheets
3. Select desired options for the combining process
4. Click "Run" to start the process

### SKU List Processing
1. Complete the spreadsheet combining process steps above
2. Select a SKU list CSV file using the "SKU List Location" section
3. Choose a destination folder for the final output files using the "Final Output Location" section
4. Click "Run" to process the combined spreadsheets against the SKU list
5. GSxx files will be created in the final output location with matched data
6. GS_Missing will be created in the final output location with non-matched data

### Label Creation
1. Click on "Labels" in the menu bar, then select "Create Labels"
2. In the dialog that appears:
   - Select the input folder containing your GSxx.xlsx files
   - Choose an output folder where label files will be saved
   - (Optional) Select a .param template file for printer configuration
   - (Optional) Select a .prt template file containing label layout with data markers
   - (Optional) Select a Dymo template file for Dymo label creation
3. Click "Create Labels" to start the process
4. The application will create:
   - GSxx.tsk files (if param/prt templates are provided)
   - GSxx.dymo files (if a Dymo template is provided)
5. All files will be saved to the selected output folder

### Configuration
- **File → New Configuration**: Reset all settings to default
- **File → Open Configuration**: Load settings from an XML file
- **File → Save Configuration**: Save settings to the current file
- **File → Save Configuration As**: Save settings to a new file

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add some amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## Author

Created by Bryant Welch

If you find this project helpful, consider supporting its development:

[![Buy Me a Coffee](https://storage.ko-fi.com/cdn/kofi5.png)](https://ko-fi.com/V7V01A0SJC)
