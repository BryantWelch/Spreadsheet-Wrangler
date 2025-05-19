# Spreadsheet Wrangler

A PowerShell GUI application for folder backups, spreadsheet combining operations, and SKU list processing.

<img src="https://github.com/user-attachments/assets/d3791579-ae79-45d9-866e-95077f9baccc" width=75% height=75%>


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
- Microsoft Excel (for spreadsheet operations)

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
