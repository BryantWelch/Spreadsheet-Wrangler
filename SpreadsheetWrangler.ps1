# SpreadsheetWrangler.ps1
# GUI application for spreadsheet operations and folder backups

# Check for ImportExcel module availability
$script:UseImportExcel = $false

# Check if ImportExcel module is installed
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "ImportExcel module not found. Attempting to install..." -ForegroundColor Yellow
    try {
        Install-Module -Name ImportExcel -Scope CurrentUser -Force -ErrorAction Stop
        Write-Host "ImportExcel module installed successfully." -ForegroundColor Green
    } catch {
        $errorMessage = $_.Exception.Message
        Write-Host "Failed to install ImportExcel module. Please run 'Install-Module -Name ImportExcel -Scope CurrentUser -Force' manually." -ForegroundColor Red
        Write-Host "Error``: $errorMessage" -ForegroundColor Red
        exit
    }
}

# Import the module
try {
    Import-Module -Name ImportExcel -ErrorAction Stop
    $script:UseImportExcel = $true
    Write-Host "ImportExcel module loaded successfully." -ForegroundColor Green
} catch {
    $errorMessage = $_.Exception.Message
    Write-Host "Failed to load ImportExcel module. Error``: $errorMessage" -ForegroundColor Red
    exit
}

#region Helper Functions

# Function to create a timestamp string for folder naming
function Get-TimeStampString {
    return Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
}

# Function to extract number from filename
function Get-FileNumber {
    param (
        [Parameter(Mandatory=$true)]
        [string]$FileName
    )
    
    # Remove file extension first
    $nameWithoutExtension = [System.IO.Path]::GetFileNameWithoutExtension($FileName)
    
    # Debug output
    Write-Log "  Extracting number from: $nameWithoutExtension" "Gray"
    
    # Look for patterns like "(1)", "[1]", "_1", etc.
    if ($nameWithoutExtension -match "[\(\[_\s-](\d+)[\)\]\s]*$") {
        Write-Log "    Found number: $($matches[1])" "Gray"
        return $matches[1]
    }
    # Look for patterns where the number is at the end
    elseif ($nameWithoutExtension -match "(\d+)[\)\]\s]*$") {
        Write-Log "    Found number: $($matches[1])" "Gray"
        return $matches[1]
    }
    
    # If no number found, return null
    Write-Log "    No number found" "Gray"
    return $null
}

# Global variable for log file path
$script:LogFilePath = $null

# Function to log messages to the output textbox and optionally to a file
function Write-Log {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Message,
        
        [Parameter(Mandatory=$false)]
        [string]$Color = "LightGreen"
    )
    
    # Get timestamp for log file
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    
    # Ensure we're on the UI thread for textbox updates
    if ($outputTextbox.InvokeRequired) {
        $outputTextbox.Invoke([Action[string, string]]{ param($msg, $clr) 
            $outputTextbox.SelectionColor = [System.Drawing.Color]::$clr
            $outputTextbox.AppendText("$msg`r`n")
            $outputTextbox.ScrollToCaret()
        }, $Message, $Color)
    } else {
        $outputTextbox.SelectionColor = [System.Drawing.Color]::$Color
        $outputTextbox.AppendText("$Message`r`n")
        $outputTextbox.ScrollToCaret()
    }
    
    # If logging to file is enabled, write to the log file
    if ($script:LogFilePath -and (Test-Path $script:LogFilePath)) {
        "[$timestamp] $Message" | Out-File -FilePath $script:LogFilePath -Append
    }
}

# Function to update the progress bar
function Update-ProgressBar {
    param (
        [Parameter(Mandatory=$true)]
        [int]$Value
    )
    
    # Ensure we're on the UI thread
    if ($progressBar.InvokeRequired) {
        $progressBar.Invoke([Action[int]]{ param($val) 
            $progressBar.Value = $val
        }, $Value)
    } else {
        $progressBar.Value = $Value
    }
}

# Function to create backup of a folder
function Backup-Folder {
    param (
        [Parameter(Mandatory=$true)]
        [string]$SourcePath
    )
    
    try {
        # Create backup directory if it doesn't exist
        $backupRootDir = Join-Path -Path $PSScriptRoot -ChildPath ".backup"
        if (-not (Test-Path -Path $backupRootDir)) {
            New-Item -Path $backupRootDir -ItemType Directory -Force | Out-Null
            Write-Log "Created backup directory: $backupRootDir"
        }
        
        # Get folder name from source path
        $folderName = Split-Path -Path $SourcePath -Leaf
        
        # Create timestamped backup folder
        $timestamp = Get-TimeStampString
        $backupFolderName = "$folderName-$timestamp"
        $backupPath = Join-Path -Path $backupRootDir -ChildPath $backupFolderName
        
        # Create the backup folder
        New-Item -Path $backupPath -ItemType Directory -Force | Out-Null
        Write-Log "Created backup folder: $backupPath"
        
        # Copy all items from source to backup
        Write-Log "Starting backup of $SourcePath..."
        Copy-Item -Path "$SourcePath\*" -Destination $backupPath -Recurse -Force
        Write-Log "Backup completed successfully!" "Yellow"
        
        return $true
    }
    catch {
        Write-Log "Error during backup: $_" "Red"
        return $false
    }
}

# Function to perform backup of all selected folders
function Start-BackupProcess {
    if ($backupLocations.Items.Count -eq 0) {
        Write-Log "No backup locations selected." "Yellow"
        return
    }
    
    Write-Log "Starting backup process..." "Cyan"
    Update-ProgressBar 0
    
    $totalFolders = $backupLocations.Items.Count
    $completedFolders = 0
    
    foreach ($item in $backupLocations.Items) {
        $folderPath = $item.Text
        Write-Log "Processing backup for: $folderPath" "White"
        
        $success = Backup-Folder -SourcePath $folderPath
        $completedFolders++
        
        # Update progress
        $progressPercentage = [int](($completedFolders / $totalFolders) * 100)
        Update-ProgressBar $progressPercentage
    }
    
    Write-Log "Backup process completed." "Cyan"
    Update-ProgressBar 100
}

# Function to combine spreadsheets
function Combine-Spreadsheets {
    param (
        [Parameter(Mandatory=$true)]
        [System.Collections.ArrayList]$FolderPaths,
        
        [Parameter(Mandatory=$true)]
        [string]$DestinationPath,
        
        [Parameter(Mandatory=$false)]
        [string]$FileExtension = "*.xlsx",
        
        [Parameter(Mandatory=$false)]
        [bool]$ExcludeHeaders = $false,
        
        [Parameter(Mandatory=$false)]
        [bool]$DuplicateQuantityTwoRows = $false,
        
        [Parameter(Mandatory=$false)]
        [bool]$NormalizeQuantities = $false,
        
        [Parameter(Mandatory=$false)]
        [bool]$InsertBlankRows = $false,
        
        [Parameter(Mandatory=$false)]
        [bool]$ReverseDataRows = $false
    )
    
    try {
        # Using ImportExcel module for all operations
        Write-Log "Using ImportExcel module for all operations..." "White"
        
        # Create a dictionary to store spreadsheets by number
        $spreadsheetGroups = @{}
        
        # First, scan all folders and group spreadsheets by their number
        Write-Log "Scanning folders for spreadsheets..." "White"
        
        foreach ($folderPath in $FolderPaths) {
            Write-Log "Scanning folder: $folderPath" "White"
            
            # Handle different file formats if All Formats option is selected
            if ($FileExtension -eq "*.*") {
                # Use a hash table to track unique base filenames to prevent duplicates
                $uniqueFiles = @{}
                
                # Get all spreadsheet files
                $allFiles = @()
                $allFiles += Get-ChildItem -Path $folderPath -Filter "*.xlsx" -File
                $allFiles += Get-ChildItem -Path $folderPath -Filter "*.xls" -File
                $allFiles += Get-ChildItem -Path $folderPath -Filter "*.csv" -File
                
                $files = @()
                foreach ($file in $allFiles) {
                    $baseFileName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
                    if (-not $uniqueFiles.ContainsKey($baseFileName)) {
                        $uniqueFiles[$baseFileName] = $true
                        $files += $file
                    }
                }
                
                Write-Log "  Found $($files.Count) unique spreadsheet files" "White"
            } else {
                $files = Get-ChildItem -Path $folderPath -Filter $FileExtension -File
            }
            
            foreach ($file in $files) {
                $fileNumber = Get-FileNumber -FileName $file.Name
                
                if ($fileNumber) {
                    if (-not $spreadsheetGroups.ContainsKey($fileNumber)) {
                        $spreadsheetGroups[$fileNumber] = New-Object System.Collections.ArrayList
                    }
                    
                    $null = $spreadsheetGroups[$fileNumber].Add($file.FullName)
                    Write-Log "  Found spreadsheet: $($file.Name) (Group: $fileNumber)" "White"
                }
            }
        }
        
        # Check if we found any spreadsheets
        if ($spreadsheetGroups.Count -eq 0) {
            Write-Log "No spreadsheets with matching numbers found in the selected folders." "Yellow"
            return $false
        }
        
        # Create destination directory if it doesn't exist
        if (-not (Test-Path -Path $DestinationPath)) {
            New-Item -Path $DestinationPath -ItemType Directory -Force | Out-Null
            Write-Log "Created destination directory: $DestinationPath" "White"
        }
        
        # Process each group of spreadsheets
        $totalGroups = $spreadsheetGroups.Count
        $completedGroups = 0
        
        foreach ($groupNumber in $spreadsheetGroups.Keys) {
            $files = $spreadsheetGroups[$groupNumber]
            
            if ($files.Count -lt 2) {
                Write-Log "Skipping group $groupNumber - only one spreadsheet found." "Yellow"
                continue
            }
            
            Write-Log "Processing spreadsheet group $groupNumber..." "Cyan"
            
            # Variables to track data across both approaches
            $combinedData = @()
            $headers = @()
            $rowIndex = 1
            $isFirstFile = $true
            $lastColumn = 0
            
            # Process using ImportExcel
                #region ImportExcel Processing
                # Process each file in the group
                foreach ($file in $files) {
                    Write-Log "  Combining: $file" "White"
                    
                    try {
                        # Import the spreadsheet
                        $importParams = @{
                            Path = $file
                            ErrorAction = 'Stop'
                        }
                        
                        # Handle headers based on the ExcludeHeaders option
                        if ($ExcludeHeaders) {
                            # No headers mode - don't treat first row as headers
                            $importParams.Add('NoHeader', $true)
                            
                            # In no-headers mode, we need to create column names
                            if ($isFirstFile) {
                                # Initialize headers for no-headers mode
                                $isFirstFile = $false
                            }
                        } else {
                            # Use the first row as headers
                            $importParams.Add('HeaderRow', 1)
                        }
                        
                        $data = Import-Excel @importParams
                        
                        # Debug output to see what we're getting
                        Write-Log "  Imported data with $($data.Count) rows" "White"
                        if ($data.Count -gt 0) {
                            $columnCount = ($data[0].PSObject.Properties | Measure-Object).Count
                            Write-Log "  First row has $columnCount columns" "White"
                        }
                        
                        # If this is the first file and headers are included, save the headers
                        if (-not $ExcludeHeaders) {
                            if ($isFirstFile) {
                                $headers = $data[0].PSObject.Properties.Name
                                $lastColumn = $headers.Count
                                Write-Log "  Using headers: $($headers -join ', ')" "White"
                                $isFirstFile = $false
                            }
                        }
                        
                        # Add data to the combined data array, skipping headers if needed
                        if ($ExcludeHeaders) {
                            # Skip the first row of each file when No Headers is selected
                            if ($data.Count -gt 1) {
                                # Skip the first row (header) for each file
                                $combinedData += $data | Select-Object -Skip 1
                                Write-Log "  Skipping header row from file" "White"
                            } else {
                                # If there's only one row, it might be just a header - skip it entirely
                                Write-Log "  File only has one row, skipping it entirely" "White"
                            }
                        } else {
                            # Add all data when headers are included
                            $combinedData += $data
                        }
                        
                        # Insert BLANK row between spreadsheets if option is enabled and this is not the last file
                        if ($InsertBlankRows -and ($file -ne $files[-1])) {
                            Write-Log "  Inserting BLANK row after data from: $file" "White"
                            
                            # Create a blank row object
                            $blankRow = New-Object PSObject
                            
                            if (-not $ExcludeHeaders -and $headers.Count -gt 0) {
                                # Add BLANK to each column that has a header
                                foreach ($header in $headers) {
                                    $blankRow | Add-Member -MemberType NoteProperty -Name $header -Value "BLANK"
                                }
                            } else {
                                # If no headers, add BLANK to all columns in the data
                                $columnCount = if ($data.Count -gt 0) { $data[0].PSObject.Properties.Count } else { 0 }
                                for ($i = 0; $i -lt $columnCount; $i++) {
                                    $propName = if ($ExcludeHeaders) { "Column$($i+1)" } else { $headers[$i] }
                                    $blankRow | Add-Member -MemberType NoteProperty -Name $propName -Value "BLANK"
                                }
                            }
                            
                            # Add the blank row to the combined data
                            $combinedData += $blankRow
                        }
                    } catch {
                        $errorMessage = $_.Exception.Message
                        Write-Log "  Error processing file $file``: $errorMessage" "Red"
                    }
                }
            
            # Process special options and save the combined spreadsheet
            $combinedFilePath = Join-Path -Path $DestinationPath -ChildPath "Combined_Spreadsheet_$groupNumber.xlsx"
                #region ImportExcel Special Options Processing
                # Process Duplicate Qty=2 and Normalize Qty to 1 options
                if ($DuplicateQuantityTwoRows -or $NormalizeQuantities) {
                    # Find the 'Add to Quantity' column if it exists
                    $addToQuantityColName = $null
                    
                    if (-not $ExcludeHeaders -and $headers.Count -gt 0) {
                        foreach ($header in $headers) {
                            if ($header -eq "Add to Quantity") {
                                $addToQuantityColName = $header
                                Write-Log "  Found 'Add to Quantity' column" "White"
                                break
                            }
                        }
                    }
                    
                    # Process the column if found
                    if ($addToQuantityColName) {
                        # First duplicate rows with '2' in the 'Add to Quantity' column if option is enabled
                        if ($DuplicateQuantityTwoRows) {
                            Write-Log "  Processing 'Duplicate Qty=2' option..." "White"
                            
                            $rowsToAdd = @()
                            
                            # Find rows with quantity 2 and duplicate them
                            for ($i = 0; $i -lt $combinedData.Count; $i++) {
                                $row = $combinedData[$i]
                                if ($row.$addToQuantityColName -eq "2") {
                                    $rowsToAdd += $row
                                }
                            }
                            
                            # Add the duplicated rows
                            $combinedData += $rowsToAdd
                            Write-Log "  Duplicated $($rowsToAdd.Count) rows with quantity 2" "Green"
                        }
                        
                        # Then normalize all quantities to '1' if option is enabled
                        if ($NormalizeQuantities) {
                            Write-Log "  Processing 'Normalize Qty to 1' option..." "White"
                            $changedCount = 0
                            
                            foreach ($row in $combinedData) {
                                if ($row.$addToQuantityColName -ne "1") {
                                    $row.$addToQuantityColName = "1"
                                    $changedCount++
                                }
                            }
                            
                            Write-Log "  Normalized $changedCount cells to quantity 1" "Green"
                        }
                    } else {
                        Write-Log "  'Add to Quantity' column not found, skipping quantity processing" "Yellow"
                    }
                }
                
                # Apply Reverse, Reverse option if enabled
                if ($ReverseDataRows) {
                    Write-Log "  Applying 'Reverse, Reverse' option..." "White"
                    
                    # If headers are included, we need to handle them separately
                    if (-not $ExcludeHeaders -and $headers.Count -gt 0) {
                        # Reverse all rows except the header information (which is embedded in the object properties)
                        $dataCount = $combinedData.Count
                        if ($dataCount -gt 1) {
                            $reversedData = @()
                            
                            # Reverse the order of the data rows
                            for ($i = $dataCount - 1; $i -ge 0; $i--) {
                                $reversedData += $combinedData[$i]
                            }
                            
                            $combinedData = $reversedData
                            Write-Log "  Data rows reversed successfully" "Green"
                        } else {
                            Write-Log "  Not enough data rows to reverse" "Yellow"
                        }
                    } else {
                        # Reverse all rows since there are no headers
                        $combinedData = $combinedData | Sort-Object -Descending
                        Write-Log "  Data rows reversed successfully" "Green"
                    }
                }
                
                # Export the combined data to Excel
                try {
                    $exportParams = @{
                        Path = $combinedFilePath
                        WorksheetName = "Combined"
                        AutoSize = $true
                        FreezeTopRow = (-not $ExcludeHeaders)
                        TableName = "CombinedData"
                        TableStyle = "Medium2"
                        ErrorAction = "Stop"
                    }
                    
                    # Export the data
                    $combinedData | Export-Excel @exportParams
                    Write-Log "  Exported combined data to $combinedFilePath" "Green"
                } catch {
                    $errorMessage = $_.Exception.Message
                    Write-Log "  Error exporting combined data``: $errorMessage" "Red"
                }
            
            Write-Log "  Saved combined spreadsheet: $combinedFilePath" "Green"
            
            $completedGroups++
            $progressPercentage = [int](($completedGroups / $totalGroups) * 100)
            Update-ProgressBar $progressPercentage
        }
        
        # No Excel COM resources to clean up
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        
        Write-Log "Spreadsheet combining process completed." "Cyan"
        Update-ProgressBar 100
        
        return $true
    }
    catch {
        Write-Log "Error during spreadsheet combining: $_" "Red"
        
        # No Excel COM resources to clean up
        
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        
        return $false
    }
}

# Function to process SKU list and create GSxx spreadsheets
function Process-SKUList {
    param (
        [Parameter(Mandatory=$true)]
        [string]$CombinedSpreadsheetPath,
        
        [Parameter(Mandatory=$true)]
        [string]$SKUListPath,
        
        [Parameter(Mandatory=$true)]
        [string]$FinalOutputPath
    )
    
    try {
        # Check if paths exist
        if (-not (Test-Path -Path $CombinedSpreadsheetPath)) {
            Write-Log "Combined spreadsheet path does not exist: $CombinedSpreadsheetPath" "Red"
            return $false
        }
        
        if (-not (Test-Path -Path $SKUListPath)) {
            Write-Log "SKU list file does not exist: $SKUListPath" "Red"
            return $false
        }
        
        if (-not (Test-Path -Path $FinalOutputPath)) {
            # Create the final output directory if it doesn't exist
            New-Item -Path $FinalOutputPath -ItemType Directory -Force | Out-Null
            Write-Log "Created final output directory: $FinalOutputPath" "White"
        }
        
        Write-Log "Starting SKU list processing..." "Cyan"
        Update-ProgressBar 0
        
        # Import the SKU list CSV with a custom approach to handle duplicate headers
        Write-Log "Importing SKU list from: $SKUListPath" "White"
        
        # Use Import-Csv with a temporary file to handle duplicate headers
        Write-Log "Using a more robust method to import the SKU list..." "White"
        
        try {
            # Create a temporary file with only the columns we need
            $tempCsvPath = [System.IO.Path]::GetTempFileName()
            
            # Read the original CSV file
            $csvContent = Get-Content -Path $SKUListPath -Encoding UTF8
            
            # Find the indices of the columns we need
            $headerLine = $csvContent[0]
            $headerValues = $headerLine.Split(',') | ForEach-Object { $_.Trim('"').Trim() }
            
            Write-Log "Searching for required columns in headers..." "White"
            
            # Function to find column index with flexible matching
            function Find-ColumnIndex {
                param (
                    [string]$ColumnName,
                    [array]$Headers
                )
                
                # Try exact match first
                $index = $Headers.IndexOf($ColumnName)
                
                # If not found, try with trimming spaces
                if ($index -eq -1) {
                    for ($i = 0; $i -lt $Headers.Count; $i++) {
                        if ($Headers[$i].Trim() -eq $ColumnName) {
                            return $i
                        }
                    }
                    
                    # If still not found, try case-insensitive match
                    for ($i = 0; $i -lt $Headers.Count; $i++) {
                        if ($Headers[$i].Trim() -eq $ColumnName -or 
                            $Headers[$i].Trim() -eq " $ColumnName " -or
                            $Headers[$i].Trim() -eq " $ColumnName" -or
                            $Headers[$i].Trim() -eq "$ColumnName ") {
                            return $i
                        }
                    }
                    
                    # Try contains match as last resort
                    for ($i = 0; $i -lt $Headers.Count; $i++) {
                        if ($Headers[$i].Contains($ColumnName)) {
                            return $i
                        }
                    }
                }
                
                return $index
            }
            
            $tidIndex = Find-ColumnIndex -ColumnName "TID" -Headers $headerValues
            $gmeSkuIndex = Find-ColumnIndex -ColumnName "GME SKU" -Headers $headerValues
            $gmeNameIndex = Find-ColumnIndex -ColumnName "GME POS Name (36 Character Limit)" -Headers $headerValues
            $conditionIndex = Find-ColumnIndex -ColumnName "Condition Abbreviated" -Headers $headerValues
            $costIndex = Find-ColumnIndex -ColumnName "Cost" -Headers $headerValues
            $priceRoundedIndex = Find-ColumnIndex -ColumnName "Price (Rounded)" -Headers $headerValues
            $priceIndex = Find-ColumnIndex -ColumnName "Price" -Headers $headerValues
            
            Write-Log "Found column indices: TID=$tidIndex, GME SKU=$gmeSkuIndex, Name=$gmeNameIndex, Condition=$conditionIndex, Cost=$costIndex, Price Rounded=$priceRoundedIndex, Price=$priceIndex" "White"
            
            # Check if all required columns were found
            if ($tidIndex -eq -1 -or $gmeSkuIndex -eq -1 -or $gmeNameIndex -eq -1 -or 
                $conditionIndex -eq -1 -or $costIndex -eq -1 -or $priceRoundedIndex -eq -1 -or $priceIndex -eq -1) {
                Write-Log "Error: Could not find all required columns in the SKU list CSV" "Red"
                Write-Log "Required columns: TID, GME SKU, GME POS Name (36 Character Limit), Condition Abbreviated, Cost, Price (Rounded), Price" "Red"
                Write-Log "Found headers: $headerLine" "Yellow"
                return $false
            }
            
            # Write the header line to the temporary file
            "TID,GME_SKU,Card_Name,Condition,Cost,Price_Rounded,Price" | Out-File -FilePath $tempCsvPath -Encoding UTF8
            
            # Process each data row
            for ($i = 1; $i -lt $csvContent.Count; $i++) {
                $line = $csvContent[$i]
                
                # Skip empty lines
                if ([string]::IsNullOrWhiteSpace($line)) {
                    continue
                }
                
                # Parse the CSV line properly, handling quoted values
                $values = @()
                $inQuotes = $false
                $currentValue = ""
                
                for ($j = 0; $j -lt $line.Length; $j++) {
                    $char = $line[$j]
                    
                    if ($char -eq '"') {
                        $inQuotes = -not $inQuotes
                    } elseif ($char -eq ',' -and -not $inQuotes) {
                        $values += $currentValue
                        $currentValue = ""
                    } else {
                        $currentValue += $char
                    }
                }
                
                # Add the last value
                $values += $currentValue
                
                # Extract the values we need
                $tid = if ($tidIndex -lt $values.Count) { $values[$tidIndex].Trim('"') } else { "" }
                $gmeSku = if ($gmeSkuIndex -lt $values.Count) { $values[$gmeSkuIndex].Trim('"') } else { "" }
                $cardName = if ($gmeNameIndex -lt $values.Count) { $values[$gmeNameIndex].Trim('"') } else { "" }
                $condition = if ($conditionIndex -lt $values.Count) { $values[$conditionIndex].Trim('"') } else { "" }
                $cost = if ($costIndex -lt $values.Count) { $values[$costIndex].Trim('"') } else { "" }
                $priceRounded = if ($priceRoundedIndex -lt $values.Count) { $values[$priceRoundedIndex].Trim('"') } else { "" }
                $price = if ($priceIndex -lt $values.Count) { $values[$priceIndex].Trim('"') } else { "" }
                
                # Write the extracted values to the temporary file
                "$tid,$gmeSku,$cardName,$condition,$cost,$priceRounded,$price" | Out-File -FilePath $tempCsvPath -Append -Encoding UTF8
            }
            
            # Import the temporary CSV file
            $skuListData = Import-Csv -Path $tempCsvPath -Header "TID", "GME SKU", "GME POS Name (36 Character Limit)", "Condition Abbreviated", "Cost", "Price (Rounded)", "Price" -Encoding UTF8
            
            # Skip the header row
            $skuListData = $skuListData | Select-Object -Skip 1
            
            # Clean up the temporary file
            Remove-Item -Path $tempCsvPath -Force
            
            Write-Log "Successfully imported SKU list with $($skuListData.Count) rows" "Green"
        } catch {
            Write-Log "Error importing SKU list: $_" "Red"
            return $false
        }
        
        Write-Log "Imported SKU list with $($skuListData.Count) rows" "White"
        
        # Get all combined spreadsheets
        $combinedFiles = Get-ChildItem -Path $CombinedSpreadsheetPath -Filter "Combined_Spreadsheet_*.xlsx"
        
        if ($combinedFiles.Count -eq 0) {
            Write-Log "No combined spreadsheets found in: $CombinedSpreadsheetPath" "Yellow"
            return $false
        }
        
        Write-Log "Found $($combinedFiles.Count) combined spreadsheets to process" "White"
        
        $totalFiles = $combinedFiles.Count
        $processedFiles = 0
        
        foreach ($combinedFile in $combinedFiles) {
            # Extract the number from the combined spreadsheet filename
            if ($combinedFile.Name -match "Combined_Spreadsheet_(\d+)\.xlsx") {
                $fileNumber = $matches[1]
                $gsFileName = "GS$fileNumber.xlsx"
                $gsFilePath = Join-Path -Path $FinalOutputPath -ChildPath $gsFileName
                
                Write-Log "Processing combined spreadsheet: $($combinedFile.Name) -> $gsFileName" "Cyan"
                
                # Import the combined spreadsheet
                $combinedData = Import-Excel -Path $combinedFile.FullName
                Write-Log "  Imported combined spreadsheet with $($combinedData.Count) rows" "White"
                
                # Create a new array to hold the processed data
                $processedData = @()
                $matchCount = 0
                $skipCount = 0
                $noMatchCount = 0
                $multipleMatchCount = 0
                
                # Process each row in the combined spreadsheet
                foreach ($row in $combinedData) {
                    # Skip empty rows or rows with "BLANK"
                    $isBlankRow = $true
                    foreach ($prop in $row.PSObject.Properties) {
                        if ($prop.Value -and $prop.Value -ne "BLANK") {
                            $isBlankRow = $false
                            break
                        }
                    }
                    
                    if ($isBlankRow) {
                        $skipCount++
                        continue
                    }
                    
                    # Get the TCGplayer Id
                    $tcgplayerId = $row.'TCGplayer Id'
                    
                    if (-not $tcgplayerId) {
                        Write-Log "  Row missing TCGplayer Id, skipping" "Yellow"
                        $skipCount++
                        continue
                    }
                    
                    # Find matching row(s) in SKU list
                    $matchingRows = $skuListData | Where-Object { $_.'TID' -eq $tcgplayerId }
                    
                    if (-not $matchingRows -or $matchingRows.Count -eq 0) {
                        Write-Log "  No match found in SKU list for TCGplayer Id: $tcgplayerId" "Yellow"
                        $noMatchCount++
                        continue
                    }
                    
                    if ($matchingRows.Count -gt 1) {
                        Write-Log "  Multiple matches found in SKU list for TCGplayer Id: $tcgplayerId" "Yellow"
                        $multipleMatchCount++
                        continue
                    }
                    
                    # Get the matched row
                    $matchedRow = $matchingRows[0]
                    
                    # Extract required data
                    $gmeSku = $matchedRow.'GME SKU'
                    $cardName = $matchedRow.'GME POS Name (36 Character Limit)'
                    
                    # Use the 'Condition Abbreviated' field
                    $condition = $matchedRow.'Condition Abbreviated'
                    
                    # Extract and clean monetary values
                    $cost = $matchedRow.'Cost' -replace '\$', '' -replace ',', ''
                    $priceRounded = $matchedRow.'Price (Rounded)' -replace '\$', '' -replace ',', ''
                    $price = $matchedRow.'Price' -replace '\$', '' -replace ',', ''
                    
                    # Create barcode (GME SKU + P + whole number from Price Rounded)
                    try {
                        # First, remove the dollar sign and any commas
                        $cleanPriceRounded = $priceRounded -replace '\$', '' -replace ',', ''
                        
                        # Handle empty or invalid price values
                        if ([string]::IsNullOrWhiteSpace($cleanPriceRounded)) {
                            $wholePriceNumber = 0
                            Write-Log "  Warning: Invalid or empty price for TCGplayer Id: $tcgplayerId, using 0" "Yellow"
                        } else {
                            # Remove any non-numeric characters except decimal point
                            $cleanPriceRounded = $cleanPriceRounded -replace '[^0-9\.]', ''
                            
                            # Try to parse as double and get the floor value
                            $wholePriceNumber = [math]::Floor([double]::Parse($cleanPriceRounded))
                            Write-Log "  Extracted whole price number: $wholePriceNumber from '$priceRounded'" "Gray"
                        }
                    } catch {
                        # If parsing fails for any reason, use 0 and log a warning
                        $wholePriceNumber = 0
                        Write-Log "  Warning: Could not parse price '$priceRounded' for TCGplayer Id: $tcgplayerId, using 0" "Yellow"
                    }
                    
                    $barcode = "$gmeSku" + "P" + "$wholePriceNumber"
                    
                    # Create a new object with the required properties
                    $newRow = [PSCustomObject]@{
                        'SKU' = $gmeSku
                        'Barcode' = $barcode
                        'Card Name' = $cardName
                        'Condition' = $condition
                        'Cost' = $cost
                        'Price (Rounded)' = $priceRounded
                        'Price' = $price
                    }
                    
                    # Add to processed data
                    $processedData += $newRow
                    $matchCount++
                    
                    Write-Log "  Matched TCGplayer Id: $tcgplayerId with SKU: $gmeSku" "White"
                }
                
                # Export the processed data to the GS file
                if ($processedData.Count -gt 0) {
                    $exportParams = @{
                        Path = $gsFilePath
                        WorksheetName = "Sheet1"
                        AutoSize = $true
                        TableName = "GSData"
                        TableStyle = "Medium2"
                        ErrorAction = "Stop"
                    }
                    
                    $processedData | Export-Excel @exportParams
                    
                    Write-Log "  Created $gsFileName with $matchCount matched rows" "Green"
                    Write-Log "  Skipped rows: $skipCount, No matches: $noMatchCount, Multiple matches: $multipleMatchCount" "White"
                } else {
                    Write-Log "  No matching data found for $($combinedFile.Name), GS file not created" "Yellow"
                }
            } else {
                Write-Log "  Could not extract number from filename: $($combinedFile.Name)" "Yellow"
            }
            
            $processedFiles++
            $progressPercentage = [int](($processedFiles / $totalFiles) * 100)
            Update-ProgressBar $progressPercentage
        }
        
        Write-Log "SKU list processing completed." "Cyan"
        Update-ProgressBar 100
        
        return $true
    }
    catch {
        Write-Log "Error during SKU list processing: $_" "Red"
        return $false
    }
}

# Function to start the spreadsheet combining process
function Start-SpreadsheetCombiningProcess {
    if ($spreadsheetLocations.Items.Count -lt 2) {
        Write-Log "At least two spreadsheet folder locations are required." "Yellow"
        return $false
    }
    
    if ([string]::IsNullOrWhiteSpace($destinationLocation.Text)) {
        Write-Log "Please select a combined destination location." "Yellow"
        return $false
    }
    
    Write-Log "Starting spreadsheet combining process..." "Cyan"
    Update-ProgressBar 0
    
    # Get all folder paths
    $folderPaths = New-Object System.Collections.ArrayList
    foreach ($item in $spreadsheetLocations.Items) {
        $null = $folderPaths.Add($item.Text)
    }
    
    # Get options from checkboxes
    $fileExtension = if ($optionCheckboxes[1].Checked) { "*.*" } else { "*.xlsx" }
    $excludeHeaders = $optionCheckboxes[2].Checked
    $duplicateQuantityTwoRows = $optionCheckboxes[3].Checked
    $normalizeQuantities = $optionCheckboxes[4].Checked
    $insertBlankRows = $optionCheckboxes[6].Checked  # BLANK option (Option 7)
    $reverseDataRows = $optionCheckboxes[7].Checked  # Reverse, Reverse option (Option 8)
    
    # Start combining spreadsheets with selected options
    $success = Combine-Spreadsheets `
        -FolderPaths $folderPaths `
        -DestinationPath $destinationLocation.Text `
        -FileExtension $fileExtension `
        -ExcludeHeaders $excludeHeaders `
        -DuplicateQuantityTwoRows $duplicateQuantityTwoRows `
        -NormalizeQuantities $normalizeQuantities `
        -InsertBlankRows $insertBlankRows `
        -ReverseDataRows $reverseDataRows
    
    if ($success) {
        Write-Log "Spreadsheet combining completed successfully." "Green"
        return $true
    } else {
        Write-Log "Spreadsheet combining completed with errors." "Red"
        return $false
    }
}

# Function to browse for a single folder
function Select-FolderDialog {
    # Use the standard Windows folder browser dialog
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = "Select a folder"
    $folderBrowser.SelectedPath = $PSScriptRoot  # Start in the application directory
    $folderBrowser.ShowNewFolderButton = $true   # Allow creating new folders
    
    $result = $folderBrowser.ShowDialog()
    
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return $folderBrowser.SelectedPath
    }
    
    return $null
}

# Function to browse for a single file
function Select-FileDialog {
    param (
        [Parameter(Mandatory=$false)]
        [string]$Filter = "All Files (*.*)|*.*",
        
        [Parameter(Mandatory=$false)]
        [string]$Title = "Select a file"
    )
    
    # Use the standard Windows open file dialog
    $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $fileDialog.Filter = $Filter
    $fileDialog.Title = $Title
    $fileDialog.InitialDirectory = $PSScriptRoot  # Start in the application directory
    $fileDialog.CheckFileExists = $true
    
    $result = $fileDialog.ShowDialog()
    
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return $fileDialog.FileName
    }
    
    return $null
}

#endregion

# Load required assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Xml
Add-Type -AssemblyName System.Xml.Linq

# Create the main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Spreadsheet Wrangler"
$form.Size = New-Object System.Drawing.Size(900, 850)
$form.MinimumSize = New-Object System.Drawing.Size(800, 750)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.MinimizeBox = $true
$form.Font = New-Object System.Drawing.Font("Segoe UI", 10)

# Create the menu bar
$menuBar = New-Object System.Windows.Forms.MenuStrip
$menuBar.BackColor = [System.Drawing.SystemColors]::Control
$form.MainMenuStrip = $menuBar

# File Menu
$fileMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$fileMenu.Text = "File"

# New Configuration
$newConfigMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$newConfigMenuItem.Text = "New Configuration"
$newConfigMenuItem.ShortcutKeys = [System.Windows.Forms.Keys]::Control -bor [System.Windows.Forms.Keys]::N
$newConfigMenuItem.Add_Click({
    # Reset all settings
    $backupLocations.Items.Clear()
    $spreadsheetLocations.Items.Clear()
    $destinationLocation.Text = ""
    
    # Reset all checkboxes
    foreach ($checkbox in $optionCheckboxes) {
        $checkbox.Checked = $false
    }
    
    Write-Log "Configuration reset to default." "Cyan"
})
$fileMenu.DropDownItems.Add($newConfigMenuItem)

# Open Configuration
$openConfigMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$openConfigMenuItem.Text = "Open Configuration..."
$openConfigMenuItem.ShortcutKeys = [System.Windows.Forms.Keys]::Control -bor [System.Windows.Forms.Keys]::O
$openConfigMenuItem.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "XML Files (*.xml)|*.xml|All Files (*.*)|*.*"
    $openFileDialog.Title = "Open Configuration"
    $openFileDialog.InitialDirectory = $PSScriptRoot
    
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        Load-Configuration -ConfigPath $openFileDialog.FileName
    }
})
$fileMenu.DropDownItems.Add($openConfigMenuItem)

# Save Configuration
$saveConfigMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$saveConfigMenuItem.Text = "Save Configuration"
$saveConfigMenuItem.ShortcutKeys = [System.Windows.Forms.Keys]::Control -bor [System.Windows.Forms.Keys]::S
$saveConfigMenuItem.Add_Click({
    # If we have a current config file, save to it, otherwise prompt for location
    if ($script:CurrentConfigFile -and (Test-Path $script:CurrentConfigFile)) {
        Save-Configuration -ConfigPath $script:CurrentConfigFile
    } else {
        $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveFileDialog.Filter = "XML Files (*.xml)|*.xml|All Files (*.*)|*.*"
        $saveFileDialog.Title = "Save Configuration"
        $saveFileDialog.InitialDirectory = $PSScriptRoot
        $saveFileDialog.DefaultExt = "xml"
        
        if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            Save-Configuration -ConfigPath $saveFileDialog.FileName
            $script:CurrentConfigFile = $saveFileDialog.FileName
        }
    }
})
$fileMenu.DropDownItems.Add($saveConfigMenuItem)

# Save Configuration As
$saveAsConfigMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$saveAsConfigMenuItem.Text = "Save Configuration As..."
$saveAsConfigMenuItem.Add_Click({
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "XML Files (*.xml)|*.xml|All Files (*.*)|*.*"
    $saveFileDialog.Title = "Save Configuration As"
    $saveFileDialog.InitialDirectory = $PSScriptRoot
    $saveFileDialog.DefaultExt = "xml"
    
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        Save-Configuration -ConfigPath $saveFileDialog.FileName
        $script:CurrentConfigFile = $saveFileDialog.FileName
    }
})
$fileMenu.DropDownItems.Add($saveAsConfigMenuItem)

# Separator
$fileMenu.DropDownItems.Add("-")

# Exit
$exitMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$exitMenuItem.Text = "Exit"
$exitMenuItem.ShortcutKeys = [System.Windows.Forms.Keys]::Alt -bor [System.Windows.Forms.Keys]::F4
$exitMenuItem.Add_Click({ $form.Close() })
$fileMenu.DropDownItems.Add($exitMenuItem)

# Help Menu
$helpMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$helpMenu.Text = "Help"

# About
$aboutMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$aboutMenuItem.Text = "About"
$aboutMenuItem.Add_Click({
    $aboutForm = New-Object System.Windows.Forms.Form
    $aboutForm.Text = "About Spreadsheet Wrangler"
    $aboutForm.Size = New-Object System.Drawing.Size(450, 300)
    $aboutForm.StartPosition = "CenterParent"
    $aboutForm.FormBorderStyle = "FixedDialog"
    $aboutForm.MaximizeBox = $false
    $aboutForm.MinimizeBox = $false
    
    $aboutPanel = New-Object System.Windows.Forms.TableLayoutPanel
    $aboutPanel.Dock = "Fill"
    $aboutPanel.RowCount = 3
    $aboutPanel.ColumnCount = 1
    $aboutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 60)))
    $aboutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 20)))
    $aboutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 20)))
    $aboutForm.Controls.Add($aboutPanel)
    
    # Main about text
    $aboutLabel = New-Object System.Windows.Forms.Label
    $aboutLabel.Text = "Spreadsheet Wrangler v1.4.0`n`nA powerful tool for backing up folders and combining spreadsheets.`n`nCreated by Bryant Welch`nCreated: $(Get-Date -Format 'yyyy-MM-dd')`n`n(c) 2025 Bryant Welch. All Rights Reserved"
    $aboutLabel.AutoSize = $false
    $aboutLabel.Dock = "Fill"
    $aboutLabel.TextAlign = "MiddleCenter"
    $aboutPanel.Controls.Add($aboutLabel, 0, 0)
    
    # GitHub link
    $linkLabel = New-Object System.Windows.Forms.LinkLabel
    $linkLabel.Text = "https://github.com/BryantWelch/Spreadsheet-Wrangler"
    $linkLabel.AutoSize = $false
    $linkLabel.Dock = "Fill"
    $linkLabel.TextAlign = "MiddleCenter"
    $linkLabel.LinkColor = [System.Drawing.Color]::Blue
    $linkLabel.ActiveLinkColor = [System.Drawing.Color]::Red
    $linkLabel.Add_LinkClicked({
        param($sender, $e)
        Start-Process "https://github.com/BryantWelch/Spreadsheet-Wrangler"
    })
    $aboutPanel.Controls.Add($linkLabel, 0, 1)
    
    # OK button
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "OK"
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $okButton.Dock = "Fill"
    $okButton.Margin = New-Object System.Windows.Forms.Padding(150, 10, 150, 10)
    $aboutPanel.Controls.Add($okButton, 0, 2)
    $aboutForm.AcceptButton = $okButton
    
    $aboutForm.ShowDialog() | Out-Null
})
$helpMenu.DropDownItems.Add($aboutMenuItem)

# Add menus to menu bar
$menuBar.Items.Add($fileMenu)
$menuBar.Items.Add($helpMenu)

# Add menu bar to form
$form.Controls.Add($menuBar)

# Create a container panel to hold everything below the menu bar
$containerPanel = New-Object System.Windows.Forms.Panel
$containerPanel.Dock = "Fill"
$containerPanel.Padding = New-Object System.Windows.Forms.Padding(0, $menuBar.Height, 0, 0)
$form.Controls.Add($containerPanel)

# Create a table layout panel for the main layout
$mainLayout = New-Object System.Windows.Forms.TableLayoutPanel
$mainLayout.Dock = "Fill"
$mainLayout.RowCount = 1
$mainLayout.ColumnCount = 2
$mainLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 30)))
$mainLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 70)))
$containerPanel.Controls.Add($mainLayout)

# Initialize current config file variable
$script:CurrentConfigFile = $null

# Create tooltip component for the entire form
$toolTip = New-Object System.Windows.Forms.ToolTip
$toolTip.AutoPopDelay = 5000
$toolTip.InitialDelay = 500
$toolTip.ReshowDelay = 200
$toolTip.ShowAlways = $true

#region Left Panel
$leftPanel = New-Object System.Windows.Forms.Panel
$leftPanel.Dock = "Fill"
$leftPanel.Padding = New-Object System.Windows.Forms.Padding(10)
$mainLayout.Controls.Add($leftPanel, 0, 0)

# Create a table layout for the left panel
$leftLayout = New-Object System.Windows.Forms.TableLayoutPanel
$leftLayout.Dock = "Fill"
$leftLayout.RowCount = 6
$leftLayout.ColumnCount = 1
$leftLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 20))) # Backup
$leftLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 20))) # Spreadsheet
$leftLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 15))) # Combined
$leftLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 15))) # SKU List
$leftLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 15))) # Final Output
$leftLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 15))) # Run Button
$leftPanel.Controls.Add($leftLayout)

# Backup Locations Section
$backupPanel = New-Object System.Windows.Forms.GroupBox
$backupPanel.Text = "Backup Locations"
$backupPanel.Dock = "Fill"
$leftLayout.Controls.Add($backupPanel, 0, 0)

$backupLayout = New-Object System.Windows.Forms.TableLayoutPanel
$backupLayout.Dock = "Fill"
$backupLayout.RowCount = 2
$backupLayout.ColumnCount = 1
$backupLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 85)))
$backupLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 35)))
$backupPanel.Controls.Add($backupLayout)

# List of backup locations
$backupLocations = New-Object System.Windows.Forms.ListView
$backupLocations.View = "Details"
$backupLocations.FullRowSelect = $true
$backupLocations.Columns.Add("Folder Path", -2)
$backupLocations.Dock = "Fill"
# Set tooltip for backup locations using the tooltip component
$toolTip.SetToolTip($backupLocations, "List of folders to back up. Select an item and press Delete or use the minus button to remove it.")
$backupLocations.Add_KeyDown({
    param($sender, $e)
    # Delete selected item when Delete key is pressed
    if ($e.KeyCode -eq 'Delete' -and $backupLocations.SelectedItems.Count -gt 0) {
        foreach ($item in $backupLocations.SelectedItems) {
            Write-Log "Removed backup location: $($item.Text)" "Yellow"
            $backupLocations.Items.Remove($item)
        }
    }
})
$backupLayout.Controls.Add($backupLocations, 0, 0)

# Button panel for backup locations
$backupButtonPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$backupButtonPanel.Dock = "Fill"
$backupButtonPanel.FlowDirection = "RightToLeft"
$backupButtonPanel.WrapContents = $false
$backupButtonPanel.Padding = New-Object System.Windows.Forms.Padding(5, 2, 5, 2)
$backupLayout.Controls.Add($backupButtonPanel, 0, 1)

# Remove button for backup locations
$removeBackupBtn = New-Object System.Windows.Forms.Button
$removeBackupBtn.Text = "-"
$removeBackupBtn.Width = 40
$removeBackupBtn.Height = 25
$removeBackupBtn.Margin = New-Object System.Windows.Forms.Padding(3, 0, 3, 0)
$removeBackupBtn.FlatStyle = "Flat"
$toolTip = New-Object System.Windows.Forms.ToolTip
$toolTip.SetToolTip($removeBackupBtn, "Remove selected backup location")
$removeBackupBtn.Add_Click({
    if ($backupLocations.SelectedItems.Count -gt 0) {
        foreach ($item in $backupLocations.SelectedItems) {
            Write-Log "Removed backup location: $($item.Text)" "Yellow"
            $backupLocations.Items.Remove($item)
        }
    } else {
        Write-Log "Please select a backup location to remove" "Yellow"
    }
})
$backupButtonPanel.Controls.Add($removeBackupBtn)

# Add button for backup locations
$addBackupBtn = New-Object System.Windows.Forms.Button
$addBackupBtn.Text = "+"
$addBackupBtn.Width = 40
$addBackupBtn.Height = 25
$addBackupBtn.Margin = New-Object System.Windows.Forms.Padding(3, 0, 3, 0)
$addBackupBtn.FlatStyle = "Flat"
$toolTip.SetToolTip($addBackupBtn, "Add a new folder to back up")
$addBackupBtn.Add_Click({
    $folderPath = Select-FolderDialog
    if ($folderPath) {
        $item = New-Object System.Windows.Forms.ListViewItem($folderPath)
        $backupLocations.Items.Add($item)
        Write-Log "Added backup location: $folderPath"
    }
})
$backupButtonPanel.Controls.Add($addBackupBtn)

# Spreadsheet Folders Section
$spreadsheetPanel = New-Object System.Windows.Forms.GroupBox
$spreadsheetPanel.Text = "Spreadsheet Folder Locations"
$spreadsheetPanel.Dock = "Fill"
$leftLayout.Controls.Add($spreadsheetPanel, 0, 1)

$spreadsheetLayout = New-Object System.Windows.Forms.TableLayoutPanel
$spreadsheetLayout.Dock = "Fill"
$spreadsheetLayout.RowCount = 2
$spreadsheetLayout.ColumnCount = 1
$spreadsheetLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 85)))
$spreadsheetLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 35)))
$spreadsheetPanel.Controls.Add($spreadsheetLayout)

# List of spreadsheet folder locations
$spreadsheetLocations = New-Object System.Windows.Forms.ListView
$spreadsheetLocations.View = "Details"
$spreadsheetLocations.FullRowSelect = $true
$spreadsheetLocations.Columns.Add("Folder Path", -2)
$spreadsheetLocations.Dock = "Fill"
# Set tooltip for spreadsheet locations using the tooltip component
$toolTip.SetToolTip($spreadsheetLocations, "List of folders containing spreadsheets to combine. Select an item and press Delete or use the minus button to remove it.")
$spreadsheetLocations.Add_KeyDown({
    param($sender, $e)
    # Delete selected item when Delete key is pressed
    if ($e.KeyCode -eq 'Delete' -and $spreadsheetLocations.SelectedItems.Count -gt 0) {
        foreach ($item in $spreadsheetLocations.SelectedItems) {
            Write-Log "Removed spreadsheet folder location: $($item.Text)" "Yellow"
            $spreadsheetLocations.Items.Remove($item)
        }
    }
})
$spreadsheetLayout.Controls.Add($spreadsheetLocations, 0, 0)

# Button panel for spreadsheet locations
$spreadsheetButtonPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$spreadsheetButtonPanel.Dock = "Fill"
$spreadsheetButtonPanel.FlowDirection = "RightToLeft"
$spreadsheetButtonPanel.WrapContents = $false
$spreadsheetButtonPanel.Padding = New-Object System.Windows.Forms.Padding(5, 2, 5, 2)
$spreadsheetLayout.Controls.Add($spreadsheetButtonPanel, 0, 1)

# Remove button for spreadsheet locations
$removeSpreadsheetBtn = New-Object System.Windows.Forms.Button
$removeSpreadsheetBtn.Text = "-"
$removeSpreadsheetBtn.Width = 40
$removeSpreadsheetBtn.Height = 25
$removeSpreadsheetBtn.Margin = New-Object System.Windows.Forms.Padding(3, 0, 3, 0)
$removeSpreadsheetBtn.FlatStyle = "Flat"
$toolTip.SetToolTip($removeSpreadsheetBtn, "Remove selected spreadsheet folder location")
$removeSpreadsheetBtn.Add_Click({
    if ($spreadsheetLocations.SelectedItems.Count -gt 0) {
        foreach ($item in $spreadsheetLocations.SelectedItems) {
            Write-Log "Removed spreadsheet folder location: $($item.Text)" "Yellow"
            $spreadsheetLocations.Items.Remove($item)
        }
    } else {
        Write-Log "Please select a spreadsheet folder location to remove" "Yellow"
    }
})
$spreadsheetButtonPanel.Controls.Add($removeSpreadsheetBtn)

# Add button for spreadsheet locations
$addSpreadsheetBtn = New-Object System.Windows.Forms.Button
$addSpreadsheetBtn.Text = "+"
$addSpreadsheetBtn.Width = 40
$addSpreadsheetBtn.Height = 25
$addSpreadsheetBtn.Margin = New-Object System.Windows.Forms.Padding(3, 0, 3, 0)
$addSpreadsheetBtn.FlatStyle = "Flat"
$toolTip.SetToolTip($addSpreadsheetBtn, "Add a new folder containing spreadsheets to combine")
$addSpreadsheetBtn.Add_Click({
    $folderPath = Select-FolderDialog
    if ($folderPath) {
        $item = New-Object System.Windows.Forms.ListViewItem($folderPath)
        $spreadsheetLocations.Items.Add($item)
        Write-Log "Added spreadsheet folder location: $folderPath"
    }
})
$spreadsheetButtonPanel.Controls.Add($addSpreadsheetBtn)

# Combined Destination Location Section
$destinationPanel = New-Object System.Windows.Forms.GroupBox
$destinationPanel.Text = "Combined Destination Location"
$destinationPanel.Dock = "Fill"
$leftLayout.Controls.Add($destinationPanel, 0, 2)

$destinationLayout = New-Object System.Windows.Forms.TableLayoutPanel
$destinationLayout.Dock = "Fill"
$destinationLayout.RowCount = 2
$destinationLayout.ColumnCount = 1
$destinationLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 85)))
$destinationLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 35)))
$destinationPanel.Controls.Add($destinationLayout)

# Destination location display
$destinationLocation = New-Object System.Windows.Forms.TextBox
$destinationLocation.ReadOnly = $true
$destinationLocation.Dock = "Fill"
$destinationLocation.BackColor = [System.Drawing.Color]::White
$toolTip.SetToolTip($destinationLocation, "Location where combined spreadsheets will be saved")
$destinationLayout.Controls.Add($destinationLocation, 0, 0)

# Browse button for destination location
$browseDestinationBtn = New-Object System.Windows.Forms.Button
$browseDestinationBtn.Text = "Browse..."
$browseDestinationBtn.Dock = "Right"
$browseDestinationBtn.Width = 80
$browseDestinationBtn.Height = 25
$browseDestinationBtn.Margin = New-Object System.Windows.Forms.Padding(3, 2, 3, 2)
$browseDestinationBtn.FlatStyle = "Flat"
$toolTip.SetToolTip($browseDestinationBtn, "Select a folder where combined spreadsheets will be saved")
$browseDestinationBtn.Add_Click({
    $folderPath = Select-FolderDialog
    if ($folderPath) {
        $destinationLocation.Text = $folderPath
        Write-Log "Set combined destination location: $folderPath"
    }
})
$destinationLayout.Controls.Add($browseDestinationBtn, 0, 1)

# Run Button
$runBtn = New-Object System.Windows.Forms.Button
$runBtn.Text = "Run"
$runBtn.Dock = "Fill"
$runBtn.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
$runBtn.ForeColor = [System.Drawing.Color]::White
$runBtn.FlatStyle = "Flat"
$runBtn.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
$toolTip.SetToolTip($runBtn, "Start the backup and spreadsheet combining process")
$runBtn.Add_Click({
    # Clear previous output
    $outputTextbox.Clear()
    
    # Initialize log file if logging is enabled
    if ($optionCheckboxes[5].Checked) { # Log to File option
        $logFileName = "SpreadsheetWrangler_Log_$(Get-TimeStampString).txt"
        $script:LogFilePath = Join-Path -Path $PWD.Path -ChildPath $logFileName
        
        # Create the log file with header
        "Spreadsheet Wrangler Log - Started at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" | Out-File -FilePath $script:LogFilePath
        "--------------------------------------------------------------" | Out-File -FilePath $script:LogFilePath -Append
        
        Write-Log "Logging to file: $script:LogFilePath" "Cyan"
    } else {
        $script:LogFilePath = $null
    }
    
    Write-Log "Starting operations..." "Cyan"
    
    # Start backup process if not skipped
    if (-not $optionCheckboxes[0].Checked) {
        Start-BackupProcess
    } else {
        Write-Log "Backup process skipped due to 'Skip Backup' option." "Yellow"
    }
    
    # Start spreadsheet combining process
    $combineSuccess = Start-SpreadsheetCombiningProcess
    
    # Process SKU list if spreadsheet combining was successful and SKU list path is provided
    if ($combineSuccess -and -not [string]::IsNullOrWhiteSpace($skuListLocation.Text) -and -not [string]::IsNullOrWhiteSpace($finalOutputLocation.Text)) {
        Write-Log "Starting SKU list processing..." "Cyan"
        $skuListSuccess = Process-SKUList -CombinedSpreadsheetPath $destinationLocation.Text -SKUListPath $skuListLocation.Text -FinalOutputPath $finalOutputLocation.Text
        
        if ($skuListSuccess) {
            Write-Log "SKU list processing completed successfully." "Green"
        } else {
            Write-Log "SKU list processing completed with errors." "Red"
        }
    } elseif ($combineSuccess) {
        if ([string]::IsNullOrWhiteSpace($skuListLocation.Text)) {
            Write-Log "SKU list processing skipped - No SKU list file specified." "Yellow"
        }
        if ([string]::IsNullOrWhiteSpace($finalOutputLocation.Text)) {
            Write-Log "SKU list processing skipped - No final output location specified." "Yellow"
        }
    }
    
    Write-Log "All operations completed." "Cyan"
    
    # Add final log entry if logging is enabled
    if ($script:LogFilePath -and (Test-Path $script:LogFilePath)) {
        Write-Log "Log file saved to: $script:LogFilePath" "Yellow"
        "--------------------------------------------------------------" | Out-File -FilePath $script:LogFilePath -Append
        "Spreadsheet Wrangler Log - Completed at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" | Out-File -FilePath $script:LogFilePath -Append
    }
})
# SKU List Location Section
$skuListPanel = New-Object System.Windows.Forms.GroupBox
$skuListPanel.Text = "SKU List Location"
$skuListPanel.Dock = "Fill"
$leftLayout.Controls.Add($skuListPanel, 0, 3)

$skuListLayout = New-Object System.Windows.Forms.TableLayoutPanel
$skuListLayout.Dock = "Fill"
$skuListLayout.RowCount = 2
$skuListLayout.ColumnCount = 1
$skuListLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 85)))
$skuListLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 35)))
$skuListPanel.Controls.Add($skuListLayout)

# SKU List location display
$skuListLocation = New-Object System.Windows.Forms.TextBox
$skuListLocation.ReadOnly = $true
$skuListLocation.Dock = "Fill"
$skuListLocation.BackColor = [System.Drawing.Color]::White
$toolTip.SetToolTip($skuListLocation, "Location of the SKUList.csv file")
$skuListLayout.Controls.Add($skuListLocation, 0, 0)

# Browse button for SKU List location
$browseSkuListBtn = New-Object System.Windows.Forms.Button
$browseSkuListBtn.Text = "Browse..."
$browseSkuListBtn.Dock = "Right"
$browseSkuListBtn.Width = 80
$browseSkuListBtn.Height = 25
$browseSkuListBtn.Margin = New-Object System.Windows.Forms.Padding(3, 2, 3, 2)
$browseSkuListBtn.FlatStyle = "Flat"
$toolTip.SetToolTip($browseSkuListBtn, "Select the SKUList.csv file")
$browseSkuListBtn.Add_Click({
    $filePath = Select-FileDialog -Filter "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*" -Title "Select SKU List File"
    if ($filePath) {
        $skuListLocation.Text = $filePath
        Write-Log "Set SKU List location: $filePath"
    }
})
$skuListLayout.Controls.Add($browseSkuListBtn, 0, 1)

# Final Output Location Section
$finalOutputPanel = New-Object System.Windows.Forms.GroupBox
$finalOutputPanel.Text = "Final Output Location"
$finalOutputPanel.Dock = "Fill"
$leftLayout.Controls.Add($finalOutputPanel, 0, 4)

$finalOutputLayout = New-Object System.Windows.Forms.TableLayoutPanel
$finalOutputLayout.Dock = "Fill"
$finalOutputLayout.RowCount = 2
$finalOutputLayout.ColumnCount = 1
$finalOutputLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 85)))
$finalOutputLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 35)))
$finalOutputPanel.Controls.Add($finalOutputLayout)

# Final Output location display
$finalOutputLocation = New-Object System.Windows.Forms.TextBox
$finalOutputLocation.ReadOnly = $true
$finalOutputLocation.Dock = "Fill"
$finalOutputLocation.BackColor = [System.Drawing.Color]::White
$toolTip.SetToolTip($finalOutputLocation, "Final GS##.xlsx file location")
$finalOutputLayout.Controls.Add($finalOutputLocation, 0, 0)

# Browse button for Final Output location
$browseFinalOutputBtn = New-Object System.Windows.Forms.Button
$browseFinalOutputBtn.Text = "Browse..."
$browseFinalOutputBtn.Dock = "Right"
$browseFinalOutputBtn.Width = 80
$browseFinalOutputBtn.Height = 25
$browseFinalOutputBtn.Margin = New-Object System.Windows.Forms.Padding(3, 2, 3, 2)
$browseFinalOutputBtn.FlatStyle = "Flat"
$toolTip.SetToolTip($browseFinalOutputBtn, "Select a folder for the final GS##.xlsx file")
$browseFinalOutputBtn.Add_Click({
    $folderPath = Select-FolderDialog
    if ($folderPath) {
        $finalOutputLocation.Text = $folderPath
        Write-Log "Set Final Output location: $folderPath"
    }
})
$finalOutputLayout.Controls.Add($browseFinalOutputBtn, 0, 1)

$leftLayout.Controls.Add($runBtn, 0, 5)
#endregion

#region Right Panel
$rightPanel = New-Object System.Windows.Forms.Panel
$rightPanel.Dock = "Fill"
$rightPanel.Padding = New-Object System.Windows.Forms.Padding(10)
$mainLayout.Controls.Add($rightPanel, 1, 0)

# Create a table layout for the right panel
$rightLayout = New-Object System.Windows.Forms.TableLayoutPanel
$rightLayout.Dock = "Fill"
$rightLayout.RowCount = 3
$rightLayout.ColumnCount = 1
$rightLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 20)))
$rightLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 70)))
$rightLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 10)))
$rightPanel.Controls.Add($rightLayout)

# Options Panel
$optionsPanel = New-Object System.Windows.Forms.GroupBox
$optionsPanel.Text = "Options"
$optionsPanel.Dock = "Fill"
$rightLayout.Controls.Add($optionsPanel, 0, 0)

$optionsLayout = New-Object System.Windows.Forms.TableLayoutPanel
$optionsLayout.Dock = "Fill"
$optionsLayout.RowCount = 3
$optionsLayout.ColumnCount = 3
$optionsLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 33.33)))
$optionsLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 33.33)))
$optionsLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 33.33)))
$optionsLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 33.33)))
$optionsLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 33.33)))
$optionsLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 33.33)))
$optionsPanel.Controls.Add($optionsLayout)

# Create checkboxes for options with specific functionality
$optionCheckboxes = @()

# Option 1: Skip backup process
$optionCheckboxes += $checkbox1 = New-Object System.Windows.Forms.CheckBox
$checkbox1.Text = "Skip Backup"
$checkbox1.Dock = "Fill"
$checkbox1.Margin = New-Object System.Windows.Forms.Padding(5)
$toolTip.SetToolTip($checkbox1, "Skip the backup process and only combine spreadsheets")
$optionsLayout.Controls.Add($checkbox1, 0, 0)

# Option 2: Support multiple file formats
$optionCheckboxes += $checkbox2 = New-Object System.Windows.Forms.CheckBox
$checkbox2.Text = "All Formats"
$checkbox2.Dock = "Fill"
$checkbox2.Margin = New-Object System.Windows.Forms.Padding(5)
$toolTip.SetToolTip($checkbox2, "Process all spreadsheet formats (.xlsx, .xls, .csv)")
$optionsLayout.Controls.Add($checkbox2, 1, 0)

# Option 3: Exclude headers
$optionCheckboxes += $checkbox3 = New-Object System.Windows.Forms.CheckBox
$checkbox3.Text = "No Headers"
$checkbox3.Dock = "Fill"
$checkbox3.Margin = New-Object System.Windows.Forms.Padding(5)
$toolTip.SetToolTip($checkbox3, "Exclude headers when combining spreadsheets")
$optionsLayout.Controls.Add($checkbox3, 2, 0)

# Option 4: Duplicate rows with '2' in 'Add to Quantity' column
$optionCheckboxes += $checkbox4 = New-Object System.Windows.Forms.CheckBox
$checkbox4.Text = "Duplicate Qty=2"
$checkbox4.Dock = "Fill"
$checkbox4.Margin = New-Object System.Windows.Forms.Padding(5)
$toolTip.SetToolTip($checkbox4, "Duplicate rows with '2' in the 'Add to Quantity' column")
$optionsLayout.Controls.Add($checkbox4, 0, 1)

# Option 5: Normalize all quantities to '1'
$optionCheckboxes += $checkbox5 = New-Object System.Windows.Forms.CheckBox
$checkbox5.Text = "Normalize Qty to 1"
$checkbox5.Dock = "Fill"
$checkbox5.Margin = New-Object System.Windows.Forms.Padding(5)
$toolTip.SetToolTip($checkbox5, "Change all values in 'Add to Quantity' column to '1' (runs after duplication)")
$optionsLayout.Controls.Add($checkbox5, 1, 1)

# Option 6: Log to File
$optionCheckboxes += $checkbox6 = New-Object System.Windows.Forms.CheckBox
$checkbox6.Text = "Log to File"
$checkbox6.Dock = "Fill"
$checkbox6.Margin = New-Object System.Windows.Forms.Padding(5)
$toolTip.SetToolTip($checkbox6, "Save terminal output to a log file in the application directory")
$optionsLayout.Controls.Add($checkbox6, 2, 1)

# Option 7: BLANK - Insert separator rows between spreadsheets
$optionCheckboxes += $checkbox7 = New-Object System.Windows.Forms.CheckBox
$checkbox7.Text = "BLANK"
$checkbox7.Dock = "Fill"
$checkbox7.Margin = New-Object System.Windows.Forms.Padding(5)
$toolTip.SetToolTip($checkbox7, "Insert 'BLANK' rows between data from different spreadsheets")
$optionsLayout.Controls.Add($checkbox7, 0, 2)

# Option 8: Reverse, Reverse - Reverse the order of data rows
$optionCheckboxes += $checkbox8 = New-Object System.Windows.Forms.CheckBox
$checkbox8.Text = "Reverse, Reverse"
$checkbox8.Dock = "Fill"
$checkbox8.Margin = New-Object System.Windows.Forms.Padding(5)
$toolTip.SetToolTip($checkbox8, "Reverse the order of data rows in the final combined spreadsheet")
$optionsLayout.Controls.Add($checkbox8, 1, 2)

# Option 9: Placeholder
$optionCheckboxes += $checkbox9 = New-Object System.Windows.Forms.CheckBox
$checkbox9.Text = "Option 9"
$checkbox9.Dock = "Fill"
$checkbox9.Margin = New-Object System.Windows.Forms.Padding(5)
$toolTip.SetToolTip($checkbox9, "Reserved for future functionality")
$optionsLayout.Controls.Add($checkbox9, 2, 2)

# Output Panel
$outputPanel = New-Object System.Windows.Forms.GroupBox
$outputPanel.Text = "Terminal Output"
$outputPanel.Dock = "Fill"
$rightLayout.Controls.Add($outputPanel, 0, 1)

# Terminal output textbox
$outputTextbox = New-Object System.Windows.Forms.RichTextBox
$outputTextbox.Dock = "Fill"
$outputTextbox.ReadOnly = $true
$outputTextbox.BackColor = [System.Drawing.Color]::Black
$outputTextbox.ForeColor = [System.Drawing.Color]::LightGreen
$outputTextbox.Font = New-Object System.Drawing.Font("Consolas", 10)
$toolTip.SetToolTip($outputTextbox, "Displays real-time progress and status information")
$outputPanel.Controls.Add($outputTextbox)

# Progress Bar
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Dock = "Fill"
$progressBar.Style = "Continuous"
$progressBar.Value = 0
$toolTip.SetToolTip($progressBar, "Shows overall progress of the current operation")
$rightLayout.Controls.Add($progressBar, 0, 2)
#endregion

# Function to save configuration to XML file
function Save-Configuration {
    param (
        [Parameter(Mandatory=$true)]
        [string]$ConfigPath
    )
    
    try {
        # Create XML document
        $xmlDoc = New-Object System.Xml.XmlDocument
        $xmlDeclaration = $xmlDoc.CreateXmlDeclaration("1.0", "UTF-8", $null)
        $xmlDoc.AppendChild($xmlDeclaration) | Out-Null
        
        # Create root element
        $rootElement = $xmlDoc.CreateElement("SpreadsheetWranglerConfig")
        $xmlDoc.AppendChild($rootElement) | Out-Null
        
        # Add backup locations
        $backupLocationsElement = $xmlDoc.CreateElement("BackupLocations")
        $rootElement.AppendChild($backupLocationsElement) | Out-Null
        
        foreach ($item in $backupLocations.Items) {
            $locationElement = $xmlDoc.CreateElement("Location")
            $locationElement.InnerText = $item.Text
            $backupLocationsElement.AppendChild($locationElement) | Out-Null
        }
        
        # Add spreadsheet locations
        $spreadsheetLocationsElement = $xmlDoc.CreateElement("SpreadsheetLocations")
        $rootElement.AppendChild($spreadsheetLocationsElement) | Out-Null
        
        foreach ($item in $spreadsheetLocations.Items) {
            $locationElement = $xmlDoc.CreateElement("Location")
            $locationElement.InnerText = $item.Text
            $spreadsheetLocationsElement.AppendChild($locationElement) | Out-Null
        }
        
        # Add destination location
        $destinationElement = $xmlDoc.CreateElement("DestinationLocation")
        $destinationElement.InnerText = $destinationLocation.Text
        $rootElement.AppendChild($destinationElement) | Out-Null
        
        # Add SKU List location
        $skuListElement = $xmlDoc.CreateElement("SKUListLocation")
        $skuListElement.InnerText = $skuListLocation.Text
        $rootElement.AppendChild($skuListElement) | Out-Null
        
        # Add Final Output location
        $finalOutputElement = $xmlDoc.CreateElement("FinalOutputLocation")
        $finalOutputElement.InnerText = $finalOutputLocation.Text
        $rootElement.AppendChild($finalOutputElement) | Out-Null
        
        # Add options
        $optionsElement = $xmlDoc.CreateElement("Options")
        $rootElement.AppendChild($optionsElement) | Out-Null
        
        for ($i = 0; $i -lt $optionCheckboxes.Count; $i++) {
            $optionElement = $xmlDoc.CreateElement("Option")
            $optionElement.SetAttribute("Index", $i)
            $optionElement.SetAttribute("Checked", $optionCheckboxes[$i].Checked)
            $optionsElement.AppendChild($optionElement) | Out-Null
        }
        
        # Save the XML document
        $xmlDoc.Save($ConfigPath)
        
        Write-Log "Configuration saved to: $ConfigPath" "Green"
        return $true
    }
    catch {
        Write-Log "Error saving configuration: $_" "Red"
        return $false
    }
}

# Function to load configuration from XML file
function Load-Configuration {
    param (
        [Parameter(Mandatory=$true)]
        [string]$ConfigPath
    )
    
    try {
        # Check if file exists
        if (-not (Test-Path -Path $ConfigPath)) {
            Write-Log "Configuration file not found: $ConfigPath" "Red"
            return $false
        }
        
        # Load XML document
        $xmlDoc = New-Object System.Xml.XmlDocument
        $xmlDoc.Load($ConfigPath)
        
        # Clear current settings
        $backupLocations.Items.Clear()
        $spreadsheetLocations.Items.Clear()
        $destinationLocation.Text = ""
        
        foreach ($checkbox in $optionCheckboxes) {
            $checkbox.Checked = $false
        }
        
        # Load backup locations
        $backupLocationsElement = $xmlDoc.SelectSingleNode("//BackupLocations")
        if ($backupLocationsElement) {
            foreach ($locationElement in $backupLocationsElement.SelectNodes("Location")) {
                $item = New-Object System.Windows.Forms.ListViewItem($locationElement.InnerText)
                $backupLocations.Items.Add($item)
            }
        }
        
        # Load spreadsheet locations
        $spreadsheetLocationsElement = $xmlDoc.SelectSingleNode("//SpreadsheetLocations")
        if ($spreadsheetLocationsElement) {
            foreach ($locationElement in $spreadsheetLocationsElement.SelectNodes("Location")) {
                $item = New-Object System.Windows.Forms.ListViewItem($locationElement.InnerText)
                $spreadsheetLocations.Items.Add($item)
            }
        }
        
        # Load destination location
        $destinationElement = $xmlDoc.SelectSingleNode("//DestinationLocation")
        if ($destinationElement) {
            $destinationLocation.Text = $destinationElement.InnerText
        }
        
        # Load SKU List location
        $skuListElement = $xmlDoc.SelectSingleNode("//SKUListLocation")
        if ($skuListElement) {
            $skuListLocation.Text = $skuListElement.InnerText
        }
        
        # Load Final Output location
        $finalOutputElement = $xmlDoc.SelectSingleNode("//FinalOutputLocation")
        if ($finalOutputElement) {
            $finalOutputLocation.Text = $finalOutputElement.InnerText
        }
        
        # Load options
        $optionsElement = $xmlDoc.SelectSingleNode("//Options")
        if ($optionsElement) {
            foreach ($optionElement in $optionsElement.SelectNodes("Option")) {
                $index = [int]$optionElement.GetAttribute("Index")
                $checked = [System.Convert]::ToBoolean($optionElement.GetAttribute("Checked"))
                
                if ($index -ge 0 -and $index -lt $optionCheckboxes.Count) {
                    $optionCheckboxes[$index].Checked = $checked
                }
            }
        }
        
        # Set current config file
        $script:CurrentConfigFile = $ConfigPath
        
        Write-Log "Configuration loaded from: $ConfigPath" "Green"
        return $true
    }
    catch {
        Write-Log "Error loading configuration: $_" "Red"
        return $false
    }
}

# Initialize the form with some sample data for visualization
Write-Log "Application initialized and ready to run." "Cyan"
Write-Log "Please add backup, spreadsheet, and combined destination folder locations." "White"
Write-Log "Tip: You can remove locations by selecting them and pressing Delete." "Yellow"

# Show the form
$form.ShowDialog()
