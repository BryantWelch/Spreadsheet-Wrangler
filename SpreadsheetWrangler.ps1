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

# Global variables for application state
$script:LogFilePath = $null
$script:RecentFiles = @() # List of recently used configuration files
$script:MaxRecentFiles = 5 # Maximum number of recent files to track
$script:AppSettingsFile = Join-Path -Path $PSScriptRoot -ChildPath "SpreadsheetWrangler.settings.xml" # Settings file path

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
        
        # Variables to track handbuilt folder and its spreadsheets
        $handbuiltFolderPath = $null
        $handbuiltSpreadsheet = $null
        $useHandbuiltSingleSpreadsheet = $false
        
        # Check if any of the folders is named "handbuilt" (case-insensitive)
        foreach ($folderPath in $FolderPaths) {
            $folderName = Split-Path -Path $folderPath -Leaf
            if ($folderName -match "(?i)handbuilt") {
                $handbuiltFolderPath = $folderPath
                Write-Log "Found handbuilt folder: $handbuiltFolderPath" "Cyan"
                break
            }
        }
        
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
        
        # Check if handbuilt folder has only one spreadsheet
        if ($handbuiltFolderPath) {
            # Get all spreadsheet files in the handbuilt folder
            $handbuiltFiles = @()
            
            if ($FileExtension -eq "*.*") {
                $handbuiltFiles += Get-ChildItem -Path $handbuiltFolderPath -Filter "*.xlsx" -File
                $handbuiltFiles += Get-ChildItem -Path $handbuiltFolderPath -Filter "*.xls" -File
                $handbuiltFiles += Get-ChildItem -Path $handbuiltFolderPath -Filter "*.csv" -File
            } else {
                $handbuiltFiles = Get-ChildItem -Path $handbuiltFolderPath -Filter $FileExtension -File
            }
            
            # Check if there's only one spreadsheet in the handbuilt folder
            if ($handbuiltFiles.Count -eq 1) {
                $handbuiltSpreadsheet = $handbuiltFiles[0].FullName
                $useHandbuiltSingleSpreadsheet = $true
                Write-Log "Handbuilt folder contains only one spreadsheet: $($handbuiltFiles[0].Name)" "Cyan"
                Write-Log "Will combine this spreadsheet with each spreadsheet group" "Cyan"
            } else {
                Write-Log "Handbuilt folder contains $($handbuiltFiles.Count) spreadsheets - using standard grouping logic" "White"
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
            
            # If we have a handbuilt spreadsheet with only one file, add it to all groups
            if ($useHandbuiltSingleSpreadsheet) {
                Write-Log "Adding handbuilt spreadsheet to group $groupNumber" "Cyan"
                
                # Create a new list with the handbuilt spreadsheet first (to ensure its headers are used)
                $combinedFiles = New-Object System.Collections.ArrayList
                $null = $combinedFiles.Add($handbuiltSpreadsheet)
                
                # Add all existing files from this group
                foreach ($file in $files) {
                    $null = $combinedFiles.Add($file)
                }
                
                # Replace the files array with our new combined list
                $files = $combinedFiles
            } 
            # Skip groups with only one spreadsheet if we're not using handbuilt mode
            elseif ($files.Count -lt 2) {
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
                        # Duplicate rows based on the numeric value in the 'Add to Quantity' column if option is enabled
                        if ($DuplicateQuantityTwoRows) {
                            Write-Log "  Processing 'Duplicate by Quantity' option..." "White"
                            
                            $rowsToAdd = @()
                            $totalDuplicates = 0
                            
                            # Process each row based on its quantity value
                            for ($i = 0; $i -lt $combinedData.Count; $i++) {
                                $row = $combinedData[$i]
                                $qtyValue = $row.$addToQuantityColName
                                
                                # Try to parse the quantity as an integer
                                if ([int]::TryParse($qtyValue, [ref]$null)) {
                                    $qty = [int]$qtyValue
                                    
                                    # If quantity is greater than 1, duplicate the row (qty-1) times
                                    if ($qty -gt 1) {
                                        for ($j = 1; $j -lt $qty; $j++) {
                                            # Create a clone of the row to avoid reference issues
                                            $clonedRow = [PSCustomObject]@{}
                                            foreach ($prop in $row.PSObject.Properties) {
                                                $clonedRow | Add-Member -MemberType NoteProperty -Name $prop.Name -Value $prop.Value
                                            }
                                            $rowsToAdd += $clonedRow
                                            $totalDuplicates++
                                        }
                                    }
                                }
                            }
                            
                            # Add the duplicated rows
                            $combinedData += $rowsToAdd
                            Write-Log "  Duplicated $totalDuplicates rows based on 'Add to Quantity' values" "Green"
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

# Function to process a single spreadsheet with duplication and blank row insertion
function Process-SingleSpreadsheet {
    param (
        [Parameter(Mandatory=$true)]
        [string]$SingleFilePath,

        [Parameter(Mandatory=$true)]
        [string]$DestinationPath
    )

    Write-Log "Processing single spreadsheet: $SingleFilePath" "Cyan"
    Update-ProgressBar 0

    try {
        # Prompt for duplication count
        Add-Type -AssemblyName Microsoft.VisualBasic
        $duplicationCountInput = [Microsoft.VisualBasic.Interaction]::InputBox("How many times do you want to duplicate the data (including the original)? Enter a number >= 1.", "Duplication Count", "1")
        
        if ([string]::IsNullOrWhiteSpace($duplicationCountInput)) {
            Write-Log "Duplication cancelled by user." "Yellow"
            # Ensure progress bar is reset if user cancels early
            Update-ProgressBar 0
            return $false
        }

        $duplicationCount = 0
        if (-not [int]::TryParse($duplicationCountInput, [ref]$duplicationCount) -or $duplicationCount -lt 1) {
            Write-Log "Invalid duplication count entered. Please enter a number greater than or equal to 1." "Red"
            Update-ProgressBar 0
            return $false
        }

        Write-Log "Data will be duplicated $duplicationCount time(s)." "White"

        # Import the spreadsheet
        $data = Import-Excel -Path $SingleFilePath -ErrorAction Stop
        
        if (-not $data -or $data.Count -eq 0) {
            Write-Log "No data found in the spreadsheet or spreadsheet is empty: $SingleFilePath" "Red"
            Update-ProgressBar 0
            return $false
        }

        # Get header names from the properties of the first data object
        $header = $data[0].PSObject.Properties | ForEach-Object { $_.Name }
        $dataRows = $data # All rows including the first (which Import-Excel uses for property names)

        $processedData = New-Object System.Collections.ArrayList

        $totalIterations = $duplicationCount
        for ($i = 1; $i -le $totalIterations; $i++) {
            Write-Log "Processing duplication iteration $i of $totalIterations..." "White"
            # Add data rows
            foreach ($row in $dataRows) {
                [void]$processedData.Add($row.PSObject.Copy()) # Add a copy to avoid issues if modifying objects later
            }

            # Add two blank rows if not the last iteration
            if ($i -lt $totalIterations) {
                Write-Log "  Adding 2 blank rows..." "White"
                $blankRow = New-Object PSObject
                foreach ($colName in $header) {
                    Add-Member -InputObject $blankRow -MemberType NoteProperty -Name $colName -Value "BLANK"
                }
                [void]$processedData.Add($blankRow)
                [void]$processedData.Add($blankRow.PSObject.Copy()) # Add a second, distinct blank row object
            }
            Update-ProgressBar ([int](($i / $totalIterations) * 90)) # Progress up to 90% for processing
        }

        # Save the processed data
        $baseFileName = [System.IO.Path]::GetFileNameWithoutExtension($SingleFilePath)
        $fileName = "${baseFileName}_Duplicated.xlsx"
        $outputFilePath = Join-Path -Path $DestinationPath -ChildPath $fileName

        Write-Log "Saving processed single spreadsheet to: $outputFilePath" "White"
        $processedData | Export-Excel -Path $outputFilePath -AutoSize -TableName "Data" -ErrorAction Stop
        
        Write-Log "Single spreadsheet processing completed successfully: $outputFilePath" "Green"
        Update-ProgressBar 100
        return $outputFilePath

    } catch {
        $errorMessage = $_.Exception.Message
        Write-Log "Error during single spreadsheet processing: $errorMessage" "Red"
        Write-Log "Stack Trace: $($_.ScriptStackTrace)" "Red"
        Update-ProgressBar 100 # Reset progress bar on error, or set to 0 if preferred
        return $null
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
        
        # Create a hashtable for fast SKU lookups indexed by TID
        Write-Log "Creating SKU lookup table for faster processing..." "White"
        $skuLookup = @{}
        foreach ($item in $skuListData) {
            if ($item.TID) {
                # Convert TID to string to ensure consistent lookup
                $tidKey = $item.TID.ToString().Trim()
                $skuLookup[$tidKey] = $item
            }
        }
        Write-Log "Created lookup table with $($skuLookup.Count) SKUs" "White"
        
        # Get all combined spreadsheets and sort them numerically
        $combinedFiles = Get-ChildItem -Path $CombinedSpreadsheetPath -Filter "Combined_Spreadsheet_*.xlsx" | 
            Sort-Object { [int]($_.Name -replace 'Combined_Spreadsheet_(\d+)\.xlsx', '$1') }
        
        if ($combinedFiles.Count -eq 0) {
            Write-Log "No combined spreadsheets found in: $CombinedSpreadsheetPath" "Yellow"
            return $false
        }
        
        Write-Log "Found $($combinedFiles.Count) combined spreadsheets to process" "White"
        
        $totalFiles = $combinedFiles.Count
        $processedFiles = 0
        
        # Create an array to hold all missing matches for the GS_Missing spreadsheet
        $missingData = @()
        
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
                    
                    # Use hashtable for fast lookup instead of filtering the entire SKU list
                    # Convert TCGplayer Id to string to ensure consistent lookup
                    $tcgplayerIdKey = $tcgplayerId.ToString().Trim()
                    $matchedRow = $skuLookup[$tcgplayerIdKey]
                    
                    if (-not $matchedRow) {
                        Write-Log "  No match found in SKU list for TCGplayer Id: $tcgplayerId" "Yellow"
                        $noMatchCount++
                        
                        # Add the unmatched row to the missingData array
                        # Create a clone of the row to avoid reference issues
                        $missingRow = [PSCustomObject]@{}
                        foreach ($prop in $row.PSObject.Properties) {
                            $missingRow | Add-Member -MemberType NoteProperty -Name $prop.Name -Value $prop.Value
                        }
                        $missingData += $missingRow
                        continue
                    }
                    
                    # Check for multiple matches is no longer needed with hashtable approach
                    # as we're storing one SKU per TID in the hashtable
                    
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
            
            # If we found unmatched rows in this spreadsheet, add a separator for the next spreadsheet
            if ($noMatchCount -gt 0 -and $processedFiles -lt ($totalFiles - 1)) {
                # Get property names from the first row to ensure consistent structure
                if ($missingData.Count -gt 0) {
                    $firstRow = $missingData[0]
                    $propNames = $firstRow.PSObject.Properties.Name
                    
                    # Create a separator row with the spreadsheet name in the first column
                    $separatorRow = [PSCustomObject]@{}
                    foreach ($propName in $propNames) {
                        if ($propName -eq 'TCGplayer Id') {
                            $separatorRow | Add-Member -MemberType NoteProperty -Name $propName -Value "COMBINED_SPREADSHEET_$fileNumber"
                        } else {
                            $separatorRow | Add-Member -MemberType NoteProperty -Name $propName -Value $null
                        }
                    }
                    $missingData += $separatorRow
                }
            }
            
            $processedFiles++
            $progressPercentage = [int](($processedFiles / $totalFiles) * 100)
            Update-ProgressBar $progressPercentage
        }
        
        # Create the GS_Missing spreadsheet if we have any missing data
        if ($missingData.Count -gt 0) {
            try {
                $gsMissingFilePath = Join-Path -Path $FinalOutputPath -ChildPath "GS_Missing.xlsx"
                
                # Ensure the export path exists
                $exportDir = Split-Path -Path $gsMissingFilePath -Parent
                if (-not (Test-Path -Path $exportDir)) {
                    New-Item -Path $exportDir -ItemType Directory -Force | Out-Null
                }
                
                Write-Log "Exporting $($missingData.Count) unmatched rows to GS_Missing.xlsx..." "White"
                
                $exportParams = @{
                    Path = $gsMissingFilePath
                    WorksheetName = "Sheet1"
                    AutoSize = $true
                    TableName = "MissingData"
                    TableStyle = "Medium2"
                    ErrorAction = "Stop"
                }
                
                # Export with error handling
                $missingData | Export-Excel @exportParams
                
                Write-Log "Created GS_Missing.xlsx with $($missingData.Count) unmatched rows" "Green"
            } catch {
                Write-Log "Error creating GS_Missing.xlsx: $_" "Red"
                
                # Fallback method if Export-Excel fails
                try {
                    Write-Log "Attempting alternative export method..." "Yellow"
                    $missingData | ConvertTo-Csv -NoTypeInformation | Out-File -FilePath "$FinalOutputPath\GS_Missing.csv" -Encoding UTF8
                    Write-Log "Created GS_Missing.csv as fallback" "Green"
                } catch {
                    Write-Log "Fallback export also failed: $_" "Red"
                }
            }
        } else {
            Write-Log "No unmatched rows found, GS_Missing.xlsx not created" "Yellow"
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
    # Check if Single Spreadsheet option ($optionCheckboxes[9]) is selected
    if ($optionCheckboxes[9].Checked) {
        Write-Log "Single Spreadsheet mode selected." "Cyan"
        if ([string]::IsNullOrWhiteSpace($textBoxSingleSpreadsheetFile.Text) -or -not (Test-Path $textBoxSingleSpreadsheetFile.Text)) {
            Write-Log "Please select a valid single spreadsheet file." "Red"
            return $false
        }
        if ([string]::IsNullOrWhiteSpace($destinationLocation.Text)) {
            Write-Log "Please select a combined destination location." "Red"
            return $false
        }

        # Call the new function to process the single spreadsheet
        # This function will be defined elsewhere
        return Process-SingleSpreadsheet -SingleFilePath $textBoxSingleSpreadsheetFile.Text -DestinationPath $destinationLocation.Text
    }

    # Original multi-folder logic starts here
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
    $fileExtension = if ($optionCheckboxes[5].Checked) { "*.*" } else { "*.xlsx" }
    $excludeHeaders = $optionCheckboxes[10].Checked
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

# Function to create labels from template
function Create-Labels {
    param (
        [Parameter(Mandatory=$true)]
        [string]$InputFolder,
        
        [Parameter(Mandatory=$true)]
        [string]$OutputFolder,
        
        [Parameter(Mandatory=$false)]
        [string]$ParamTemplate = "",
        
        [Parameter(Mandatory=$false)]
        [string]$PrtTemplate = "",
        
        [Parameter(Mandatory=$false)]
        [string]$DymoTemplate = ""
    )
    
    try {
        # Check if paths exist
        if (-not (Test-Path -Path $InputFolder)) {
            Write-Log "Input folder does not exist: $InputFolder" "Red"
            return $false
        }
        
        if (-not (Test-Path -Path $OutputFolder)) {
            # Create output folder if it doesn't exist
            New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null
            Write-Log "Created output folder: $OutputFolder"
        }
        
        # Check for provided template files
        $templateFilesExist = $true
        
        # Check for param template
        if (-not [string]::IsNullOrWhiteSpace($ParamTemplate) -and (Test-Path -Path $ParamTemplate)) {
            $paramTemplatePath = $ParamTemplate
            Write-Log "Using provided param template: $paramTemplatePath" "Cyan"
        } else {
            $templateFilesExist = $false
            Write-Log "No param template provided or file not found" "Yellow"
        }
        
        # Check for prt template
        if (-not [string]::IsNullOrWhiteSpace($PrtTemplate) -and (Test-Path -Path $PrtTemplate)) {
            $prtTemplatePath = $PrtTemplate
            Write-Log "Using provided prt template: $prtTemplatePath" "Cyan"
        } else {
            $templateFilesExist = $false
            Write-Log "No prt template provided or file not found" "Yellow"
        }
        
        # Check for dymo template (not used yet but stored for future use)
        if (-not [string]::IsNullOrWhiteSpace($DymoTemplate) -and (Test-Path -Path $DymoTemplate)) {
            $dymoTemplatePath = $DymoTemplate
            Write-Log "Using provided dymo template: $dymoTemplatePath" "Cyan"
        } else {
            Write-Log "No dymo template provided or file not found (not required)" "Yellow"
        }
        
        # Check if template files exist
        $templateFilesExist = $true
        
        if (-not (Test-Path -Path $paramTemplatePath)) {
            Write-Log "Param template file not found: $paramTemplatePath" "Yellow"
            $templateFilesExist = $false
        }
        
        if (-not (Test-Path -Path $prtTemplatePath)) {
            Write-Log "PRT template file not found: $prtTemplatePath" "Yellow"
            $templateFilesExist = $false
        }
        
        if (-not $templateFilesExist) {
            Write-Log "Template files not found. Using default empty templates." "Yellow"
            # Create default empty templates
            $paramTemplateContent = @"
<?xml version="1.0" encoding="utf-8"?>
<Parameters></Parameters>
"@
            $prtTemplateContent = @"
<?xml version="1.0" encoding="utf-8"?>
<PrintTemplate></PrintTemplate>
"@
        } else {
            # Read template files
            $paramTemplateContent = Get-Content -Path $paramTemplatePath -Raw
            $prtTemplateContent = Get-Content -Path $prtTemplatePath -Raw
        }
        
        # Start the label creation process
        Write-Log "Starting label creation process..." "Cyan"
        Update-ProgressBar 0
        
        # Get all Excel files in the input folder
        $excelFiles = Get-ChildItem -Path $InputFolder -Filter "GS*.xlsx"
        $totalFiles = $excelFiles.Count
        $processedFiles = 0
        
        if ($totalFiles -eq 0) {
            Write-Log "No GS*.xlsx files found in input folder: $InputFolder" "Yellow"
            Update-ProgressBar 100
            return $true
        }
        
        Write-Log "Found $totalFiles GS*.xlsx files to process" "White"
        
        # Process each Excel file
        foreach ($excelFile in $excelFiles) {
            Write-Log "Processing file: $($excelFile.Name)" "White"
            
            # Extract the GS number from the filename
            if ($excelFile.BaseName -match "GS(\d+)") {
                $gsNumber = $matches[1]
                Write-Log "  Extracted GS number: $gsNumber" "White"
                
                # Create output file paths
                $paramFileName = "GS$gsNumber.param"
                $prtFileName = "GS$gsNumber.prt"
                $xmlFileName = "GS$gsNumber.xml"
                $tskFileName = "GS$gsNumber.tsk"
                
                $paramFilePath = Join-Path -Path $OutputFolder -ChildPath $paramFileName
                $prtFilePath = Join-Path -Path $OutputFolder -ChildPath $prtFileName
                $xmlFilePath = Join-Path -Path $OutputFolder -ChildPath $xmlFileName
                $tskFilePath = Join-Path -Path $OutputFolder -ChildPath $tskFileName
                $zipFilePath = Join-Path -Path $OutputFolder -ChildPath "GS$gsNumber.zip"
                
                # Create temporary folder for files to zip
                $tempFolder = Join-Path -Path $OutputFolder -ChildPath "temp_GS$gsNumber"
                if (-not (Test-Path -Path $tempFolder)) {
                    New-Item -Path $tempFolder -ItemType Directory -Force | Out-Null
                }
                
                $tempParamPath = Join-Path -Path $tempFolder -ChildPath $paramFileName
                $tempPrtPath = Join-Path -Path $tempFolder -ChildPath $prtFileName
                $tempXmlPath = Join-Path -Path $tempFolder -ChildPath $xmlFileName
                
                try {
                    # Import Excel data
                    Write-Log "  Importing Excel data..." "White"
                    $excelData = Import-Excel -Path $excelFile.FullName -ErrorAction Stop
                    
                    if ($excelData.Count -eq 0) {
                        Write-Log "  No data found in Excel file" "Yellow"
                        continue
                    }
                    
                    # 1. Create the .param file (exact copy with new name)
                    $paramTemplateContent | Out-File -FilePath $tempParamPath -Encoding utf8
                    Write-Log "  Created $paramFileName" "Green"
                    
                    # 2. Create the .prt file (with data from first row)
                    $firstRow = $excelData[0]
                    $cardName = $firstRow.'Card Name'
                    $price = $firstRow.'Price (Rounded)'
                    $barcode = $firstRow.Barcode
                    $sku = $firstRow.SKU
                    
                    # Replace data in .prt template
                    $prtContent = $prtTemplateContent
                    
                    # Escape XML special characters in the data (order is important: & must be first)
                    $escapedCardName = $cardName -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;'
                    $escapedPrice = $price -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;'
                    $escapedBarcode = $barcode -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;'
                    $escapedSku = $sku -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;'
                    
                    # Replace each marker with the corresponding escaped data from the Excel file
                    $prtContent = $prtContent -replace 'Change Card Name Data Here', $escapedCardName
                    $prtContent = $prtContent -replace 'Change Price \(Rounded\) Data Here', $escapedPrice
                    $prtContent = $prtContent -replace 'Change Barcode Data Here', $escapedBarcode
                    $prtContent = $prtContent -replace 'Change SKU Data Here', $escapedSku
                    
                    # Save the .prt file
                    $prtContent | Out-File -FilePath $tempPrtPath -Encoding utf8
                    Write-Log "  Created $prtFileName with data from first row" "Green"
                    
                    # 3. Create the .xml file with all data rows
                    # Create XML document
                    $xmlDoc = New-Object System.Xml.XmlDocument
                    
                    # Create declaration
                    $declaration = $xmlDoc.CreateXmlDeclaration("1.0", "utf-8", $null)
                    $xmlDoc.AppendChild($declaration) | Out-Null
                    
                    # Create root element
                    $rootElement = $xmlDoc.CreateElement("DBdata")
                    $xmlDoc.AppendChild($rootElement) | Out-Null
                    
                    # Add data rows
                    $rowId = 1
                    foreach ($row in $excelData) {
                        # Create DataRow element
                        $dataRowElement = $xmlDoc.CreateElement("DataRow")
                        $dataRowElement.SetAttribute("id", $rowId.ToString())
                        
                        # Create Card Name element
                        $cardNameElement = $xmlDoc.CreateElement("Excel")
                        $cardNameElement.SetAttribute("col", "Card Name")
                        $cardNameElement.InnerText = $row.'Card Name'
                        $dataRowElement.AppendChild($cardNameElement) | Out-Null
                        
                        # Create Price element
                        $priceElement = $xmlDoc.CreateElement("Excel")
                        $priceElement.SetAttribute("col", "Price (Rounded)")
                        $priceElement.InnerText = $row.'Price (Rounded)'
                        $dataRowElement.AppendChild($priceElement) | Out-Null
                        
                        # Create Barcode element
                        $barcodeElement = $xmlDoc.CreateElement("Excel")
                        $barcodeElement.SetAttribute("col", "Barcode")
                        $barcodeElement.InnerText = $row.Barcode
                        $dataRowElement.AppendChild($barcodeElement) | Out-Null
                        
                        # Create SKU element
                        $skuElement = $xmlDoc.CreateElement("Excel")
                        $skuElement.SetAttribute("col", "SKU")
                        $skuElement.InnerText = $row.SKU
                        $dataRowElement.AppendChild($skuElement) | Out-Null
                        
                        # Add DataRow to root
                        $rootElement.AppendChild($dataRowElement) | Out-Null
                        
                        $rowId++
                    }
                    
                    # Save XML to file
                    $xmlDoc.Save($tempXmlPath)
                    Write-Log "  Created $xmlFileName with $($excelData.Count) rows" "Green"
                    
                    # 4. Create a ZIP archive and rename to .tsk
                    Write-Log "  Creating archive..." "White"
                    
                    # Check if Compress-Archive is available (PowerShell 5.0+)
                    if (Get-Command -Name Compress-Archive -ErrorAction SilentlyContinue) {
                        Compress-Archive -Path "$tempFolder\*" -DestinationPath $zipFilePath -Force
                        
                        # Rename .zip to .tsk
                        if (Test-Path -Path $tskFilePath) {
                            Remove-Item -Path $tskFilePath -Force
                        }
                        Rename-Item -Path $zipFilePath -NewName $tskFileName
                        Write-Log "  Created $tskFileName archive" "Green"
                    }
                    else {
                        # Alternative ZIP method for older PowerShell versions
                        Add-Type -AssemblyName System.IO.Compression.FileSystem
                        [System.IO.Compression.ZipFile]::CreateFromDirectory($tempFolder, $zipFilePath)
                        
                        # Rename .zip to .tsk
                        if (Test-Path -Path $tskFilePath) {
                            Remove-Item -Path $tskFilePath -Force
                        }
                        Rename-Item -Path $zipFilePath -NewName $tskFileName
                        Write-Log "  Created $tskFileName archive" "Green"
                    }
                    
                    # We only want to keep the .tsk file, individual files are not copied to the output folder
                    Write-Log "  Only keeping the .tsk file in the output folder" "White"
                }
                catch {
                    Write-Log "  Error processing file: $_" "Red"
                }
                finally {
                    # Clean up temporary folder
                    if (Test-Path -Path $tempFolder) {
                        Remove-Item -Path $tempFolder -Recurse -Force
                    }
                }
            }
            else {
                Write-Log "  Could not extract GS number from filename: $($excelFile.Name)" "Yellow"
            }
            
            $processedFiles++
            $progressPercentage = [int](($processedFiles / $totalFiles) * 100)
            Update-ProgressBar $progressPercentage
        }
        
        # Create Dymo labels if a template was provided
        if (-not [string]::IsNullOrWhiteSpace($DymoTemplate) -and (Test-Path -Path $DymoTemplate)) {
            Write-Log "Starting Dymo label creation process..." "Cyan"
            Update-ProgressBar 0
            
            # Read the Dymo template content
            $dymoTemplateContent = Get-Content -Path $DymoTemplate -Raw
            
            # Process each Excel file for Dymo labels
            $processedFiles = 0
            foreach ($excelFile in $excelFiles) {
                Write-Log "Creating Dymo label for: $($excelFile.Name)" "White"
                
                try {
                    # Import Excel data
                    $excelData = Import-Excel -Path $excelFile.FullName -ErrorAction Stop
                    
                    if ($excelData.Count -eq 0) {
                        Write-Log "  No data found in Excel file" "Yellow"
                        continue
                    }
                    
                    # Create XML structure for the spreadsheet data
                    $xmlDoc = New-Object System.Xml.XmlDocument
                    $dataTable = $xmlDoc.CreateElement("DataTable")
                    
                    # Define Columns
                    $columns = $xmlDoc.CreateElement("Columns")
                    $dataTable.AppendChild($columns) | Out-Null
                    
                    # Get column names from the first row
                    $firstRow = $excelData[0]
                    $columnNames = $firstRow.PSObject.Properties.Name
                    
                    foreach ($colName in $columnNames) {
                        $colElem = $xmlDoc.CreateElement("DataColumn")
                        $colElem.InnerText = $colName
                        $columns.AppendChild($colElem) | Out-Null
                    }
                    
                    # Define Rows
                    $rows = $xmlDoc.CreateElement("Rows")
                    $dataTable.AppendChild($rows) | Out-Null
                    
                    foreach ($row in $excelData) {
                        $rowElem = $xmlDoc.CreateElement("DataRow")
                        $rows.AppendChild($rowElem) | Out-Null
                        
                        foreach ($colName in $columnNames) {
                            $valueElem = $xmlDoc.CreateElement("Value")
                            $value = $row.$colName
                            
                            # Convert to string and remove .0 from numeric values
                            if ($value -is [double] -and $value -eq [math]::Floor($value)) {
                                $value = [math]::Floor($value).ToString()
                            } else {
                                $value = $value.ToString()
                            }
                            
                            $valueElem.InnerText = $value
                            $rowElem.AppendChild($valueElem) | Out-Null
                        }
                    }
                    
                    # Convert to string and format
                    $xmlSettings = New-Object System.Xml.XmlWriterSettings
                    $xmlSettings.Indent = $true
                    $xmlSettings.IndentChars = "    "
                    $xmlSettings.NewLineChars = "`n"
                    $xmlSettings.NewLineHandling = [System.Xml.NewLineHandling]::Replace
                    $xmlSettings.OmitXmlDeclaration = $true
                    
                    $stringBuilder = New-Object System.Text.StringBuilder
                    $xmlWriter = [System.Xml.XmlWriter]::Create($stringBuilder, $xmlSettings)
                    $dataTable.WriteTo($xmlWriter)
                    $xmlWriter.Flush()
                    $xmlWriter.Close()
                    
                    $dataTableXml = $stringBuilder.ToString()
                    
                    # Insert the data into the template
                    $outputContent = $dymoTemplateContent -replace "</DesktopLabel>", "$dataTableXml`n</DesktopLabel>"
                    
                    # Save as .dymo file
                    $dymoFileName = [System.IO.Path]::GetFileNameWithoutExtension($excelFile.Name) + ".dymo"
                    $dymoFilePath = Join-Path -Path $OutputFolder -ChildPath $dymoFileName
                    
                    $outputContent | Out-File -FilePath $dymoFilePath -Encoding utf8
                    
                    Write-Log "  Created $dymoFileName" "Green"
                } catch {
                    Write-Log "  Error creating Dymo label: $_" "Red"
                }
                
                $processedFiles++
                $progressPercentage = [int](($processedFiles / $totalFiles) * 100)
                Update-ProgressBar $progressPercentage
            }
            
            Write-Log "Dymo label creation process completed. Labels saved to: $OutputFolder" "Cyan"
        }
        
        Write-Log "Label creation process completed. Labels saved to: $OutputFolder" "Cyan"
        Update-ProgressBar 100
        
        # Update the button state to "Finished!" if it exists
        if ($script:currentLabelButton) {
            $script:currentLabelButton.Text = "Finished!"
            $script:currentLabelButton.BackColor = [System.Drawing.Color]::FromArgb(40, 167, 69) # Green
            $toolTip.SetToolTip($script:currentLabelButton, "Label creation completed. Click to reset.")
        }
        
        return $true
    }
    catch {
        Write-Log "Error during label creation: $_" "Red"
        
        # Update the button state to show error if it exists
        if ($script:currentLabelButton) {
            $script:currentLabelButton.Text = "Error!"
            $script:currentLabelButton.BackColor = [System.Drawing.Color]::FromArgb(220, 53, 69) # Red
            $toolTip.SetToolTip($script:currentLabelButton, "Error during label creation. Click to reset.")
        }
        
        return $false
    }
}

# Function to show the Create Labels dialog
function Show-CreateLabelsDialog {
    # Create the form
    $labelsForm = New-Object System.Windows.Forms.Form
    $labelsForm.Text = "Create Labels"
    $labelsForm.Size = New-Object System.Drawing.Size(600, 370)
    $labelsForm.StartPosition = "CenterScreen"
    $labelsForm.FormBorderStyle = "FixedDialog"
    $labelsForm.MaximizeBox = $false
    $labelsForm.MinimizeBox = $false
    
    # Form is ready to display
    
    # Create the layout panel
    $labelsPanel = New-Object System.Windows.Forms.TableLayoutPanel
    $labelsPanel.Dock = "Fill"
    $labelsPanel.RowCount = 6
    $labelsPanel.ColumnCount = 3
    $labelsPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 120)))
    $labelsPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    $labelsPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 80)))
    $labelsPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 40)))
    $labelsPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 40)))
    $labelsPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 40)))
    $labelsPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 40)))
    $labelsPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 40)))
    $labelsPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 40)))
    $labelsPanel.Padding = New-Object System.Windows.Forms.Padding(10)
    $labelsForm.Controls.Add($labelsPanel)
    
    # Input folder label
    $inputFolderLabel = New-Object System.Windows.Forms.Label
    $inputFolderLabel.Text = "Input Folder:"
    $inputFolderLabel.Dock = "Fill"
    $inputFolderLabel.TextAlign = "MiddleLeft"
    $labelsPanel.Controls.Add($inputFolderLabel, 0, 0)
    
    # Input folder textbox
    $inputFolderTextBox = New-Object System.Windows.Forms.TextBox
    $inputFolderTextBox.Dock = "Fill"
    $inputFolderTextBox.ReadOnly = $true
    $labelsPanel.Controls.Add($inputFolderTextBox, 1, 0)
    
    # Immediately set the text value if we have a saved path
    if (-not [string]::IsNullOrWhiteSpace($script:LabelInputFolder)) {
        $inputFolderTextBox.Text = $script:LabelInputFolder
    }
    
    # Create tooltip for input folder
    $toolTip = New-Object System.Windows.Forms.ToolTip
    $toolTip.SetToolTip($inputFolderTextBox, "Select the folder containing your GSxx.xlsx files to process.")
    $toolTip.SetToolTip($inputFolderLabel, "Select the folder containing your GSxx.xlsx files to process.")
    
    # Use Final Output location as a fallback for Input Folder if needed
    if ([string]::IsNullOrWhiteSpace($inputFolderTextBox.Text) -and -not [string]::IsNullOrWhiteSpace($finalOutputLocation.Text)) {
        $inputFolderTextBox.Text = $finalOutputLocation.Text
    }
    
    # Input folder browse button
    $inputFolderButton = New-Object System.Windows.Forms.Button
    $inputFolderButton.Text = "Browse..."
    $inputFolderButton.Dock = "Fill"
    $inputFolderButton.Add_Click({
        $folderPath = Select-FolderDialog
        if ($folderPath) {
            $inputFolderTextBox.Text = $folderPath
        }
    })
    $labelsPanel.Controls.Add($inputFolderButton, 2, 0)
    
    # Output folder label
    $outputFolderLabel = New-Object System.Windows.Forms.Label
    $outputFolderLabel.Text = "Output Folder:"
    $outputFolderLabel.Dock = "Fill"
    $outputFolderLabel.TextAlign = "MiddleLeft"
    $labelsPanel.Controls.Add($outputFolderLabel, 0, 1)
    
    # Output folder textbox
    $outputFolderTextBox = New-Object System.Windows.Forms.TextBox
    $outputFolderTextBox.Dock = "Fill"
    $outputFolderTextBox.ReadOnly = $true
    $labelsPanel.Controls.Add($outputFolderTextBox, 1, 1)
    
    # Immediately set the text value if we have a saved path
    if (-not [string]::IsNullOrWhiteSpace($script:LabelOutputFolder)) {
        $outputFolderTextBox.Text = $script:LabelOutputFolder
    }
    
    # Tooltip for output folder
    $toolTip.SetToolTip($outputFolderTextBox, "Select the folder where the GSxx.tsk files will be saved.")
    $toolTip.SetToolTip($outputFolderLabel, "Select the folder where the GSxx.tsk files will be saved.")
    
    # Output folder browse button
    $outputFolderButton = New-Object System.Windows.Forms.Button
    $outputFolderButton.Text = "Browse..."
    $outputFolderButton.Dock = "Fill"
    $outputFolderButton.Add_Click({
        $folderPath = Select-FolderDialog
        if ($folderPath) {
            $outputFolderTextBox.Text = $folderPath
        }
    })
    $labelsPanel.Controls.Add($outputFolderButton, 2, 1)
    
    # Param Template label
    $paramTemplateLabel = New-Object System.Windows.Forms.Label
    $paramTemplateLabel.Text = "Param Template:"
    $paramTemplateLabel.Dock = "Fill"
    $paramTemplateLabel.TextAlign = "MiddleLeft"
    $labelsPanel.Controls.Add($paramTemplateLabel, 0, 2)
    
    # Param Template textbox
    $paramTemplateTextBox = New-Object System.Windows.Forms.TextBox
    $paramTemplateTextBox.Dock = "Fill"
    $paramTemplateTextBox.ReadOnly = $true
    $labelsPanel.Controls.Add($paramTemplateTextBox, 1, 2)
    
    # Immediately set the text value if we have a saved path
    if (-not [string]::IsNullOrWhiteSpace($script:LabelParamTemplate)) {
        $paramTemplateTextBox.Text = $script:LabelParamTemplate
    }
    
    # Tooltip for param template
    $toolTip.SetToolTip($paramTemplateTextBox, "Select a .param template file for printer configuration settings.")
    $toolTip.SetToolTip($paramTemplateLabel, "Select a .param template file for printer configuration settings.")
    
    # Param Template browse button
    $paramTemplateButton = New-Object System.Windows.Forms.Button
    $paramTemplateButton.Text = "Browse..."
    $paramTemplateButton.Dock = "Fill"
    $paramTemplateButton.Add_Click({
        $filePath = Select-FileDialog -Filter "Param Files (*.param)|*.param|All Files (*.*)|*.*"
        if ($filePath) {
            $paramTemplateTextBox.Text = $filePath
        }
    })
    $labelsPanel.Controls.Add($paramTemplateButton, 2, 2)
    
    # PRT Template label
    $prtTemplateLabel = New-Object System.Windows.Forms.Label
    $prtTemplateLabel.Text = "PRT Template:"
    $prtTemplateLabel.Dock = "Fill"
    $prtTemplateLabel.TextAlign = "MiddleLeft"
    $labelsPanel.Controls.Add($prtTemplateLabel, 0, 3)
    
    # PRT Template textbox
    $prtTemplateTextBox = New-Object System.Windows.Forms.TextBox
    $prtTemplateTextBox.Dock = "Fill"
    $prtTemplateTextBox.ReadOnly = $true
    $labelsPanel.Controls.Add($prtTemplateTextBox, 1, 3)
    
    # Immediately set the text value if we have a saved path
    if (-not [string]::IsNullOrWhiteSpace($script:LabelPrtTemplate)) {
        $prtTemplateTextBox.Text = $script:LabelPrtTemplate
    }
    
    # Tooltip for PRT template
    $toolTip.SetToolTip($prtTemplateTextBox, "Select a .prt template file that contains the label layout with markers for 'Change Card Name Data Here', etc.")
    $toolTip.SetToolTip($prtTemplateLabel, "Select a .prt template file that contains the label layout with markers for 'Change Card Name Data Here', etc.")
    
    # PRT Template browse button
    $prtTemplateButton = New-Object System.Windows.Forms.Button
    $prtTemplateButton.Text = "Browse..."
    $prtTemplateButton.Dock = "Fill"
    $prtTemplateButton.Add_Click({
        $filePath = Select-FileDialog -Filter "PRT Files (*.prt)|*.prt|All Files (*.*)|*.*"
        if ($filePath) {
            $prtTemplateTextBox.Text = $filePath
        }
    })
    $labelsPanel.Controls.Add($prtTemplateButton, 2, 3)
    
    # Dymo Template label
    $dymoTemplateLabel = New-Object System.Windows.Forms.Label
    $dymoTemplateLabel.Text = "Dymo Template:"
    $dymoTemplateLabel.Dock = "Fill"
    $dymoTemplateLabel.TextAlign = "MiddleLeft"
    $labelsPanel.Controls.Add($dymoTemplateLabel, 0, 4)
    
    # Dymo Template textbox
    $dymoTemplateTextBox = New-Object System.Windows.Forms.TextBox
    $dymoTemplateTextBox.Dock = "Fill"
    $dymoTemplateTextBox.ReadOnly = $true
    $labelsPanel.Controls.Add($dymoTemplateTextBox, 1, 4)
    
    # Immediately set the text value if we have a saved path
    if (-not [string]::IsNullOrWhiteSpace($script:LabelDymoTemplate)) {
        $dymoTemplateTextBox.Text = $script:LabelDymoTemplate
    }
    
    # Tooltip for Dymo template
    $toolTip.SetToolTip($dymoTemplateTextBox, "Select a Dymo template file (optional - for future functionality).")
    $toolTip.SetToolTip($dymoTemplateLabel, "Select a Dymo template file (optional - for future functionality).")
    
    # Dymo Template browse button
    $dymoTemplateButton = New-Object System.Windows.Forms.Button
    $dymoTemplateButton.Text = "Browse..."
    $dymoTemplateButton.Dock = "Fill"
    $dymoTemplateButton.Add_Click({
        $filePath = Select-FileDialog -Filter "Dymo Files (*.xml)|*.xml|All Files (*.*)|*.*"
        if ($filePath) {
            $dymoTemplateTextBox.Text = $filePath
        }
    })
    $labelsPanel.Controls.Add($dymoTemplateButton, 2, 4)
    
    # Button panel
    $buttonPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $buttonPanel.Dock = "Bottom"
    $buttonPanel.FlowDirection = "RightToLeft"
    $buttonPanel.WrapContents = $false
    $buttonPanel.Height = 40
    $buttonPanel.Padding = New-Object System.Windows.Forms.Padding(0, 5, 0, 0)
    $labelsPanel.Controls.Add($buttonPanel, 0, 5)
    $labelsPanel.SetColumnSpan($buttonPanel, 3)
    
    # Create button
    $createButton = New-Object System.Windows.Forms.Button
    $createButton.Text = "Create Labels"
    $createButton.Width = 120
    $createButton.Height = 30
    $createButton.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215) # Blue
    $createButton.ForeColor = [System.Drawing.Color]::White
    $createButton.FlatStyle = "Flat"
    $createButton.Margin = New-Object System.Windows.Forms.Padding(5, 0, 0, 0)
    $toolTip.SetToolTip($createButton, "Start the label creation process")
    $createButton.Add_Click({
        # Check if this is a reset from "Finished!" state
        if ($createButton.Text -eq "Finished!") {
            # Reset button appearance
            $createButton.Text = "Create Labels"
            $createButton.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215) # Blue
            $toolTip.SetToolTip($createButton, "Start the label creation process")
            return
        }
        
        # Validate inputs
        if ([string]::IsNullOrWhiteSpace($inputFolderTextBox.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Please select an input folder.", "Input Required", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        
        if ([string]::IsNullOrWhiteSpace($outputFolderTextBox.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Please select an output folder.", "Input Required", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        
        # Close the dialog
        $labelsForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $labelsForm.Close()
        
        # Create a global reference to the button so we can update it from the Create-Labels function
        $script:currentLabelButton = $createButton
        
        # Change button appearance to "Creating..."
        $createButton.Text = "Creating..."
        $createButton.BackColor = [System.Drawing.Color]::FromArgb(255, 204, 0) # Yellow/Orange
        $toolTip.SetToolTip($createButton, "Label creation in progress...")
        $form.Refresh() # Force UI update
        
        # Start the label creation process
        Create-Labels -InputFolder $inputFolderTextBox.Text -OutputFolder $outputFolderTextBox.Text -ParamTemplate $paramTemplateTextBox.Text -PrtTemplate $prtTemplateTextBox.Text -DymoTemplate $dymoTemplateTextBox.Text
    })
    $buttonPanel.Controls.Add($createButton)
    
    # Cancel button
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Cancel"
    $cancelButton.Width = 80
    $cancelButton.Height = 30
    $cancelButton.Margin = New-Object System.Windows.Forms.Padding(5, 0, 0, 0)
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $buttonPanel.Controls.Add($cancelButton)
    
    $labelsForm.AcceptButton = $createButton
    $labelsForm.CancelButton = $cancelButton
    
    # Show the dialog
    $result = $labelsForm.ShowDialog()
    
    # Always save the paths regardless of how the dialog was closed
    $script:LabelInputFolder = $inputFolderTextBox.Text
    $script:LabelOutputFolder = $outputFolderTextBox.Text
    $script:LabelParamTemplate = $paramTemplateTextBox.Text
    $script:LabelPrtTemplate = $prtTemplateTextBox.Text
    $script:LabelDymoTemplate = $dymoTemplateTextBox.Text
    
    # Return true if dialog was accepted
    return ($result -eq [System.Windows.Forms.DialogResult]::OK)
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
$form.Size = New-Object System.Drawing.Size(900, 870) 
$form.MinimumSize = New-Object System.Drawing.Size(800, 750)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.MinimizeBox = $true
$form.Font = New-Object System.Drawing.Font("Segoe UI", 10)

# Set application icon if logo exists
$logoPath = Join-Path -Path $PSScriptRoot -ChildPath "logo.png"
if (Test-Path -Path $logoPath) {
    try {
        # Load the logo as an icon for the application
        $logo = [System.Drawing.Image]::FromFile($logoPath)
        # Create a simple icon from the logo
        $icon = [System.Drawing.Icon]::FromHandle(($logo.GetThumbnailImage(64, 64, $null, [System.IntPtr]::Zero)).GetHicon())
        $form.Icon = $icon
    } catch {
        Write-Log "Error setting application icon: $_" "Yellow"
    }
}

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
    Write-Log "Resetting to new configuration." "White"
    # Clear text fields and list views
    $backupLocations.Items.Clear()
    $spreadsheetLocations.Items.Clear()
    $destinationLocation.Text = ""
    $skuListLocation.Text = ""
    $finalOutputLocation.Text = ""
    $textBoxSingleSpreadsheetFile.Text = ""
    
    # Reset script-level variables for label paths
    $script:LabelInputFolder = ""
    $script:LabelOutputFolder = ""
    $script:LabelParamTemplate = ""
    $script:LabelPrtTemplate = ""
    $script:LabelDymoTemplate = ""

    # Reset all checkboxes to false
    foreach ($checkbox in $optionCheckboxes) {
        $checkbox.Checked = $false
    }
    
    # Reset current config file path
    $script:CurrentConfigFile = $null
    $form.Text = "Spreadsheet Wrangler"
    Update-RecentFilesMenu # This will also save app settings if recent files are managed
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

# Recent Files submenu
$recentFilesMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$recentFilesMenuItem.Text = "Recent Files"
$fileMenu.DropDownItems.Add($recentFilesMenuItem)

# Initialize with empty item (will be updated by Update-RecentFilesMenu)
$noRecentFilesItem = New-Object System.Windows.Forms.ToolStripMenuItem
$noRecentFilesItem.Text = "(No recent files)"
$noRecentFilesItem.Enabled = $false
$recentFilesMenuItem.DropDownItems.Add($noRecentFilesItem)

# Separator
$fileMenu.DropDownItems.Add("-")

# Exit
$exitMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$exitMenuItem.Text = "Exit"
$exitMenuItem.ShortcutKeys = [System.Windows.Forms.Keys]::Alt -bor [System.Windows.Forms.Keys]::F4
$exitMenuItem.Add_Click({ $form.Close() })
$fileMenu.DropDownItems.Add($exitMenuItem)

# Labels Menu
$labelsMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$labelsMenu.Text = "Labels"

# Create Labels
$createLabelsMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$createLabelsMenuItem.Text = "Create Labels"
$createLabelsMenuItem.Add_Click({
    Show-CreateLabelsDialog
})
$labelsMenu.DropDownItems.Add($createLabelsMenuItem)

# Help Menu
$helpMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$helpMenu.Text = "Help"

# Check for Updates
$checkUpdatesMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$checkUpdatesMenuItem.Text = "Check for Updates"
$checkUpdatesMenuItem.Add_Click({
    Check-ForUpdates
})
$helpMenu.DropDownItems.Add($checkUpdatesMenuItem)

# Separator
$helpMenu.DropDownItems.Add("-")

# About
$aboutMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$aboutMenuItem.Text = "About"
$aboutMenuItem.Add_Click({
    $aboutForm = New-Object System.Windows.Forms.Form
    $aboutForm.Text = "About Spreadsheet Wrangler"
    # Adjust the size of the About dialog
    $aboutForm.Size = New-Object System.Drawing.Size(500, 400)
    $aboutForm.StartPosition = "CenterParent"
    $aboutForm.FormBorderStyle = "FixedDialog"
    $aboutForm.MaximizeBox = $false
    $aboutForm.MinimizeBox = $false
    
    $aboutPanel = New-Object System.Windows.Forms.TableLayoutPanel
    $aboutPanel.Dock = "Fill"
    $aboutPanel.RowCount = 4
    $aboutPanel.ColumnCount = 1
    # Adjust space allocation to reduce gaps
    $aboutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 40)))
    $aboutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 30)))
    $aboutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 15)))
    $aboutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 15)))
    $aboutForm.Controls.Add($aboutPanel)
    
    # Logo
    $logoPanel = New-Object System.Windows.Forms.Panel
    $logoPanel.Dock = "Fill"
    # Reduce padding to 10 pixels
    $logoPanel.Padding = New-Object System.Windows.Forms.Padding(10)
    $aboutPanel.Controls.Add($logoPanel, 0, 0)
    
    # Load the logo image
    $logoPath = Join-Path -Path $PSScriptRoot -ChildPath "logo.png"
    if (Test-Path -Path $logoPath) {
        try {
            $logoImage = [System.Drawing.Image]::FromFile($logoPath)
            $logoPictureBox = New-Object System.Windows.Forms.PictureBox
            $logoPictureBox.Image = $logoImage
            $logoPictureBox.SizeMode = "Zoom"
            $logoPictureBox.Dock = "Fill"
            $logoPanel.Controls.Add($logoPictureBox)
        } catch {
            Write-Log "Error loading logo: $_" "Yellow"
        }
    }
    
    # Main about text
    $aboutLabel = New-Object System.Windows.Forms.Label
    $aboutLabel.Text = "Spreadsheet Wrangler v1.8.5`n`nA powerful tool for backing up folders and combining spreadsheets.`n`nCreated by Bryant Welch`nCreated: $(Get-Date -Format 'yyyy-MM-dd')`n`n(c) 2025 Bryant Welch. All Rights Reserved"
    $aboutLabel.AutoSize = $false
    $aboutLabel.Dock = "Fill"
    $aboutLabel.TextAlign = "MiddleCenter"
    $aboutPanel.Controls.Add($aboutLabel, 0, 1)
    
    # GitHub link
    $linkLabel = New-Object System.Windows.Forms.LinkLabel
    $linkLabel.Text = "https://github.com/BryantWelch/Spreadsheet-Wrangler"
    $linkLabel.AutoSize = $false
    $linkLabel.Dock = "Fill"
    $linkLabel.TextAlign = "MiddleCenter"
    $linkLabel.LinkColor = [System.Drawing.Color]::Blue
    $linkLabel.ActiveLinkColor = [System.Drawing.Color]::Red
    $linkLabel.Add_LinkClicked({
        param($senderObj, $e)
        Start-Process "https://github.com/BryantWelch/Spreadsheet-Wrangler"
    })
    $aboutPanel.Controls.Add($linkLabel, 0, 2)
    
    # OK button
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "OK"
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $okButton.Dock = "Fill"
    $okButton.Margin = New-Object System.Windows.Forms.Padding(150, 10, 150, 10)
    $aboutPanel.Controls.Add($okButton, 0, 3)
    $aboutForm.AcceptButton = $okButton
    
    $aboutForm.ShowDialog() | Out-Null
})
$helpMenu.DropDownItems.Add($aboutMenuItem)

# Add menus to menu bar
$menuBar.Items.Add($fileMenu)
$menuBar.Items.Add($labelsMenu)
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
$leftLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 15))) # Backup - adjusted for 1-2 paths
$leftLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 25))) # Spreadsheet - adjusted for 3-4 paths
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
$backupLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 80)))
$backupLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 35)))
$backupPanel.Controls.Add($backupLayout)

# List of backup locations
$backupLocations = New-Object System.Windows.Forms.ListView
$backupLocations.View = "Details"
$backupLocations.FullRowSelect = $true
$backupLocations.HorizontalScrollbar = $true
$backupLocations.Scrollable = $true
$backupLocations.HeaderStyle = "None" # Remove header to save space
$backupLocations.Columns.Add("Folder Path", 400) # Set explicit width to force horizontal scrollbar
$backupLocations.Dock = "Fill"
$backupLocations.MinimumSize = New-Object System.Drawing.Size(0, 50) # Increased minimum height

# Do not auto-resize columns to ensure horizontal scrolling works
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
$spreadsheetLocations.HorizontalScrollbar = $true
$spreadsheetLocations.Scrollable = $true
$spreadsheetLocations.HeaderStyle = "None" # Remove header to save space
$spreadsheetLocations.Columns.Add("Folder Path", 400) # Set explicit width to force horizontal scrollbar
$spreadsheetLocations.Dock = "Fill"
$spreadsheetLocations.MinimumSize = New-Object System.Drawing.Size(0, 90) # Maintain minimum height for 3-4 rows

# Do not auto-resize columns to ensure horizontal scrolling works
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

# --- START: UI Elements for Single Spreadsheet Option ---
# Textbox for single spreadsheet file path
$textBoxSingleSpreadsheetFile = New-Object System.Windows.Forms.TextBox
$textBoxSingleSpreadsheetFile.ReadOnly = $true
$textBoxSingleSpreadsheetFile.Dock = [System.Windows.Forms.DockStyle]::Fill
$textBoxSingleSpreadsheetFile.BackColor = [System.Drawing.Color]::White
$toolTip.SetToolTip($textBoxSingleSpreadsheetFile, "Path to the single spreadsheet file to process")
$textBoxSingleSpreadsheetFile.Visible = $false # Initially hidden
$spreadsheetLayout.Controls.Add($textBoxSingleSpreadsheetFile, 0, 0) # Added to layout, will toggle visibility with $spreadsheetLocations

# Browse button for single spreadsheet file
$buttonBrowseSingleSpreadsheetFile = New-Object System.Windows.Forms.Button
$buttonBrowseSingleSpreadsheetFile.Text = "Browse..."
$buttonBrowseSingleSpreadsheetFile.Width = 80 # Standard browse button width
$buttonBrowseSingleSpreadsheetFile.Height = 25
$buttonBrowseSingleSpreadsheetFile.Margin = New-Object System.Windows.Forms.Padding(3, 0, 3, 0) # Consistent margin
$buttonBrowseSingleSpreadsheetFile.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$toolTip.SetToolTip($buttonBrowseSingleSpreadsheetFile, "Select the single spreadsheet file to process")
$buttonBrowseSingleSpreadsheetFile.Visible = $false # Initially hidden
$buttonBrowseSingleSpreadsheetFile.Add_Click({
    $filePath = Select-FileDialog -Filter "Spreadsheet Files (*.xlsx;*.xls;*.csv)|*.xlsx;*.xls;*.csv|All files (*.*)|*.*" -Title "Select Single Spreadsheet File"
    if ($filePath) {
        $textBoxSingleSpreadsheetFile.Text = $filePath
        Write-Log "Selected single spreadsheet: $filePath"
    }
})
$spreadsheetButtonPanel.Controls.Add($buttonBrowseSingleSpreadsheetFile) # Add to existing button panel
# --- END: UI Elements for Single Spreadsheet Option ---

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
    # Check if this is a reset from "Finished!" state
    if ($runBtn.Text -eq "Finished!") {
        # Reset button appearance
        $runBtn.Text = "Run"
        $runBtn.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215) # Blue
        $toolTip.SetToolTip($runBtn, "Start the backup and spreadsheet combining process")
        return
    }
    
    # Change button appearance to "Running..."
    $runBtn.Text = "Running..."
    $runBtn.BackColor = [System.Drawing.Color]::FromArgb(255, 204, 0) # Yellow/Orange
    $toolTip.SetToolTip($runBtn, "Processing in progress...")
    $form.Refresh() # Force UI update
    
    # Clear previous output
    $outputTextbox.Clear()
    
    # Initialize log file if logging is enabled
    if ($optionCheckboxes[8].Checked) { # Log to File option
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
    
    # Start spreadsheet combining process if not skipped
    $combineSuccess = $true
    if (-not $optionCheckboxes[1].Checked) {
        $combineSuccess = Start-SpreadsheetCombiningProcess
    } else {
        Write-Log "Spreadsheet combining process skipped due to 'Skip Combine' option." "Yellow"
    }
    
    # Process SKU list if spreadsheet combining was successful (or skipped) and SKU list path is provided
    # When combining is skipped, we still need to have a valid destination location
    # Check if SKU processing should be skipped (checkboxSkipSkuProcessing is $optionCheckboxes[10] as it's the 11th checkbox, 0-indexed)
    $skipSkuProcessing = $optionCheckboxes[2].Checked

    if (-not $skipSkuProcessing -and $combineSuccess -and -not [string]::IsNullOrWhiteSpace($skuListLocation.Text) -and 
        -not [string]::IsNullOrWhiteSpace($finalOutputLocation.Text) -and 
        (-not $optionCheckboxes[1].Checked -or -not [string]::IsNullOrWhiteSpace($destinationLocation.Text))) {
        Write-Log "Starting SKU list processing..." "Cyan"
        $skuListSuccess = Process-SKUList -CombinedSpreadsheetPath $destinationLocation.Text -SKUListPath $skuListLocation.Text -FinalOutputPath $finalOutputLocation.Text
        
        if ($skuListSuccess) {
            Write-Log "SKU list processing completed successfully." "Green"
        } else {
            Write-Log "SKU list processing completed with errors." "Red"
        }
    } elseif ($skipSkuProcessing) {
        Write-Log "SKU list processing skipped due to 'Skip Sku Processing' option." "Yellow"
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
    
    # Change button appearance to "Finished!"
    $runBtn.Text = "Finished!"
    $runBtn.BackColor = [System.Drawing.Color]::FromArgb(40, 167, 69) # Green
    $toolTip.SetToolTip($runBtn, "Operations completed. Click to reset.")
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
$optionsLayout.RowCount = 4 # Changed from 3
$optionsLayout.ColumnCount = 3 # Changed from 4
$optionsLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 25))) # Changed from 33.33
$optionsLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 25))) # Changed from 33.33
$optionsLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 25))) # Changed from 33.33
$optionsLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 25))) # Added fourth row
$optionsLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 33.33))) # Changed from 25
$optionsLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 33.33))) # Changed from 25
$optionsLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 33.33))) # Changed from 25
$optionsPanel.Controls.Add($optionsLayout)

# Create checkboxes for options with specific functionality
$optionCheckboxes = @()

# Option 1: Skip backup process
$optionCheckboxes += $checkbox1 = New-Object System.Windows.Forms.CheckBox
$checkbox1.Name = "checkboxSkipBackup"
$checkbox1.Text = "Skip Backup"
$checkbox1.Dock = "Fill"
$checkbox1.Margin = New-Object System.Windows.Forms.Padding(5)
$toolTip.SetToolTip($checkbox1, "Skip the backup process and only combine spreadsheets")
$optionsLayout.Controls.Add($checkbox1, 0, 0)

# Option 2: Skip Combine - Skip the spreadsheet combining process
$optionCheckboxes += $checkbox2 = New-Object System.Windows.Forms.CheckBox
$checkbox2.Name = "checkboxSkipCombine"
$checkbox2.Text = "Skip Combine"
$checkbox2.Dock = "Fill"
$checkbox2.Margin = New-Object System.Windows.Forms.Padding(5)
$toolTip.SetToolTip($checkbox2, "Skip the spreadsheet combining process and only perform backup and/or SKU list processing")
$optionsLayout.Controls.Add($checkbox2, 1, 0)

# Option 3: Exclude headers
$optionCheckboxes += $checkbox3 = New-Object System.Windows.Forms.CheckBox
$checkbox3.Name = "checkboxSkipSkuProcessing"
$checkbox3.Text = "Skip Sku Processing"
$checkbox3.Dock = "Fill"
$checkbox3.Margin = New-Object System.Windows.Forms.Padding(5)
$toolTip.SetToolTip($checkbox3, "If checked, the SKU list processing step will be skipped.")
$optionsLayout.Controls.Add($checkbox3, 2, 0)

# Option 4: Duplicate rows based on value in 'Add to Quantity' column
$optionCheckboxes += $checkbox4 = New-Object System.Windows.Forms.CheckBox
$checkbox4.Name = "checkboxDuplicateByQty"
$checkbox4.Text = "Duplicate by Qty"
$checkbox4.Dock = "Fill"
$checkbox4.Margin = New-Object System.Windows.Forms.Padding(5)
$toolTip.SetToolTip($checkbox4, "Duplicate rows based on the numeric value in the 'Add to Quantity' column. For example, a value of 7 will create 7 total rows.")
$optionsLayout.Controls.Add($checkbox4, 0, 1)

# Option 5: Normalize all quantities to '1'
$optionCheckboxes += $checkbox5 = New-Object System.Windows.Forms.CheckBox
$checkbox5.Name = "checkboxNormalizeQty"
$checkbox5.Text = "Normalize Qty to 1"
$checkbox5.Dock = "Fill"
$checkbox5.Margin = New-Object System.Windows.Forms.Padding(5)
$toolTip.SetToolTip($checkbox5, "Change all values in 'Add to Quantity' column to '1' (runs after duplication)")
$optionsLayout.Controls.Add($checkbox5, 1, 1)

# Option 6: Support multiple file formats
$optionCheckboxes += $checkbox6 = New-Object System.Windows.Forms.CheckBox
$checkbox6.Name = "checkboxAllFormats"
$checkbox6.Text = "All Formats"
$checkbox6.Dock = "Fill"
$checkbox6.Margin = New-Object System.Windows.Forms.Padding(5)
$toolTip.SetToolTip($checkbox6, "Process all spreadsheet formats (.xlsx, .xls, .csv)")
$optionsLayout.Controls.Add($checkbox6, 2, 1)

# Option 7: BLANK - Insert separator rows between spreadsheets
$optionCheckboxes += $checkbox7 = New-Object System.Windows.Forms.CheckBox
$checkbox7.Name = "checkboxBlankSeparator"
$checkbox7.Text = "BLANK"
$checkbox7.Dock = "Fill"
$checkbox7.Margin = New-Object System.Windows.Forms.Padding(5)
$toolTip.SetToolTip($checkbox7, "Insert 'BLANK' rows between data from different spreadsheets")
$optionsLayout.Controls.Add($checkbox7, 0, 2)

# Option 8: Reverse, Reverse - Reverse the order of data rows
$optionCheckboxes += $checkbox8 = New-Object System.Windows.Forms.CheckBox
$checkbox8.Name = "checkboxReverseRows"
$checkbox8.Text = "Reverse, Reverse"
$checkbox8.Dock = "Fill"
$checkbox8.Margin = New-Object System.Windows.Forms.Padding(5)
$toolTip.SetToolTip($checkbox8, "Reverse the order of data rows in the final combined spreadsheet")
$optionsLayout.Controls.Add($checkbox8, 1, 2)

# Option 9: Log to File
$optionCheckboxes += $checkbox9 = New-Object System.Windows.Forms.CheckBox
$checkbox9.Name = "checkboxLogToFile"
$checkbox9.Text = "Log to File"
$checkbox9.Dock = "Fill"
$checkbox9.Margin = New-Object System.Windows.Forms.Padding(5)
$toolTip.SetToolTip($checkbox9, "Save terminal output to a log file in the application directory")
$optionsLayout.Controls.Add($checkbox9, 2, 2)

# Option 10: Single Spreadsheet Duplication
$optionCheckboxes += ($checkbox10 = New-Object System.Windows.Forms.CheckBox)
$checkbox10.Name = "checkboxSingleSpreadsheet"
$checkbox10.Text = "Single Spreadsheet"
$checkbox10.Dock = "Fill"
$checkbox10.Margin = New-Object System.Windows.Forms.Padding(5)
$toolTip.SetToolTip($checkbox10, "Enable Single Spreadsheet mode. This changes the UI for spreadsheet selection and disables the 'No Headers' option.")
$optionsLayout.Controls.Add($checkbox10, 0, 3)

# Option 11: No Headers
$optionCheckboxes += ($checkboxSkipSkuProcessing = New-Object System.Windows.Forms.CheckBox)
$checkboxSkipSkuProcessing.Name = "checkboxNoHeaders"
$checkboxSkipSkuProcessing.Text = "No Headers"
$checkboxSkipSkuProcessing.Dock = "Fill"
$checkboxSkipSkuProcessing.Margin = New-Object System.Windows.Forms.Padding(5)
$toolTip.SetToolTip($checkboxSkipSkuProcessing, "Exclude headers when combining spreadsheets")
$optionsLayout.Controls.Add($checkboxSkipSkuProcessing, 1, 3)

# Option 12: Reserved for future use
$optionCheckboxes += $checkbox12 = New-Object System.Windows.Forms.CheckBox
$checkbox12.Name = "checkboxOption12"
$checkbox12.Text = "Option 12"
$checkbox12.Dock = "Fill"
$checkbox12.Margin = New-Object System.Windows.Forms.Padding(5)
$toolTip.SetToolTip($checkbox12, "Reserved for future use")
$optionsLayout.Controls.Add($checkbox12, 2, 3)

# Event handler for Single Spreadsheet checkbox ($checkbox10, which is $optionCheckboxes[9])
$checkbox10.Add_CheckedChanged({
    param($sender, $e)
    if ($sender.Checked) {
        # Single Spreadsheet mode ENABLED
        $spreadsheetPanel.Text = "Single Spreadsheet Location:"
        
        $spreadsheetLocations.Visible = $false
        $addSpreadsheetBtn.Visible = $false
        $removeSpreadsheetBtn.Visible = $false
        
        $textBoxSingleSpreadsheetFile.Visible = $true
        $buttonBrowseSingleSpreadsheetFile.Visible = $true
        
        # Disable and uncheck "No Headers" ($optionCheckboxes[10] is $checkboxSkipSkuProcessing)
        $optionCheckboxes[10].Enabled = $false
        $optionCheckboxes[10].Checked = $false
        $toolTip.SetToolTip($optionCheckboxes[10], "No Headers option is not applicable with Single Spreadsheet mode.")

        # Disable and uncheck "BLANK" ($optionCheckboxes[6] is $checkbox7)
        $optionCheckboxes[6].Enabled = $false
        $optionCheckboxes[6].Checked = $false
        $toolTip.SetToolTip($optionCheckboxes[6], "BLANK option is not applicable with Single Spreadsheet mode.")

    } else {
        # Single Spreadsheet mode DISABLED (Multi-folder mode)
        $spreadsheetPanel.Text = "Spreadsheet Folder Locations:"
        
        $spreadsheetLocations.Visible = $true
        $addSpreadsheetBtn.Visible = $true
        $removeSpreadsheetBtn.Visible = $true
        
        $textBoxSingleSpreadsheetFile.Visible = $false
        $buttonBrowseSingleSpreadsheetFile.Visible = $false
        
        # Enable "No Headers" ($optionCheckboxes[10] is $checkboxSkipSkuProcessing)
        $optionCheckboxes[10].Enabled = $true
        $originalNoHeadersTooltip = "Exclude headers when combining spreadsheets"
        $toolTip.SetToolTip($optionCheckboxes[10], $originalNoHeadersTooltip)

        # Enable "BLANK" ($optionCheckboxes[6] is $checkbox7)
        $optionCheckboxes[6].Enabled = $true
        $originalBlankTooltip = "Insert 'BLANK' rows between data from different spreadsheets"
        $toolTip.SetToolTip($optionCheckboxes[6], $originalBlankTooltip)
    }
})

# Note: The logic to initialize the UI based on $checkbox10.Checked state 
# (e.g., when loading a configuration) will be handled within the $form.Add_Load event handler.
# This ensures all controls are created and the form is ready before attempting to modify UI state. 

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
        
        # Add Label Folder and Template locations
        # Create a Labels element to group all label-related settings
        $labelsElement = $xmlDoc.CreateElement("Labels")
        $rootElement.AppendChild($labelsElement) | Out-Null
        
        # Add Label Input Folder
        $labelInputElement = $xmlDoc.CreateElement("InputFolder")
        $labelInputElement.InnerText = $script:LabelInputFolder
        $labelsElement.AppendChild($labelInputElement) | Out-Null
        
        # Add Label Output Folder
        $labelOutputElement = $xmlDoc.CreateElement("OutputFolder")
        $labelOutputElement.InnerText = $script:LabelOutputFolder
        $labelsElement.AppendChild($labelOutputElement) | Out-Null
        
        # Add Param Template
        $paramTemplateElement = $xmlDoc.CreateElement("ParamTemplate")
        $paramTemplateElement.InnerText = $script:LabelParamTemplate
        $labelsElement.AppendChild($paramTemplateElement) | Out-Null
        
        # Add PRT Template
        $prtTemplateElement = $xmlDoc.CreateElement("PrtTemplate")
        $prtTemplateElement.InnerText = $script:LabelPrtTemplate
        $labelsElement.AppendChild($prtTemplateElement) | Out-Null
        
        # Add Dymo Template
        $dymoTemplateElement = $xmlDoc.CreateElement("DymoTemplate")
        $dymoTemplateElement.InnerText = $script:LabelDymoTemplate
        $labelsElement.AppendChild($dymoTemplateElement) | Out-Null
        
        # Add options
        $optionsElement = $xmlDoc.CreateElement("Options")
        $rootElement.AppendChild($optionsElement) | Out-Null
        
        for ($i = 0; $i -lt $optionCheckboxes.Count; $i++) {
            $checkbox = $optionCheckboxes[$i]
            # Use checkbox Name if available and not empty, otherwise use index
            $elementName = if (-not [string]::IsNullOrWhiteSpace($checkbox.Name)) { $checkbox.Name } else { "Option$($i+1)" }
            $optionElement = $xmlDoc.CreateElement($elementName)
            $optionElement.InnerText = $checkbox.Checked
            $optionsElement.AppendChild($optionElement) | Out-Null
        }
        
        # Save the XML document
        $xmlDoc.Save($ConfigPath)
        
        # Add to recent files list
        Add-RecentFile -FilePath $ConfigPath
        
        # Update the Recent Files menu
        Update-RecentFilesMenu
        
        Write-Log "Configuration saved to: $ConfigPath" "Green"
        return $true
    }
    catch {
        Write-Log "Error saving configuration: $_" "Red"
        return $false
    }
}


# Function to add a file to the recent files list
function Add-RecentFile {
    param (
        [Parameter(Mandatory=$true)]
        [string]$FilePath
    )
    
    # Remove the file from the list if it already exists
    $script:RecentFiles = $script:RecentFiles | Where-Object { $_ -ne $FilePath }
    
    # Add the file to the beginning of the list
    $script:RecentFiles = @($FilePath) + $script:RecentFiles
    
    # Trim the list to the maximum number of recent files
    if ($script:RecentFiles.Count -gt $script:MaxRecentFiles) {
        $script:RecentFiles = $script:RecentFiles[0..($script:MaxRecentFiles - 1)]
    }
    
    # Save the recent files list to the settings file
    Save-AppSettings
}

# Function to save application settings
function Save-AppSettings {
    try {
        # Create XML document
        $xmlDoc = New-Object System.Xml.XmlDocument
        $xmlDeclaration = $xmlDoc.CreateXmlDeclaration("1.0", "UTF-8", $null)
        $xmlDoc.AppendChild($xmlDeclaration) | Out-Null
        
        # Create root element
        $rootElement = $xmlDoc.CreateElement("SpreadsheetWranglerSettings")
        $xmlDoc.AppendChild($rootElement) | Out-Null
        
        # Add recent files
        $recentFilesElement = $xmlDoc.CreateElement("RecentFiles")
        $rootElement.AppendChild($recentFilesElement) | Out-Null
        
        foreach ($file in $script:RecentFiles) {
            $fileElement = $xmlDoc.CreateElement("File")
            $fileElement.InnerText = $file
            $recentFilesElement.AppendChild($fileElement) | Out-Null
        }
        
        # Save the XML document
        $xmlDoc.Save($script:AppSettingsFile)
        return $true
    }
    catch {
        Write-Log "Error saving application settings: $_" "Red"
        return $false
    }
}

# Function to check for application updates
function Check-ForUpdates {
    try {
        Write-Log "Checking for updates..." "Cyan"
        
        # Current version (from the app)
        $currentVersion = "1.8.5" # This should match the version in the about dialog
        
        # Get the latest release info from GitHub API
        $apiUrl = "https://api.github.com/repos/BryantWelch/Spreadsheet-Wrangler/releases/latest"
        
        Write-Log "Connecting to GitHub..." "White"
        $response = Invoke-RestMethod -Uri $apiUrl -Headers @{
            "Accept" = "application/vnd.github.v3+json"
            "User-Agent" = "PowerShell Script"
        }
        
        # Extract version number from tag (assuming format like "v1.8.5")
        $latestVersion = $response.tag_name -replace 'v', ''
        
        Write-Log "Current version: $currentVersion" "White"
        Write-Log "Latest version: $latestVersion" "White"
        
        # Compare versions
        if ([version]$latestVersion -gt [version]$currentVersion) {
            # New version available
            Write-Log "New version available!" "Green"
            $updatePrompt = [System.Windows.Forms.MessageBox]::Show(
                "A new version ($latestVersion) is available. Would you like to update now?\n\nChanges in this version:\n$($response.body)",
                "Update Available",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
            
            if ($updatePrompt -eq [System.Windows.Forms.DialogResult]::Yes) {
                # Download and install update
                Update-Application -ReleaseInfo $response
            }
        } else {
            # No update needed
            Write-Log "You are running the latest version." "Green"
            [System.Windows.Forms.MessageBox]::Show(
                "You are running the latest version ($currentVersion).",
                "No Updates Available",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
        }
    } catch {
        $errorMessage = $_.Exception.Message
        Write-Log "Error checking for updates: $errorMessage" "Red"
        [System.Windows.Forms.MessageBox]::Show(
            "Could not check for updates. Please check your internet connection and try again.\n\nError: $errorMessage",
            "Update Check Failed",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
}

# Function to update the application
function Update-Application {
    param (
        [Parameter(Mandatory=$true)]
        $ReleaseInfo
    )
    
    try {
        # Create a progress form
        $progressForm = New-Object System.Windows.Forms.Form
        $progressForm.Text = "Updating Spreadsheet Wrangler"
        $progressForm.Size = New-Object System.Drawing.Size(400, 150)
        $progressForm.StartPosition = "CenterScreen"
        $progressForm.FormBorderStyle = "FixedDialog"
        $progressForm.MaximizeBox = $false
        $progressForm.MinimizeBox = $false
        
        $progressLabel = New-Object System.Windows.Forms.Label
        $progressLabel.Location = New-Object System.Drawing.Point(10, 20)
        $progressLabel.Size = New-Object System.Drawing.Size(380, 20)
        $progressLabel.Text = "Downloading update..."
        $progressForm.Controls.Add($progressLabel)
        
        $progressBar = New-Object System.Windows.Forms.ProgressBar
        $progressBar.Location = New-Object System.Drawing.Point(10, 50)
        $progressBar.Size = New-Object System.Drawing.Size(380, 20)
        $progressBar.Style = "Marquee"
        $progressForm.Controls.Add($progressBar)
        
        # Show the progress form
        $progressForm.Show()
        $progressForm.Refresh()
        
        # Find the main script asset
        $asset = $ReleaseInfo.assets | Where-Object { $_.name -eq "SpreadsheetWrangler.ps1" }
        
        if ($asset) {
            # Direct script file download
            $downloadUrl = $asset.browser_download_url
            Write-Log "Downloading from: $downloadUrl" "White"
            
            # Temporary file for download
            $tempFile = [System.IO.Path]::GetTempFileName()
            
            # Download the file
            Invoke-WebRequest -Uri $downloadUrl -OutFile $tempFile
            
            # Get the current script path
            $currentScriptPath = $PSCommandPath
            $scriptDirectory = [System.IO.Path]::GetDirectoryName($currentScriptPath)
            $vbsPath = [System.IO.Path]::Combine($scriptDirectory, "Launch-SpreadsheetWrangler.vbs")
            
            # Determine how to restart - use VBS if it exists, otherwise direct PowerShell
            $restartCommand = 'start powershell -NoProfile -ExecutionPolicy Bypass -File "' + $currentScriptPath + '"'
            if (Test-Path $vbsPath) {
                $restartCommand = 'start "" "' + $vbsPath + '"'
                Write-Log "Will restart using VBS launcher: $vbsPath" "White"
            } else {
                Write-Log "Will restart using PowerShell directly" "White"
            }
            
            # Create update batch file to replace the script after this process exits
            $batchFile = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "update_spreadsheet_wrangler.bat")
            
            # Batch file content to wait for process to exit, then copy new file and restart
            $batchContent = @"
@echo off
timeout /t 2 /nobreak > nul
copy /Y "$tempFile" "$currentScriptPath"
$restartCommand
del "%~f0"
exit
"@
            
            # Write batch file
            $batchContent | Out-File -FilePath $batchFile -Encoding ascii
            
            # Close the progress form
            $progressForm.Close()
            
            # Notify user
            [System.Windows.Forms.MessageBox]::Show(
                "Update downloaded. The application will restart to complete the update.",
                "Update Ready",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
            
            # Start the update batch file
            Start-Process -FilePath "cmd.exe" -ArgumentList "/c $batchFile" -WindowStyle Hidden
            
            # Exit the current instance
            $form.Close()
        } else {
            # No direct script file, try the zip download
            $progressForm.Close()
            
            [System.Windows.Forms.MessageBox]::Show(
                "No direct script file found in the release. Please download the update manually from GitHub.\n\nURL: $($ReleaseInfo.html_url)",
                "Manual Update Required",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
            
            # Open the release page in the browser
            Start-Process $ReleaseInfo.html_url
        }
    } catch {
        if ($progressForm -and $progressForm.Visible) {
            $progressForm.Close()
        }
        
        $errorMessage = $_.Exception.Message
        Write-Log "Error updating application: $errorMessage" "Red"
        [System.Windows.Forms.MessageBox]::Show(
            "Update failed: $errorMessage",
            "Update Failed",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
}

# Function to load application settings
function Load-AppSettings {
    try {
        # Check if file exists
        if (-not (Test-Path -Path $script:AppSettingsFile)) {
            # No settings file exists yet, that's okay
            Write-Host "No settings file found at: $($script:AppSettingsFile)" -ForegroundColor Yellow
            return $true
        }
        
        Write-Host "Loading settings from: $($script:AppSettingsFile)" -ForegroundColor Cyan
        
        # Load XML document
        $xmlDoc = New-Object System.Xml.XmlDocument
        $xmlDoc.Load($script:AppSettingsFile)
        
        # Load recent files
        $recentFilesElement = $xmlDoc.SelectSingleNode("//RecentFiles")
        if ($recentFilesElement) {
            $script:RecentFiles = @()
            foreach ($fileElement in $recentFilesElement.SelectNodes("File")) {
                $filePath = $fileElement.InnerText
                # Only add files that still exist
                if (Test-Path -Path $filePath) {
                    Write-Host "  Found recent file: $filePath" -ForegroundColor Green
                    $script:RecentFiles += $filePath
                } else {
                    Write-Host "  Skipping missing file: $filePath" -ForegroundColor Yellow
                }
            }
            
            Write-Host "Loaded $($script:RecentFiles.Count) recent files" -ForegroundColor Cyan
        } else {
            Write-Host "No recent files found in settings" -ForegroundColor Yellow
        }
        
        return $true
    }
    catch {
        $errorMessage = $_.Exception.Message
        Write-Host "Error loading application settings: $errorMessage" -ForegroundColor Red
        # Reset to defaults
        $script:RecentFiles = @()
        return $false
    }
}

# Function to update the Recent Files menu
function Update-RecentFilesMenu {
    # Clear existing items
    $recentFilesMenuItem.DropDownItems.Clear()
    
    if ($script:RecentFiles.Count -eq 0) {
        # Add a disabled item if there are no recent files
        $noRecentFilesItem = New-Object System.Windows.Forms.ToolStripMenuItem
        $noRecentFilesItem.Text = "(No recent files)"
        $noRecentFilesItem.Enabled = $false
        $recentFilesMenuItem.DropDownItems.Add($noRecentFilesItem)
    } else {
        # Add each recent file to the menu
        foreach ($file in $script:RecentFiles) {
            $fileItem = New-Object System.Windows.Forms.ToolStripMenuItem
            $fileItem.Text = [System.IO.Path]::GetFileName($file)
            $fileItem.ToolTipText = $file
            # Store the full path in the Tag property
            $fileItem.Tag = $file
            $fileItem.Add_Click({
                # Use the Tag property to get the file path
                $clickedItem = $this
                $configPath = $clickedItem.Tag
                Write-Host "Loading configuration from recent file: $configPath" -ForegroundColor Cyan
                Load-Configuration -ConfigPath $configPath
            })
            $recentFilesMenuItem.DropDownItems.Add($fileItem)
        }
        
        # Add separator and Clear Recent Files option
        $recentFilesMenuItem.DropDownItems.Add("-")
        
        $clearRecentFilesItem = New-Object System.Windows.Forms.ToolStripMenuItem
        $clearRecentFilesItem.Text = "Clear Recent Files"
        $clearRecentFilesItem.Add_Click({
            $script:RecentFiles = @()
            Save-AppSettings
            Update-RecentFilesMenu
        })
        $recentFilesMenuItem.DropDownItems.Add($clearRecentFilesItem)
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
        
        # Load Label Folder and Template locations
        # Label Input Folder
        $labelInputElement = $xmlDoc.SelectSingleNode("//Labels/InputFolder")
        if ($labelInputElement) {
            $script:LabelInputFolder = $labelInputElement.InnerText
        }
        
        # Label Output Folder
        $labelOutputElement = $xmlDoc.SelectSingleNode("//Labels/OutputFolder")
        if ($labelOutputElement) {
            $script:LabelOutputFolder = $labelOutputElement.InnerText
        }
        
        # Param Template
        $paramTemplateElement = $xmlDoc.SelectSingleNode("//Labels/ParamTemplate")
        if ($paramTemplateElement) {
            $script:LabelParamTemplate = $paramTemplateElement.InnerText
        }
        
        # PRT Template
        $prtTemplateElement = $xmlDoc.SelectSingleNode("//Labels/PrtTemplate")
        if ($prtTemplateElement) {
            $script:LabelPrtTemplate = $prtTemplateElement.InnerText
        }
        
        # Dymo Template
        $dymoTemplateElement = $xmlDoc.SelectSingleNode("//Labels/DymoTemplate")
        if ($dymoTemplateElement) {
            $script:LabelDymoTemplate = $dymoTemplateElement.InnerText
        }
        
        # Load options states
        $optionsElement = $xmlDoc.SelectSingleNode("//Options")
        if ($optionsElement) {
            for ($i = 0; $i -lt $optionCheckboxes.Count; $i++) {
                $checkbox = $optionCheckboxes[$i]
                # Use checkbox Name if available and not empty, otherwise use index for lookup
                $elementName = if (-not [string]::IsNullOrWhiteSpace($checkbox.Name)) { $checkbox.Name } else { "Option$($i+1)" }
                $optionNode = $optionsElement.SelectSingleNode($elementName)
                if ($optionNode) {
                    $checkbox.Checked = ($optionNode.InnerText -eq 'True')
                } else {
                    # Fallback for older config files that might use indexed names like 'Option11'
                    $fallbackNode = $optionsElement.SelectSingleNode("Option$($i+1)")
                    if ($fallbackNode) {
                        $checkbox.Checked = ($fallbackNode.InnerText -eq 'True')
                    }
                }
            }
        }
        
        # Set current config file
        $script:CurrentConfigFile = $ConfigPath
        
        # Add to recent files list
        Add-RecentFile -FilePath $ConfigPath
        
        # Update the Recent Files menu
        Update-RecentFilesMenu
        
        Write-Log "Configuration loaded from: $ConfigPath" "Green"
        return $true
    }
    catch {
        Write-Log "Error loading configuration: $_" "Red"
        return $false
    }
}

# Initialize script variables for label paths
$script:LabelInputFolder = ""
$script:LabelOutputFolder = ""
$script:LabelParamTemplate = ""
$script:LabelPrtTemplate = ""
$script:LabelDymoTemplate = ""

# Initialize application settings and Recent Files menu
Write-Log "Loading application settings..." "Cyan"
Load-AppSettings
Update-RecentFilesMenu

# Display welcome and helpful information
Write-Log "=== Spreadsheet Wrangler v1.8.5 ===" "Cyan"
Write-Log "Application initialized and ready to use." "Green"

# Getting started section
Write-Log "GETTING STARTED:" "Yellow"
Write-Log "1. Add backup folders" "White" 
Write-Log "   Folders to create timestamped backups of" "Gray"
Write-Log "2. Add spreadsheet folders" "White" 
Write-Log "   Source folders containing spreadsheets to combine" "Gray"
Write-Log "3. Set combined destination folder" "White" 
Write-Log "   Where combined spreadsheets will be saved" "Gray"
Write-Log "4. Set SKU list location" "White" 
Write-Log "   Path to the SKU list CSV file for processing" "Gray"
Write-Log "5. Set final output location" "White" 
Write-Log "   Where final processed files will be saved" "Gray"
Write-Log "6. Select options from checkboxes" "White" 
Write-Log "   Customize how files are processed" "Gray"
Write-Log "7. Click 'Run' to start" "White"

# Options section
Write-Log "OPTIONS:" "Yellow"
Write-Log "- Skip Backup: Skip the backup process" "White"
Write-Log "- Skip Combine: Skip the spreadsheet combining process" "White"
Write-Log "- No Headers: Exclude headers when combining spreadsheets" "White"
Write-Log "- Duplicate by Qty: Duplicate rows based on quantity value" "White"
Write-Log "- Normalize Qty to 1: Change all quantity values to 1" "White"
Write-Log "- All Formats: Process multiple spreadsheet formats" "White"
Write-Log "- BLANK: Insert separator rows between spreadsheets" "White"
Write-Log "- Reverse, Reverse: Reverse the order of data rows" "White"
Write-Log "- Log to File: Save terminal output to a log file" "White"
Write-Log "- Single Spreadsheet: Duplicate the contents of a single spreadsheet" "White"

# Tips section
Write-Log "TIPS:" "Yellow"
Write-Log "- Save/load settings: File > Save/Open Configuration" "White"
Write-Log "- Quick access: File > Recent Files" "White"
Write-Log "- Remove items: Select and press Delete" "White"
Write-Log "- Create labels: Labels > Create Labels" "White"
Write-Log "- Button states: Blue (Ready), Yellow (Running), Green (Finished)" "White"

# Show the form
$form.ShowDialog()
