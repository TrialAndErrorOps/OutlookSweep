Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

# Create and show splash screen before anything else
function Show-SplashScreen {
    # Create splash screen window
    $script:splashWindow = [System.Windows.Window]::new()
    $script:splashWindow.Title = "Email Sweep Tool"
    $script:splashWindow.Width = 400
    $script:splashWindow.Height = 200
    $script:splashWindow.WindowStartupLocation = "CenterScreen"
    $script:splashWindow.WindowStyle = "None"
    $script:splashWindow.AllowsTransparency = $true
    $script:splashWindow.Background = "Transparent"
    $script:splashWindow.Topmost = $true
    $script:splashWindow.ShowInTaskbar = $false

    # Main border
    $border = [System.Windows.Controls.Border]::new()
    $border.CornerRadius = [System.Windows.CornerRadius]::new(8)
    $border.Background = "#FFFFFF"
    $border.Effect = [System.Windows.Media.Effects.DropShadowEffect]::new()
    $border.Effect.ShadowDepth = 3
    $border.Effect.BlurRadius = 15
    $border.Effect.Opacity = 0.3

    # Content stack panel
    $stackPanel = [System.Windows.Controls.StackPanel]::new()
    $stackPanel.Margin = "20"
    $border.Child = $stackPanel

    # App logo/icon
    $iconText = [System.Windows.Controls.TextBlock]::new()
    $iconText.Text = "📧"
    $iconText.FontSize = 36
    $iconText.HorizontalAlignment = "Center"
    $iconText.Margin = "0,0,0,10"
    $iconText.Foreground = "#2196F3"
    $stackPanel.Children.Add($iconText)

    # App name
    $titleText = [System.Windows.Controls.TextBlock]::new()
    $titleText.Text = "Email Sweep Tool"
    $titleText.FontSize = 20
    $titleText.FontWeight = "Bold"
    $titleText.HorizontalAlignment = "Center"
    $titleText.Margin = "0,0,0,20"
    $titleText.Foreground = "#212121"
    $stackPanel.Children.Add($titleText)

    # Loading text
    $script:loadingText = [System.Windows.Controls.TextBlock]::new()
    $script:loadingText.Text = "Initializing..."
    $script:loadingText.HorizontalAlignment = "Center"
    $script:loadingText.Margin = "0,0,0,10"
    $script:loadingText.Foreground = "#757575"
    $stackPanel.Children.Add($script:loadingText)

    # Progress bar
    $script:progressBar = [System.Windows.Controls.ProgressBar]::new()
    $script:progressBar.Height = 10
    $script:progressBar.Minimum = 0
    $script:progressBar.Maximum = 100
    $script:progressBar.Value = 0
    $script:progressBar.Foreground = "#2196F3"
    $stackPanel.Children.Add($script:progressBar)

    # Set content and show
    $script:splashWindow.Content = $border
    
    # Show the window without blocking
    $script:splashWindow.Show()
    
    # Force UI update
    [System.Windows.Forms.Application]::DoEvents()
}

# Function to update splash screen progress
function Update-SplashProgress {
    param (
        [int]$PercentComplete,
        [string]$StatusText
    )
    
    if ($script:progressBar -and $script:loadingText) {
        $script:progressBar.Value = $PercentComplete
        $script:loadingText.Text = $StatusText
        
        # Force UI update
        $script:splashWindow.Dispatcher.Invoke([Action]{}, [Windows.Threading.DispatcherPriority]::Render)
        [System.Windows.Forms.Application]::DoEvents()
    }
}

# Function to close splash screen
function Close-SplashScreen {
    if ($script:splashWindow) {
        $script:splashWindow.Close()
        $script:splashWindow = $null
        $script:progressBar = $null
        $script:loadingText = $null
    }
}

# Display splash screen immediately
Show-SplashScreen
Update-SplashProgress -PercentComplete 5 -StatusText "Starting initialization..."

Start-Sleep -Milliseconds 1000


# Improved Release-ComObject function with explicit cleanup
function Release-ComObject {
    param([object]$ComObject)
    
    if ($null -ne $ComObject) {
        try {
            # Check if it's a collection and handle specially
            if ($ComObject -is [System.__ComObject] -and 
                ($ComObject.GetType().FullName -like "*Collection*")) {
                for ($i = $ComObject.Count; $i -gt 0; $i--) {
                    try {
                        $item = $ComObject.Item($i)
                        if ($item) {
                            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($item) | Out-Null
                        }
                    } catch {}
                }
            }
            
            # Release the actual object
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ComObject) | Out-Null
        }
        catch {}
    }
}

Update-SplashProgress -PercentComplete 15 -StatusText "Loading helper functions..."

# Add this at the top with other helper functions
function Clean-OutlookExternalProcess {
    # Create a temporary script file
    $tempFile = [System.IO.Path]::GetTempFileName() + ".ps1"
    
    # Write cleanup code to the temp file
    @"
# Wait for parent process to exit
Start-Sleep -Seconds 2

# Get Outlook processes
try {
    `$processes = Get-Process -Name "outlook" -ErrorAction SilentlyContinue
    foreach (`$process in `$processes) {
        # Check if process was started recently (within last 5 minutes)
        if ((`$process.StartTime) -and ([datetime]::Now - `$process.StartTime).TotalMinutes -lt 5) {
            Write-Host "Terminating Outlook process ID: `$(`$process.Id)"
            Stop-Process -Id `$process.Id -Force
        }
    }
}
catch {
    # Log error but continue
    `$_ | Out-File -FilePath "`$env:TEMP\OutlookCleanupError.log" -Append
}

# Clean up the script itself
Remove-Item -Path '$tempFile' -Force
"@ | Out-File -FilePath $tempFile
    
    # Execute the cleanup script in a new, hidden PowerShell process
    Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$tempFile`"" -WindowStyle Hidden
}

# Function to save current sweep configuration - fixed version
function Save-SweepConfiguration {
    param(
        [string]$Name
    )
    
    if ([string]::IsNullOrWhiteSpace($Name)) {
        [System.Windows.MessageBox]::Show("Please enter a name for this sweep configuration.", "Name Required", 
            [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        return
    }
    
     # Create the sweep configuration object - add category information
     $sweepConfig = [PSCustomObject]@{
        Name = $Name
        SenderName = $senderNameTextBox.Text.Trim()
        SenderEmail = $senderEmailTextBox.Text.Trim()
        Subject = $subjectTextBox.Text.Trim()
        SearchFolder = $searchFolderComboBox.Text.Trim()
        TargetFolder = $folderComboBox.Text.Trim()
        Category = $categoryComboBox.Text.Trim()
        RemoveCategories = $removeCategoriesCheckBox.IsChecked
        MarkAsRead = $markAsReadCheckBox.IsChecked  # Add this line
        Created = Get-Date
    }
    
    # Check if we need to create the save directory
    $appDataPath = [Environment]::GetFolderPath('ApplicationData')
    $saveDir = Join-Path -Path $appDataPath -ChildPath "EmailSweepTool"
    
    if (-not (Test-Path -Path $saveDir)) {
        New-Item -Path $saveDir -ItemType Directory | Out-Null
    }
    
    # Save path for sweep configurations
    $savePath = Join-Path -Path $saveDir -ChildPath "SavedSweeps.xml"
    
    # Load existing configurations or create a new array
    $savedSweeps = @()
    if (Test-Path -Path $savePath) {
        try {
            # Load the data and ensure it's an array
            $importedData = Import-Clixml -Path $savePath
            
            # Convert to array if it's not already one
            if ($importedData -is [System.Array]) {
                $savedSweeps = $importedData
            } 
            elseif ($null -ne $importedData) {
                # If it's a single object, create a new array with that item
                $savedSweeps = @($importedData)
            }
        }
        catch {
            Write-Log "Warning: Failed to load existing sweep configurations. Starting fresh." -isError
            $savedSweeps = @()
        }
    }
    
    # Check for duplicates
    $existingSweep = $savedSweeps | Where-Object { $_.Name -eq $Name }
    if ($existingSweep) {
        $result = Show-CustomDialog -Title "Confirm Replace" `
            -Message "A sweep configuration named '$Name' already exists. Do you want to replace it?" `
            -YesNoButtons
        
        if ($result -eq $false) {
            return
        }
        
        # Remove the existing item
        $savedSweeps = @($savedSweeps | Where-Object { $_.Name -ne $Name })
    }
    
    # Add the new configuration - use array construction instead of +=
    $savedSweeps = @($savedSweeps) + @($sweepConfig)
    
    # Save to file
    try {
        $savedSweeps | Export-Clixml -Path $savePath -Force
        Write-Log "Sweep configuration '$Name' saved successfully."
        
        # Refresh the saved sweeps list
        Load-SavedSweeps
    }
    catch {
        Write-Log "Error saving sweep configuration: $($_.Exception.Message)" -isError
    }
}

# Function to load saved sweeps into the ListView - fixed version
function Load-SavedSweeps {
    # Clear existing items
    if ($savedSweepsListView) {
        $savedSweepsListView.Items.Clear()
    }
    
    # Load saved configurations
    $appDataPath = [Environment]::GetFolderPath('ApplicationData')
    $saveDir = Join-Path -Path $appDataPath -ChildPath "EmailSweepTool"
    $savePath = Join-Path -Path $saveDir -ChildPath "SavedSweeps.xml"
    
    $hasSavedSweeps = $false
    
    if (Test-Path -Path $savePath) {
        try {
            # Load the data and ensure it's an array
            $importedData = Import-Clixml -Path $savePath
            
            # Convert to array if it's not already one
            if ($importedData -is [System.Array]) {
                $savedSweeps = $importedData
            } 
            elseif ($null -ne $importedData) {
                # If it's a single object, create a new array with that item
                $savedSweeps = @($importedData)
            }
            else {
                $savedSweeps = @()
            }
                      
            # Add each sweep to the ListView
            foreach ($sweep in $savedSweeps) {
                $hasSavedSweeps = $true
                
                # Create the ListView item
                $item = New-Object System.Windows.Controls.ListViewItem
                
                # Create a grid for the item content with minimal margins
                $grid = New-Object System.Windows.Controls.Grid
                $grid.Margin = "2,2,2,2"  # Minimal margin
                
                # Define a single row only - everything will be in this row
                $row1 = New-Object System.Windows.Controls.RowDefinition
                $row1.Height = "Auto"
                $grid.RowDefinitions.Add($row1)
                
                # Create a single vertical stack panel for all content
                $contentPanel = New-Object System.Windows.Controls.StackPanel
                $contentPanel.Orientation = "Vertical"
                $contentPanel.Margin = "2,2,2,2"  # Ultra compact margins
                
                # LINE 1: Name and date on same line using DockPanel
                $headerPanel = New-Object System.Windows.Controls.DockPanel
                $headerPanel.LastChildFill = $false  # Important - don't let last child fill
                $headerPanel.Margin = New-Object System.Windows.Thickness(0, 0, 0, 2)
                
                # Name at left
                $nameText = New-Object System.Windows.Controls.TextBlock
                $nameText.Text = $sweep.Name
                $nameText.FontWeight = "Bold"
                $nameText.FontSize = 13
                $nameText.Foreground = "#2196F3"
                [System.Windows.Controls.DockPanel]::SetDock($nameText, "Left")
                $headerPanel.Children.Add($nameText)
                
                # Date at right
                $dateText = New-Object System.Windows.Controls.TextBlock
                $dateText.Text = "Created: " + $sweep.Created.ToString("MM/dd/yyyy")
                $dateText.FontSize = 11
                $dateText.Foreground = "#9E9E9E"
                [System.Windows.Controls.DockPanel]::SetDock($dateText, "Right")
                $headerPanel.Children.Add($dateText)
                
                $contentPanel.Children.Add($headerPanel)
                
                # LINE 2: All criteria on a single line WITH folder path included
                $allDetailsText = New-Object System.Windows.Controls.TextBlock
                $details = @()
                
                # Add criteria first
                if (-not [string]::IsNullOrWhiteSpace($sweep.SenderName)) { 
                    $details += "Name: $($sweep.SenderName)" 
                }
                if (-not [string]::IsNullOrWhiteSpace($sweep.SenderEmail)) { 
                    $details += "Email: $($sweep.SenderEmail)" 
                }
                if (-not [string]::IsNullOrWhiteSpace($sweep.Subject)) { 
                    $details += "Subject: $($sweep.Subject)" 
                }
                
                # Always add folder path to the criteria line
                $details += "Folders: $($sweep.SearchFolder) → $($sweep.TargetFolder)"
                
                # Set up the combined line
                $allDetailsText.Text = $details -join " | "
                $allDetailsText.TextWrapping = "NoWrap"
                $allDetailsText.TextTrimming = "CharacterEllipsis"
                $allDetailsText.FontSize = 11
                $allDetailsText.Foreground = "#606060"
                $contentPanel.Children.Add($allDetailsText)
                
                # Add the content panel to the grid
                $grid.Children.Add($contentPanel)
                
                # Set the grid as content
                $item.Content = $grid
                $item.Tag = $sweep
                
                # Minimal padding for the item
                $item.Padding = New-Object System.Windows.Thickness(5, 3, 5, 3)
                
                # Add to ListView
                $savedSweepsListView.Items.Add($item)
            }        }
        catch {
            Write-Log "Error loading saved sweeps: $($_.Exception.Message)" -isError
        }
    }
    
    # Show empty state message if no sweeps found
    if (-not $hasSavedSweeps) {
        $emptyStateItem = New-Object System.Windows.Controls.ListViewItem
        
        $emptyPanel = New-Object System.Windows.Controls.StackPanel
        $emptyPanel.Margin = "20"
        $emptyPanel.HorizontalAlignment = "Center"
        
        $emptyIcon = New-Object System.Windows.Controls.TextBlock
        $emptyIcon.Text = "📁"
        $emptyIcon.FontSize = 48
        $emptyIcon.Foreground = "#BDBDBD"
        $emptyIcon.HorizontalAlignment = "Center"
        $emptyPanel.Children.Add($emptyIcon)
        
        $emptyText = New-Object System.Windows.Controls.TextBlock
        $emptyText.Text = "No saved sweeps found"
        $emptyText.FontSize = 16
        $emptyText.Foreground = "#757575"
        $emptyText.Margin = "0,10,0,5"
        $emptyText.HorizontalAlignment = "Center"
        $emptyPanel.Children.Add($emptyText)
        
        $emptySubText = New-Object System.Windows.Controls.TextBlock
        $emptySubText.Text = "Create a sweep configuration and click 'Save Current Sweep'"
        $emptySubText.Foreground = "#9E9E9E"
        $emptySubText.HorizontalAlignment = "Center"
        $emptyPanel.Children.Add($emptySubText)
        
        $emptyStateItem.Content = $emptyPanel
        $emptyStateItem.IsEnabled = $false
        $emptyStateItem.Background = "Transparent"
        $emptyStateItem.BorderThickness = 0
        
        $savedSweepsListView.Items.Add($emptyStateItem)
        
        # Disable buttons that need a selection
        $loadSweepButton.IsEnabled = $false
        $deleteSweepButton.IsEnabled = $false
        $runSweepButton.IsEnabled = $false
        $runAllSweepsButton.IsEnabled = $false
    }
    else {
        # Enable buttons when we have sweeps
        $loadSweepButton.IsEnabled = $true
        $deleteSweepButton.IsEnabled = $true
        $runSweepButton.IsEnabled = $true
        $runAllSweepsButton.IsEnabled = $true
    }
}

# Function to load a saved sweep configuration into the UI
function Load-SweepConfiguration {
    param($sweep)
    
    if ($sweep) {
        # Update UI controls with the saved configuration
        # Don't set date values from saved configuration - use current values instead
        $senderNameTextBox.Text = $sweep.SenderName
        $senderEmailTextBox.Text = $sweep.SenderEmail
        $subjectTextBox.Text = $sweep.Subject
        $searchFolderComboBox.Text = $sweep.SearchFolder
        $folderComboBox.Text = $sweep.TargetFolder
        
        # Update category controls if the saved sweep has category information
        if ($sweep.PSObject.Properties.Name -contains "Category") {
            if ([string]::IsNullOrWhiteSpace($sweep.Category) -or $sweep.Category -eq "(None)") {
                $categoryComboBox.SelectedIndex = 0
            } else {
                $categoryComboBox.Text = $sweep.Category
            }
        } else {
            # Default to "None" for older saved sweeps
            $categoryComboBox.SelectedIndex = 0
        }
        
        # Set RemoveCategories checkbox if property exists
        if ($sweep.PSObject.Properties.Name -contains "RemoveCategories") {
            $removeCategoriesCheckBox.IsChecked = $sweep.RemoveCategories
        } else {
            $removeCategoriesCheckBox.IsChecked = $false
        }

        # Set MarkAsRead checkbox if property exists
        if ($sweep.PSObject.Properties.Name -contains "MarkAsRead") {
            $markAsReadCheckBox.IsChecked = $sweep.MarkAsRead
        } else {
            $markAsReadCheckBox.IsChecked = $false
        }
        
        # Switch to the Mail Sweep tab more reliably
        if ($null -ne $tabControl) {
            try {
                # Ensure UI update by using dispatcher with high priority
                $window.Dispatcher.Invoke([Action]{
                    $tabControl.SelectedIndex = 0
                }, [Windows.Threading.DispatcherPriority]::Send)
            }
            catch {
                Write-Log "Error switching tab: $($_.Exception.Message)" -isError
            }
        }
        
        Write-Log "Loaded sweep configuration: $($sweep.Name)"
        Write-Log "From folder: $($sweep.SearchFolder) → To folder: $($sweep.TargetFolder)"
        
        # Removed the date range log line
        
        if (-not [string]::IsNullOrWhiteSpace($sweep.SenderName)) {
            Write-Log "Sender name contains: $($sweep.SenderName)"
        }
        if (-not [string]::IsNullOrWhiteSpace($sweep.SenderEmail)) {
            Write-Log "Email address equals: $($sweep.SenderEmail)"
        }
        if (-not [string]::IsNullOrWhiteSpace($sweep.Subject)) {
            Write-Log "Subject contains: $($sweep.Subject)"
        }
        
        # Add log entry for category if applicable
        if (-not [string]::IsNullOrWhiteSpace($sweep.Category) -and $sweep.Category -ne "(None)") {
            Write-Log "Apply category: $($sweep.Category)"
        }
    }
}

# Function to delete a saved sweep configuration - fixed version
function Delete-SweepConfiguration {
    param($sweep)
    
    if ($sweep) {
        $result = Show-CustomDialog -Title "Confirm Delete" `
            -Message "Are you sure you want to delete the sweep configuration '$($sweep.Name)'?" `
            -YesNoButtons
        
        if ($result -eq $true) {
            # Load all configurations
            $appDataPath = [Environment]::GetFolderPath('ApplicationData')
            $saveDir = Join-Path -Path $appDataPath -ChildPath "EmailSweepTool"
            $savePath = Join-Path -Path $saveDir -ChildPath "SavedSweeps.xml"
            
            if (Test-Path -Path $savePath) {
                try {
                    # Load the data and ensure it's an array
                    $importedData = Import-Clixml -Path $savePath
                    
                    # Convert to array if it's not already one
                    if ($importedData -is [System.Array]) {
                        $savedSweeps = $importedData
                    } 
                    elseif ($null -ne $importedData) {
                        # If it's a single object, create a new array with that item
                        $savedSweeps = @($importedData)
                    }
                    else {
                        $savedSweeps = @()
                    }
                    
                    # Remove the selected sweep
                    $savedSweeps = @($savedSweeps | Where-Object { $_.Name -ne $sweep.Name })
                    
                    # Save the updated list
                    if ($savedSweeps.Count -gt 0) {
                        $savedSweeps | Export-Clixml -Path $savePath -Force
                    }
                    else {
                        # If no sweeps left, remove the file
                        Remove-Item -Path $savePath -Force -ErrorAction SilentlyContinue
                    }
                    
                    # Refresh the list
                    Load-SavedSweeps
                    
                    Write-Log "Deleted sweep configuration: $($sweep.Name)"
                }
                catch {
                    Write-Log "Error deleting sweep configuration: $($_.Exception.Message)" -isError
                }
            }
        }
    }
}

# Add these functions after the Delete-SweepConfiguration function
# Function to export sweep configurations to a file
function Export-SweepConfigurations {
    param(
        [array]$SweepsToExport,
        [bool]$ExportAll = $false
    )
    
    try {
        # Get all saved configurations
        $appDataPath = [Environment]::GetFolderPath('ApplicationData')
        $saveDir = Join-Path -Path $appDataPath -ChildPath "EmailSweepTool"
        $savePath = Join-Path -Path $saveDir -ChildPath "SavedSweeps.xml"
        
        $savedSweeps = @()
        if (Test-Path -Path $savePath) {
            # Load existing configurations
            $importedData = Import-Clixml -Path $savePath
            
            # Convert to array if it's not already one
            if ($importedData -is [System.Array]) {
                $savedSweeps = $importedData
            } 
            elseif ($null -ne $importedData) {
                $savedSweeps = @($importedData)
            }
        }
        
        # Determine which sweeps to export
        $sweepsForExport = @()
        if ($ExportAll) {
            $sweepsForExport = $savedSweeps
        } else {
            foreach ($sweep in $SweepsToExport) {
                $sweepsForExport += $sweep
            }
        }
        
        if ($sweepsForExport.Count -eq 0) {
            Write-Log "No sweep configurations to export." -isError
            return $false
        }
        
        # Create a save file dialog
        $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveFileDialog.Filter = "XML Files (*.xml)|*.xml|All Files (*.*)|*.*"
        $saveFileDialog.Title = "Export Sweep Configurations"
        $saveFileDialog.FileName = "EmailSweepConfigurations.xml"
        $saveFileDialog.DefaultExt = "xml"
        
        # Show the dialog and process if OK was clicked
        if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            # Export configurations to the selected file
            $sweepsForExport | Export-Clixml -Path $saveFileDialog.FileName -Force
            Write-Log "Exported $($sweepsForExport.Count) sweep configurations to $($saveFileDialog.FileName)"
            return $true
        }
        
        return $false
    }
    catch {
        Write-Log "Error exporting sweep configurations: $($_.Exception.Message)" -isError
        return $false
    }
}

# Function to import sweep configurations from a file
# Simplified Import function that bypasses complex dialog handling
function Import-SweepConfigurations {
    try {
        # Create an open file dialog
        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openFileDialog.Filter = "XML Files (*.xml)|*.xml|All Files (*.*)|*.*" 
        $openFileDialog.Title = "Import Sweep Configurations"
        $openFileDialog.Multiselect = $false
        
        # Show the dialog and process if OK was clicked
        if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            # Load the configurations from the selected file
            try {
                Write-Log "Importing from: $($openFileDialog.FileName)"
                $importedSweeps = Import-Clixml -Path $openFileDialog.FileName
                Write-Log "Data loaded from file successfully"
                
                # Convert to array if it's not already one
                if ($importedSweeps -isnot [System.Array]) {
                    $importedSweeps = @($importedSweeps)
                }
                
                # Validate the imported data
                if ($importedSweeps.Count -eq 0) {
                    [System.Windows.MessageBox]::Show("No sweep configurations found in the selected file.", 
                        "Import Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
                    return $false
                }
                
                Write-Log "Found $($importedSweeps.Count) sweeps in file"
                
                # Check if the imported data has the correct structure
                $validSweeps = @()
                $invalidCount = 0
                
                foreach ($sweep in $importedSweeps) {
                    # Check for minimal required properties
                    if (($sweep.PSObject.Properties.Name -contains "Name") -and 
                        ($sweep.PSObject.Properties.Name -contains "SearchFolder") -and
                        ($sweep.PSObject.Properties.Name -contains "TargetFolder")) {
                        $validSweeps += $sweep
                    } else {
                        $invalidCount++
                    }
                }
                
                if ($validSweeps.Count -eq 0) {
                    [System.Windows.MessageBox]::Show("No valid sweep configurations found in the selected file.", 
                        "Import Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
                    return $false
                }
                
                Write-Log "Valid sweeps to import: $($validSweeps.Count)"
                
                # Get existing configurations
                $appDataPath = [Environment]::GetFolderPath('ApplicationData')
                $saveDir = Join-Path -Path $appDataPath -ChildPath "EmailSweepTool"
                $savePath = Join-Path -Path $saveDir -ChildPath "SavedSweeps.xml"
                
                Write-Log "Save path will be: $savePath"
                
                $savedSweeps = @()
                if (Test-Path -Path $savePath) {
                    # Load existing configurations
                    $existingData = Import-Clixml -Path $savePath
                    
                    # Convert to array if it's not already one
                    if ($existingData -is [System.Array]) {
                        $savedSweeps = $existingData
                    } 
                    elseif ($null -ne $existingData) {
                        $savedSweeps = @($existingData)
                    }
                }
                
                # Check for duplicates
                $duplicateCount = 0
                $duplicateNames = @()
                foreach ($sweep in $validSweeps) {
                    foreach ($existingSweep in $savedSweeps) {
                        if ($sweep.Name -eq $existingSweep.Name) {
                            $duplicateCount++
                            $duplicateNames += $sweep.Name
                            break
                        }
                    }
                }
                
                # Determine how to handle duplicates (if any)
                $replaceMode = $false
                if ($duplicateCount -gt 0) {
                    $message = "$($validSweeps.Count) sweep configurations will be imported, but $duplicateCount have the same names as existing sweeps.`n`nDo you want to replace the existing configurations?"
                    $result = [System.Windows.MessageBox]::Show($message, "Handle Duplicate Sweeps", 
                                [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
                    $replaceMode = ($result -eq [System.Windows.MessageBoxResult]::Yes)
                } 
                else {
                    # Just confirm import with standard MessageBox
                    $message = "Ready to import $($validSweeps.Count) sweep configurations. Continue?"
                    $result = [System.Windows.MessageBox]::Show($message, "Confirm Import", 
                              [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
                    if ($result -ne [System.Windows.MessageBoxResult]::Yes) {
                        Write-Log "User cancelled the import operation"
                        return $false
                    }
                }
                
                Write-Log "Replace mode: $replaceMode"
                
                # Process the import
                $importedCount = 0
                
                # Create save directory if it doesn't exist
                if (-not (Test-Path -Path $saveDir)) {
                    New-Item -Path $saveDir -ItemType Directory -Force | Out-Null
                }
                
                # Build the new list of sweep configurations
                $newSweepList = @()
                
                # First add all existing sweeps we're keeping 
                foreach ($existingSweep in $savedSweeps) {
                    $keepSweep = $true
                    
                    # Check if this is a duplicate that should be replaced
                    if ($replaceMode) {
                        foreach ($newSweep in $validSweeps) {
                            if ($existingSweep.Name -eq $newSweep.Name) {
                                $keepSweep = $false
                                break
                            }
                        }
                    }
                    
                    if ($keepSweep) {
                        $newSweepList += $existingSweep
                    }
                }
                
                # Now add all the imported sweeps
                foreach ($importSweep in $validSweeps) {
                    $addSweep = $true
                    
                    # If not replacing duplicates, check if it's a duplicate
                    if (-not $replaceMode) {
                        foreach ($existingSweep in $savedSweeps) {
                            if ($importSweep.Name -eq $existingSweep.Name) {
                                $addSweep = $false
                                break
                            }
                        }
                    }
                    
                    if ($addSweep) {
                        $newSweepList += $importSweep
                        $importedCount++
                    }
                }
                
                # Save the updated configurations
                if ($importedCount -gt 0) {
                    # Save the new sweep list
                    $newSweepList | Export-Clixml -Path $savePath -Force
                    
                    Write-Log "Successfully imported $importedCount sweep configurations"
                    
                    # Refresh the saved sweeps list
                    Load-SavedSweeps
                    return $true
                } else {
                    Write-Log "No new sweep configurations were imported"
                    return $false
                }
            }
            catch {
                Write-Log "Error importing sweep configurations: $($_.Exception.Message)" -isError
                [System.Windows.MessageBox]::Show("Failed to import sweep configurations: $($_.Exception.Message)", 
                        "Import Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
                return $false
            }
        }
        
        return $false
    }
    catch {
        Write-Log "Error importing sweep configurations: $($_.Exception.Message)" -isError
        return $false
    }
}

# Add this function to create custom styled dialog boxes that match the app's aesthetics
# Function to show a styled dialog using built-in MessageBox
# Function to show a styled dialog that matches the app's aesthetics
# Set up our script-scoped variable to hold dialog results
$script:DialogUserChoice = $false

# Function to show a styled dialog that stores result in a script variable
# Update the Show-CustomDialog function to properly handle dialog results
# Set up our script-scoped variable to hold dialog results
$script:DialogUserChoice = $false

# Modified Show-CustomDialog function that works like V5
function Show-CustomDialog {
    param (
        [string]$Title = "Confirmation",
        [string]$Message,
        [string]$OkButtonText = "OK",
        [string]$CancelButtonText = "Cancel",
        [System.Windows.Window]$Owner = $window,
        [switch]$YesNoButtons,
        [switch]$InfoOnly
    )
    
    # Reset the result value at the beginning
    $script:DialogUserChoice = $false
    
    # Create the dialog window
    $dialog = [System.Windows.Window]::new()
    $dialog.Title = $Title
    $dialog.Width = 460
    $dialog.Height = 350
    $dialog.WindowStartupLocation = "CenterOwner"
    $dialog.Owner = $Owner
    $dialog.ResizeMode = "NoResize"
    $dialog.Background = "#F5F5F5"
    
    # Create main border with shadow effect - match the app's card style
    $mainBorder = [System.Windows.Controls.Border]::new()
    $mainBorder.Margin = "10"
    $mainBorder.Background = "#FFFFFF"
    $mainBorder.CornerRadius = [System.Windows.CornerRadius]::new(8)
    $mainBorder.Effect = [System.Windows.Media.Effects.DropShadowEffect]::new()
    $mainBorder.Effect.ShadowDepth = 1
    $mainBorder.Effect.BlurRadius = 10
    $mainBorder.Effect.Opacity = 0.2
    $mainBorder.Effect.Color = [System.Windows.Media.Colors]::Black
    
    # Create content grid
    $contentGrid = [System.Windows.Controls.Grid]::new()
    $contentGrid.Margin = "20,20,20,20"
    
    # Define rows: content and buttons
    $rowContent = [System.Windows.Controls.RowDefinition]::new()
    $rowContent.Height = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
    $rowButtons = [System.Windows.Controls.RowDefinition]::new()
    $rowButtons.Height = [System.Windows.GridLength]::Auto
    
    $contentGrid.RowDefinitions.Add($rowContent)
    $contentGrid.RowDefinitions.Add($rowButtons)
    
    # Add message text with app style
    $messageText = [System.Windows.Controls.TextBlock]::new()
    $messageText.Text = $Message
    $messageText.TextWrapping = "Wrap"
    $messageText.Margin = "0,0,0,20"
    $messageText.VerticalAlignment = "Center"
    $messageText.HorizontalAlignment = "Left"
    $messageText.FontSize = 14
    $messageText.FontFamily = "Segoe UI"
    $messageText.Foreground = "#212121"
    [System.Windows.Controls.Grid]::SetRow($messageText, 0)
    $contentGrid.Children.Add($messageText)
    
    # Create button panel
    $buttonPanel = [System.Windows.Controls.StackPanel]::new()
    $buttonPanel.Orientation = "Horizontal"
    $buttonPanel.HorizontalAlignment = "Right"
    $buttonPanel.Margin = "0,10,0,0"
    [System.Windows.Controls.Grid]::SetRow($buttonPanel, 1)
    
    if ($InfoOnly) {
        # Just an OK button
        $okButton = [System.Windows.Controls.Button]::new()
        $okButton.Content = $OkButtonText
        $okButton.Width = 100
        $okButton.Height = 38
        $okButton.Margin = "5,0,0,0"
        $okButton.Background = "#2196F3"
        $okButton.Foreground = "White"
        $okButton.BorderThickness = 0
        $okButton.FontWeight = "SemiBold"
        $okButton.FontFamily = "Segoe UI"
        $okButton.Cursor = "Hand"
        $okButton.Add_Click({
            # Set both dialog result and script variable
            $script:DialogUserChoice = $true
            $dialog.DialogResult = $true
            $dialog.Close()
        })
        
        # Set button style with rounded corners
        $okButton.Template = [System.Windows.Markup.XamlReader]::Parse(@"
<ControlTemplate xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
                 TargetType="Button">
    <Border Background="{TemplateBinding Background}" 
            BorderBrush="{TemplateBinding BorderBrush}" 
            BorderThickness="{TemplateBinding BorderThickness}"
            CornerRadius="4">
        <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
    </Border>
</ControlTemplate>
"@)
        
        $buttonPanel.Children.Add($okButton)
        
        # Add hover effect for button
        $okButton.Add_MouseEnter({
            $this.Background = "#1976D2"
        })
        $okButton.Add_MouseLeave({
            $this.Background = "#2196F3"
        })
    }
    else {
        # Create primary action button
        $primaryButton = [System.Windows.Controls.Button]::new()
        if ($YesNoButtons) {
            $primaryButton.Content = "Yes"
        }
        else {
            $primaryButton.Content = $OkButtonText
        }
        $primaryButton.Width = 100
        $primaryButton.Height = 38
        $primaryButton.Margin = "5,0,0,0"
        $primaryButton.Background = "#2196F3"
        $primaryButton.Foreground = "White"
        $primaryButton.BorderThickness = 0
        $primaryButton.FontWeight = "SemiBold"
        $primaryButton.FontFamily = "Segoe UI"
        $primaryButton.Cursor = "Hand"
        $primaryButton.Add_Click({
            # IMPORTANT FIX: Set both dialog result and script variable
            $script:DialogUserChoice = $true
            $dialog.DialogResult = $true
            $dialog.Close()
        })
        
        # Set button style with rounded corners
        $primaryButton.Template = [System.Windows.Markup.XamlReader]::Parse(@"
<ControlTemplate xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
                 TargetType="Button">
    <Border Background="{TemplateBinding Background}" 
            BorderBrush="{TemplateBinding BorderBrush}" 
            BorderThickness="{TemplateBinding BorderThickness}"
            CornerRadius="4">
        <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
    </Border>
</ControlTemplate>
"@)
        
        # Create cancel button
        $cancelButton = [System.Windows.Controls.Button]::new()
        if ($YesNoButtons) {
            $cancelButton.Content = "No"
        }
        else {
            $cancelButton.Content = $CancelButtonText
        }
        $cancelButton.Width = 100
        $cancelButton.Height = 38
        $cancelButton.Margin = "10,0,0,0"
        $cancelButton.Background = "#E0E0E0"
        $cancelButton.Foreground = "#212121"
        $cancelButton.BorderThickness = 0
        $cancelButton.FontWeight = "Medium"
        $cancelButton.FontFamily = "Segoe UI"
        $cancelButton.Cursor = "Hand"
        $cancelButton.Add_Click({
            # IMPORTANT FIX: Set both dialog result and script variable explicitly to false
            $script:DialogUserChoice = $false
            $dialog.DialogResult = $false
            $dialog.Close()
        })
        
        # Set button style with rounded corners
        $cancelButton.Template = [System.Windows.Markup.XamlReader]::Parse(@"
<ControlTemplate xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
                 TargetType="Button">
    <Border Background="{TemplateBinding Background}" 
            BorderBrush="{TemplateBinding BorderBrush}" 
            BorderThickness="{TemplateBinding BorderThickness}"
            CornerRadius="4">
        <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
    </Border>
</ControlTemplate>
"@)
        
        # Add buttons to panel
        $buttonPanel.Children.Add($primaryButton)
        $buttonPanel.Children.Add($cancelButton)
        
        # Add hover effect for buttons
        $primaryButton.Add_MouseEnter({
            $this.Background = "#1976D2"
        })
        $primaryButton.Add_MouseLeave({
            $this.Background = "#2196F3"
        })
        
        $cancelButton.Add_MouseEnter({
            $this.Background = "#BDBDBD"
        })
        $cancelButton.Add_MouseLeave({
            $this.Background = "#E0E0E0"
        })
    }
    
    # Add button panel to grid
    $contentGrid.Children.Add($buttonPanel)
    
    # Set content
    $mainBorder.Child = $contentGrid
    $dialog.Content = $mainBorder
    
    # Make the primary button the default button
    $dialog.Add_Loaded({
        if ($InfoOnly) {
            $okButton.Focus()
        }
        else {
            $primaryButton.Focus()
        }
    })
    
    # Add escape key handling
    $dialog.Add_KeyDown({
        param($sender, $e)
        if ($e.Key -eq 'Escape' -and -not $InfoOnly) {
            # IMPORTANT FIX: Set both dialog result and script variable to false for Escape key
            $script:DialogUserChoice = $false
            $dialog.DialogResult = $false
            $dialog.Close()
        }
    })
    
    # Show dialog and wait for it to close
    $dialog.ShowDialog() | Out-Null
    
    # Return the script variable directly
    return $script:DialogUserChoice
}

Update-SplashProgress -PercentComplete 25 -StatusText "Preparing interface..."

# Create the XAML for our WPF GUI
[xml]$xaml = @"
<Window 
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Email Sweep Tool" Height="1000" Width="1200" 
    WindowStartupLocation="CenterScreen" ResizeMode="CanMinimize"
    Background="#F5F5F5">
    <Window.Resources>
        <!-- Modern color palette -->
        <SolidColorBrush x:Key="PrimaryColor" Color="#2196F3"/>
        <SolidColorBrush x:Key="AccentColor" Color="#FF4081"/>
        <SolidColorBrush x:Key="BackgroundColor" Color="#FFFFFF"/>
        <SolidColorBrush x:Key="TextColor" Color="#212121"/>
        <SolidColorBrush x:Key="SubtleTextColor" Color="#757575"/>
        <SolidColorBrush x:Key="BorderColor" Color="#E0E0E0"/>
        
        <!-- Text block style -->
        <Style TargetType="TextBlock">
            <Setter Property="Margin" Value="5"/>
            <Setter Property="VerticalAlignment" Value="Center"/>
            <Setter Property="Foreground" Value="{StaticResource TextColor}"/>
            <Setter Property="FontFamily" Value="Segoe UI"/>
            <Setter Property="FontSize" Value="14"/>
        </Style>
        
        <!-- Text box style -->
        <Style TargetType="TextBox">
            <Setter Property="Margin" Value="5"/>
            <Setter Property="Padding" Value="10,8"/>
            <Setter Property="Height" Value="38"/>
            <Setter Property="BorderBrush" Value="{StaticResource BorderColor}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="FontFamily" Value="Segoe UI"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
            <Style.Triggers>
                <Trigger Property="IsFocused" Value="True">
                    <Setter Property="BorderBrush" Value="{StaticResource PrimaryColor}"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        
        <!-- Date picker style -->
        <Style TargetType="DatePicker">
            <Setter Property="Margin" Value="5"/>
            <Setter Property="Height" Value="38"/>
            <Setter Property="BorderBrush" Value="{StaticResource BorderColor}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="FontFamily" Value="Segoe UI"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
            <Style.Triggers>
                <Trigger Property="IsFocused" Value="True">
                    <Setter Property="BorderBrush" Value="{StaticResource PrimaryColor}"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        
        <!-- Combo box style -->
        <Style TargetType="ComboBox">
            <Setter Property="Margin" Value="5"/>
            <Setter Property="Height" Value="38"/>
            <Setter Property="BorderBrush" Value="{StaticResource BorderColor}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="FontFamily" Value="Segoe UI"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
            <Setter Property="Padding" Value="8,4"/>
            <Style.Triggers>
                <Trigger Property="IsFocused" Value="True">
                    <Setter Property="BorderBrush" Value="{StaticResource PrimaryColor}"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        
        <!-- Button style -->
        <Style TargetType="Button">
            <Setter Property="Margin" Value="5"/>
            <Setter Property="Padding" Value="20,8"/>
            <Setter Property="Height" Value="38"/>
            <Setter Property="Background" Value="{StaticResource PrimaryColor}"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="FontFamily" Value="Segoe UI"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" 
                                BorderBrush="{TemplateBinding BorderBrush}" 
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="4">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#1976D2"/>
                </Trigger>
                <Trigger Property="IsPressed" Value="True">
                    <Setter Property="Background" Value="#0D47A1"/>
                </Trigger>
                <Trigger Property="IsEnabled" Value="False">
                    <Setter Property="Background" Value="#BDBDBD"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        
        <!-- Tab Item Style -->
        <Style TargetType="TabItem">
            <Setter Property="Padding" Value="15,8"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="FontWeight" Value="Medium"/>
            <Setter Property="BorderBrush" Value="{StaticResource BorderColor}"/>
            <Setter Property="Background" Value="#F2F2F2"/>
            <Style.Triggers>
                <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="#FFFFFF"/>
                    <Setter Property="BorderBrush" Value="{StaticResource PrimaryColor}"/>
                    <Setter Property="Foreground" Value="{StaticResource PrimaryColor}"/>
                </Trigger>
            </Style.Triggers>
        </Style>
    </Window.Resources>
    
    <TabControl Margin="10" BorderThickness="0">
        <!-- Mail Sweep Tab -->
        <TabItem Header="Mail Sweep">
            <Border Padding="20" Background="{StaticResource BackgroundColor}" CornerRadius="8" Margin="0,10,0,0">
                <Border.Effect>
                    <DropShadowEffect ShadowDepth="1" BlurRadius="10" Opacity="0.2" Color="#000000"/>
                </Border.Effect>
                
                    <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/> <!-- Category selection row -->
                        <RowDefinition Height="Auto"/> <!-- New Sweep Options row -->
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>

                    <!-- Title, Description and Info Button -->
                    <Grid Grid.Row="0" Grid.ColumnSpan="2" Margin="0,0,0,15">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>
                        
                        <!-- Title and Info Button on the same row -->
                        <TextBlock Grid.Column="0" Grid.Row="0" Text="Email Sweep Tool" FontWeight="Bold" 
                                  FontSize="24" Foreground="{StaticResource PrimaryColor}" 
                                  VerticalAlignment="Center"/>
                        <Button Grid.Column="1" Grid.Row="0" Content="ⓘ" FontSize="18" Width="30" Height="30" 
                                VerticalAlignment="Center" Margin="10,0,0,0" Padding="0"
                                Background="Transparent" BorderThickness="0" Foreground="{StaticResource PrimaryColor}"
                                x:Name="InfoButton"/>
                        
                        <!-- Description on the second row -->
                        <TextBlock Grid.Column="0" Grid.Row="1" Grid.ColumnSpan="3" 
                                  Text="Search, organize, and move emails in your Outlook inbox" 
                                  Foreground="{StaticResource SubtleTextColor}" FontSize="14" Margin="0,5,0,0"/>
                    </Grid>

                    <!-- Date Range Selections -->
                    <TextBlock Grid.Row="1" Grid.Column="0" Text="Date Range:" FontWeight="Medium"/>
                    <Grid Grid.Row="1" Grid.Column="1">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>
                        <DatePicker Grid.Column="0" x:Name="StartDatePicker" 
                               Width="200" HorizontalAlignment="Left"/>
                        <TextBlock Grid.Column="1" Text="to" Margin="10,0" VerticalAlignment="Center"/>
                        <DatePicker Grid.Column="2" x:Name="EndDatePicker" 
                               Width="200" HorizontalAlignment="Left"/>
                        <Button Grid.Column="3" x:Name="LastWeekButton" Content="Set to Last Week" 
                               Width="140" Margin="15,5,0,5" HorizontalAlignment="Left"/>
                    </Grid>
                    
                    <!-- Sender Name and Email - side by side -->
                    <TextBlock Grid.Row="2" Grid.Column="0" Text="Sender Name:" FontWeight="Medium"/>
                    <Grid Grid.Row="2" Grid.Column="1">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>
                        <TextBox Grid.Column="0" x:Name="SenderNameTextBox" 
                                Width="250" HorizontalAlignment="Left">
                            <TextBox.ToolTip>
                                <ToolTip Background="#F5F5F5" BorderBrush="#BDBDBD" BorderThickness="1" Padding="8,5">
                                    <TextBlock Text="This will search for a single name or full name." 
                                              TextWrapping="Wrap" MaxWidth="250" Foreground="#212121" />
                                </ToolTip>
                            </TextBox.ToolTip>
                        </TextBox>
                        <TextBlock Grid.Column="1" Text="Email Address:" Margin="15,0,5,0" FontWeight="Medium"/>
                        <TextBox Grid.Column="2" x:Name="SenderEmailTextBox" 
                                Width="250" HorizontalAlignment="Left">
                            <TextBox.ToolTip>
                                <ToolTip Background="#F5F5F5" BorderBrush="#BDBDBD" BorderThickness="1" Padding="8,5">
                                    <TextBlock Text="Enter entire email. This only searches for exact email address matches." 
                                              TextWrapping="Wrap" MaxWidth="250" Foreground="#212121" />
                                </ToolTip>
                            </TextBox.ToolTip>
                        </TextBox>
                    </Grid>
                                
                    <!-- Subject with tooltip -->
                    <TextBlock Grid.Row="3" Grid.Column="0" Text="Subject Contains:" FontWeight="Medium"/>
                    <TextBox Grid.Row="3" Grid.Column="1" x:Name="SubjectTextBox" 
                            Width="250" HorizontalAlignment="Left">
                        <TextBox.ToolTip>
                            <ToolTip Background="#F5F5F5" BorderBrush="#BDBDBD" BorderThickness="1" Padding="8,5">
                                <TextBlock Text="This only searches in the subject field. It does not search the contents of emails." 
                                           TextWrapping="Wrap" MaxWidth="250" Foreground="#212121" />
                            </ToolTip>
                        </TextBox.ToolTip>
                    </TextBox>

                    <!-- Folder Selection -->
                    <TextBlock Grid.Row="4" Grid.Column="0" Text="Search Folder:" FontWeight="Medium"/>
                    <Grid Grid.Row="4" Grid.Column="1">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>
                        <ComboBox Grid.Column="1" x:Name="SearchFolderComboBox" IsEditable="True" 
                                  Width="200" HorizontalAlignment="Left"/>
                        <TextBlock Grid.Column="2" Text="Destination Folder:" Margin="15,0,5,0" VerticalAlignment="Center" FontWeight="Medium"/>
                        <ComboBox Grid.Column="3" x:Name="FolderComboBox" IsEditable="True" 
                                  Width="200" HorizontalAlignment="Left">
                            <ComboBox.ToolTip>
                                <ToolTip Background="#F5F5F5" BorderBrush="#BDBDBD" BorderThickness="1" Padding="8,5">
                                    <TextBlock Text="Manually enter the name of a folder to create a new folder as part of the execution process." 
                                              TextWrapping="Wrap" MaxWidth="250" Foreground="#212121" />
                                </ToolTip>
                            </ComboBox.ToolTip>
                        </ComboBox>
                        <Button Grid.Column="4" x:Name="RefreshFoldersButton" 
                                Content="⟳" Width="38" Height="38" Padding="5" 
                                HorizontalAlignment="Left" Margin="5,0,0,0"
                                ToolTip="Refresh Folders List" 
                                BorderThickness="0"/>
                    </Grid>
                    
                        <!-- Category Selection -->
                        <TextBlock Grid.Row="5" Grid.Column="0" Text="Apply Category:" FontWeight="Medium"/>
                        <Grid Grid.Row="5" Grid.Column="1">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="Auto"/>
                                <ColumnDefinition Width="*"/>
                            </Grid.ColumnDefinitions>
                            <ComboBox Grid.Column="0" x:Name="CategoryComboBox" 
                                    Width="200" HorizontalAlignment="Left" IsReadOnly="True">
                                <ComboBox.ToolTip>
                                    <ToolTip Background="#F5F5F5" BorderBrush="#BDBDBD" BorderThickness="1" Padding="8,5">
                                        <TextBlock Text="Select a category to apply to emails when moving them.
                                                    Select '(None)' to not change categories." 
                                                TextWrapping="Wrap" MaxWidth="250" Foreground="#212121" />
                                    </ToolTip>
                                </ComboBox.ToolTip>
                            </ComboBox>
                        </Grid>
                        
                        <!-- New Sweep Options Section -->
                            <TextBlock Grid.Row="6" Grid.Column="0" Text="Sweep Options:" FontWeight="Medium"/>
                            <Grid Grid.Row="6" Grid.Column="1">
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto" MinHeight="60"/> <!-- Taller row to fit wrapped text -->
                                </Grid.RowDefinitions>
                                <WrapPanel Orientation="Horizontal" Margin="0,5">
                                    <CheckBox x:Name="RemoveCategoriesCheckBox" 
                                            Margin="0,5,20,5"
                                            VerticalAlignment="Center"
                                            VerticalContentAlignment="Center">
                                        <TextBlock Text="Remove existing categories" 
                                                TextWrapping="Wrap" 
                                                Width="150"
                                                VerticalAlignment="Center"/> <!-- Fixed width for consistent wrapping -->
                                        <CheckBox.ToolTip>
                                            <ToolTip Background="#F5F5F5" BorderBrush="#BDBDBD" BorderThickness="1" Padding="8,5">
                                                <TextBlock Text="Check this to remove all existing categories from emails during sweep." 
                                                        TextWrapping="Wrap" MaxWidth="250" Foreground="#212121" />
                                            </ToolTip>
                                        </CheckBox.ToolTip>
                                    </CheckBox>
                                    
                                    <!-- Template for additional checkboxes -->
                                    <CheckBox x:Name="MarkAsReadCheckBox" 
                                            Margin="0,5,20,5"
                                            VerticalAlignment="Center"
                                            VerticalContentAlignment="Center">
                                        <TextBlock Text="Mark moved emails as read" 
                                                TextWrapping="Wrap" 
                                                Width="150"
                                                VerticalAlignment="Center"/>
                                        <CheckBox.ToolTip>
                                            <ToolTip Background="#F5F5F5" BorderBrush="#BDBDBD" BorderThickness="1" Padding="8,5">
                                                <TextBlock Text="Check this to mark all emails as read during sweep" 
                                                        TextWrapping="Wrap" MaxWidth="250" Foreground="#212121" />
                                            </ToolTip>
                                        </CheckBox.ToolTip>
                                    </CheckBox>
                                                                    </WrapPanel>
                            </Grid>

                        <!-- Log output - update the row number to 7 (it was 6 before) -->
                        <Border Grid.Row="7" Grid.Column="0" Grid.ColumnSpan="2" 
                            BorderBrush="{StaticResource BorderColor}" BorderThickness="1" 
                            CornerRadius="4" Margin="5">
                        <RichTextBox x:Name="LogOutputTextBox" IsReadOnly="True" 
                            BorderThickness="0"
                            FontFamily="Consolas" Margin="0" Padding="10"
                            Background="#FAFAFA" Height="Auto"
                            VerticalAlignment="Stretch"
                            VerticalScrollBarVisibility="Auto"
                            ScrollViewer.CanContentScroll="True"
                            ScrollViewer.VerticalScrollBarVisibility="Auto">
                            <RichTextBox.Document>
                                <FlowDocument PageWidth="1000" FontFamily="Consolas"/>
                            </RichTextBox.Document>
                        </RichTextBox>
                    </Border>

                    <!-- Buttons -->
                    <Grid Grid.Row="8" Grid.Column="0" Grid.ColumnSpan="2" Margin="0,15,0,0">

                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>
                        
                        <!-- Left-aligned Save button -->
                        <Button Grid.Column="0" x:Name="SaveSweepButton" Content="Save Current Settings" 
                                Width="170" Margin="5" Background="#4CAF50" Foreground="White"
                                HorizontalAlignment="Left"/>
                        
                        <!-- Right-aligned action buttons -->
                        <StackPanel Grid.Column="1" Orientation="Horizontal" HorizontalAlignment="Right">
                            <Button x:Name="TestButton" Content="Test" Width="120" Margin="5"
                                    Background="#4CAF50" Foreground="White"/>
                            <Button x:Name="ExecuteButton" Content="Execute" Width="120" Margin="5"/>
                            <Button x:Name="CancelButton" Content="Exit" Width="120" Margin="5" 
                                    Background="#E0E0E0" Foreground="#212121"/>
                        </StackPanel>
                    </Grid>
                </Grid>
            </Border>
        </TabItem>
        
        <!-- Saved Sweeps Tab -->
        <TabItem Header="Saved Sweeps">
            <Border Padding="20" Background="{StaticResource BackgroundColor}" CornerRadius="8" Margin="0,10,0,0">
                <Border.Effect>
                    <DropShadowEffect ShadowDepth="1" BlurRadius="10" Opacity="0.2" Color="#000000"/>
                </Border.Effect>
                
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>  <!-- Added a row for the new text -->
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    
                                        <!-- Title, description, and Run All button -->
                    <Grid Grid.Row="0" Margin="0,0,0,10">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        
                        <!-- Title -->
                        <TextBlock Grid.Column="0" Text="Saved Sweep Configurations" 
                                  FontWeight="Bold" FontSize="20" Foreground="{StaticResource PrimaryColor}" 
                                  VerticalAlignment="Center"/>
                                  
                        <!-- Run All button in upper right -->
                        <Button Grid.Column="1" x:Name="RunAllSweepsButton" Content="Run All Saved Sweeps" 
                               Width="170" Height="32" Background="#4CAF50" Foreground="White"
                               HorizontalAlignment="Right"/>
                    </Grid>
                    
                    <TextBlock Grid.Row="1" Text="Save commonly used searches to quickly access them later" 
                              Foreground="{StaticResource SubtleTextColor}" Margin="0,0,0,5"/>
                              
                    <!-- Description text about date range -->
                    <TextBlock Grid.Row="2" TextWrapping="Wrap"
                              Text="Running saved sweeps directly from this tab using the Run Selected button will automatically sweep emails 7-35 days old.&#10;Press the 'Load Selected' button and adjust the date range on the Mail Sweep tab if you would like to sweep a different date range." 
                              Foreground="{StaticResource SubtleTextColor}" Margin="0,0,0,15"/>
                    <!-- Saved sweeps list -->
                    <Border Grid.Row="3" BorderBrush="{StaticResource BorderColor}" BorderThickness="1" 
                           CornerRadius="4" Background="#FAFAFA" Padding="0">
                        <ListView x:Name="SavedSweepsListView" BorderThickness="0" Background="Transparent"
                            ScrollViewer.HorizontalScrollBarVisibility="Disabled"
                            ScrollViewer.VerticalScrollBarVisibility="Auto"
                            ScrollViewer.CanContentScroll="True"
                            SelectionMode="Extended">
                            <ListView.ItemContainerStyle>
                                <Style TargetType="ListViewItem">
                                    <Setter Property="HorizontalContentAlignment" Value="Stretch"/>
                                    <Setter Property="Padding" Value="5,3"/> <!-- Ultra-minimal padding -->
                                    <Setter Property="Margin" Value="0,1"/> <!-- Minimal margin -->
                                    <Setter Property="Background" Value="White"/>
                                    <Setter Property="BorderBrush" Value="{StaticResource BorderColor}"/>
                                    <Setter Property="BorderThickness" Value="0,0,0,1"/>
                                    <Style.Triggers>
                                        <Trigger Property="IsSelected" Value="True">
                                            <Setter Property="Background" Value="#E3F2FD"/>
                                            <Setter Property="BorderBrush" Value="{StaticResource PrimaryColor}"/>
                                            <Setter Property="BorderThickness" Value="1"/>
                                        </Trigger>
                                    </Style.Triggers>
                                </Style>
                            </ListView.ItemContainerStyle>
                        </ListView>
                    </Border>
                    
                    <!-- Buttons -->
                    <Grid Grid.Row="4" Margin="0,15,0,0">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    
                    <!-- Import/Export buttons - left aligned -->
                    <Button Grid.Column="0" x:Name="ImportSweepsButton" Content="Import" 
                        Width="80" Margin="5,5,3,5"/>
                    <Button Grid.Column="1" x:Name="ExportSweepsButton" Content="Export" 
                        Width="80" Margin="3,5,5,5"/>
                    
                    <!-- Other buttons - right aligned -->
                    <Button Grid.Column="3" x:Name="LoadSweepButton" Content="Load Selected" 
                        Width="120" Margin="5,5,5,5"/>
                    <Button Grid.Column="4" x:Name="DeleteSweepButton" Content="Delete Selected" 
                        Width="120" Margin="5,5,5,5"/>
                    <Button Grid.Column="5" x:Name="RunSweepButton" Content="Run Selection" 
                        Width="120" Margin="5,5,5,5"/>
                </Grid>
                </Grid>
            </Border>
        </TabItem>
    </TabControl>
</Window>
"@

Update-SplashProgress -PercentComplete 40 -StatusText "Building user interface..."

# Create a form object from the XAML
$reader = [System.Xml.XmlNodeReader]::new($xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

# Form controls
$startDatePicker = $window.FindName('StartDatePicker')
$endDatePicker = $window.FindName('EndDatePicker')
$lastWeekButton = $window.FindName('LastWeekButton')
$subjectTextBox = $window.FindName('SubjectTextBox')
$senderEmailTextBox = $window.FindName('SenderEmailTextBox')
$senderNameTextBox = $window.FindName('SenderNameTextBox')
$searchFolderComboBox = $window.FindName('SearchFolderComboBox')
$folderComboBox = $window.FindName('FolderComboBox')
$logOutputTextBox = $window.FindName('LogOutputTextBox')
$executeButton = $window.FindName('ExecuteButton')
$cancelButton = $window.FindName('CancelButton')
$refreshFoldersButton = $window.FindName('RefreshFoldersButton')
$testButton = $window.FindName('TestButton')

# Add these after your form controls initialization
# Instead of looking for a TabControl by name, get it directly from the window's content
$tabControl = $window.Content
$savedSweepsListView = $window.FindName('SavedSweepsListView')
$saveSweepButton = $window.FindName('SaveSweepButton')
$loadSweepButton = $window.FindName('LoadSweepButton')
$deleteSweepButton = $window.FindName('DeleteSweepButton')
$runSweepButton = $window.FindName('RunSweepButton')

# Get reference to the new Run All button
$runAllSweepsButton = $window.FindName('RunAllSweepsButton')

# Get references to the new category controls
$categoryComboBox = $window.FindName('CategoryComboBox')
$removeCategoriesCheckBox = $window.FindName('RemoveCategoriesCheckBox')
$markAsReadCheckBox = $window.FindName('MarkAsReadCheckBox')

# Add these button references after the existing button references
$importSweepsButton = $window.FindName('ImportSweepsButton')
$exportSweepsButton = $window.FindName('ExportSweepsButton')

Update-SplashProgress -PercentComplete 50 -StatusText "Configuring components..."

# Set initial tab at startup
if ($tabControl) {
    $tabControl.SelectedIndex = 0
} else {
    Write-Log "Warning: Could not find TabControl element" -isError
}

# Initialize the form with default values
$startDatePicker.SelectedDate = (Get-Date).AddDays(-14).Date
$endDatePicker.SelectedDate = (Get-Date).AddDays(-7).Date

# Optimized function to filter email items with better performance
function Filter-EmailItems {
    param(
        [object]$Items,
        [DateTime]$StartDate,
        [DateTime]$EndDate,
        [string]$SenderName,
        [string]$SenderEmail,
        [string]$Subject,
        [switch]$PreviewMode
    )
    
    # Status tracking
    $startTime = Get-Date
    $processedCount 
    $matchCount = 0
    $progressInterval = 50  # Report progress every 50 items
    
    # Create date filter - use Outlook's native filtering for dates
    $dateFilter = "[ReceivedTime] >= '" + $StartDate.ToString("MM/dd/yyyy") + 
                "' AND [ReceivedTime] < '" + $EndDate.AddDays(1).ToString("MM/dd/yyyy") + "'"
    
    Write-Log "Applying date filter..."
    $dateFilteredItems = $Items.Restrict($dateFilter)
    $totalItems = $dateFilteredItems.Count
    Write-Log "Found $totalItems items in date range $($StartDate.ToString('MM/dd/yyyy')) to $($EndDate.ToString('MM/dd/yyyy'))"
    
    # For other filters, create optimized conditions first
    $checkSender = -not [string]::IsNullOrWhiteSpace($SenderName)
    $checkEmail = -not [string]::IsNullOrWhiteSpace($SenderEmail)
    $checkSubject = -not [string]::IsNullOrWhiteSpace($Subject)
    $searchEmail = ""
    
    if ($checkEmail) {
        $searchEmail = $SenderEmail.Trim().ToLower()
    }
    
    # Create a list to hold matching items - ArrayList is faster than array for additions
    $matchingItems = New-Object System.Collections.ArrayList
    
    # Process in batches to improve UI responsiveness
    $batchSize = 50
    $itemCount = $dateFilteredItems.Count
    
    Write-Log "Filtering emails with criteria..."
    if ($checkSender) { Write-Log "- Sender name contains: $SenderName" }
    if ($checkEmail) { Write-Log "- Email address equals: $searchEmail" }
    if ($checkSubject) { Write-Log "- Subject contains: $Subject" }
    
    # Process items optimally
    for ($i = 1; $i -le $itemCount; $i++) {
        try {
            $item = $dateFilteredItems.Item($i)
            $includeItem = $true
            
            # Apply filters in most efficient order (fastest checks first)
            # Subject check is usually fastest
            if ($checkSubject -and $includeItem) {
                $includeItem = $item.Subject -like "*$Subject*"
            }
            
            # Sender name check is next fastest
            if ($checkSender -and $includeItem) {
                $includeItem = $item.SenderName -like "*$SenderName*"
            }
            
            # Email address check is most expensive - do it last
            if ($checkEmail -and $includeItem) {
                $matchFound = $false
                try {
                    # Get the sender object
                    $sender = $item.Sender
                    
                    # Method for Exchange users
                    if ($sender -and $sender.AddressEntryUserType -eq 0) { # olExchangeUserAddressEntry = 0
                        $exchangeUser = $sender.GetExchangeUser()
                        if ($exchangeUser) {
                            $smtpAddress = $exchangeUser.PrimarySmtpAddress
                            if ($smtpAddress -and $smtpAddress.ToLower() -eq $searchEmail) {
                                $matchFound = $true
                            }
                        }
                    }
                    # For non-Exchange users
                    elseif ($sender) {
                        $smtpAddress = $sender.Address
                        if ($smtpAddress -and $smtpAddress.ToLower() -eq $searchEmail) {
                            $matchFound = $true
                        }
                    }
                }
                catch {
                    # Silent continue on SMTP resolution errors
                }
                
                $includeItem = $matchFound
            }
            
            # If item passed all filters, add it
            if ($includeItem) {
                $matchCount++
                
                # In preview mode, only collect first 10
                if ($PreviewMode -and $matchingItems.Count -ge 10) {
                    # We've found at least 10 matches but keep counting total matches
                } else {
                    $null = $matchingItems.Add($item)
                }
            }
            
            # Update progress and process UI events periodically
            $processedCount++
            if ($processedCount % $progressInterval -eq 0) {
                $percentComplete = [math]::Min(100, [math]::Round(($processedCount / $itemCount) * 100))
                $elapsedTime = (Get-Date) - $startTime
                $estimatedTimeRemaining = ""
                
                if ($processedCount -gt 0) {
                    $itemsPerSecond = $processedCount / $elapsedTime.TotalSeconds
                    if ($itemsPerSecond -gt 0) {
                        $remainingItems = $itemCount - $processedCount
                        $remainingSeconds = $remainingItems / $itemsPerSecond
                        $estimatedTimeRemaining = " - Est. remaining: " + [timespan]::FromSeconds($remainingSeconds).ToString("mm\:ss")
                    }
                }
                
                Write-Log "Processing: $percentComplete% complete ($processedCount of $itemCount)$estimatedTimeRemaining"
                [System.Windows.Forms.Application]::DoEvents()
            }
        }
        catch {
            # Log error but continue processing
            Write-Log ("Error processing item " + $i + ": " + $_.Exception.Message) -isError        }
    }
    
    $elapsedTime = (Get-Date) - $startTime
    Write-Log "Filter completed in $($elapsedTime.TotalSeconds.ToString('0.00')) seconds. Found $matchCount matching items."
    
    # Return both the matching items and the total count - FIXED INDENTATION HERE
    return @{
        Items = $matchingItems
        TotalCount = $matchCount
    }
}

Update-SplashProgress -PercentComplete 60 -StatusText "Setting up event handlers..."

<# Function to validate all search inputs
function Test-EmailSweepInputs {
    param(
        [DateTime]$StartDate,
        [DateTime]$EndDate,
        [string]$SenderName,
        [string]$SenderEmail,
        [string]$Subject,
        [string]$SearchFolder,
        [string]$TargetFolder
    )
    
    # Validate dates
    if (-not $StartDate -or -not $EndDate) {
        Write-Log "Error: Please select both start and end dates." -isError
        return $false
    }
    
    if ($EndDate -lt $StartDate) {
        Write-Log "Error: End date must be later than start date." -isError
        return $false
    }
    
    # Validate search criteria
    if ([string]::IsNullOrWhiteSpace($SenderName) -and 
        [string]::IsNullOrWhiteSpace($SenderEmail) -and 
        [string]::IsNullOrWhiteSpace($Subject)) {
        Write-Log "Error: Please enter a sender name, email address, or subject text." -isError
        return $false
    }
    
    # Validate folders
    if ([string]::IsNullOrWhiteSpace($SearchFolder)) {
        Write-Log "Error: Please select or enter a search folder." -isError
        return $false
    }
    
    if ([string]::IsNullOrWhiteSpace($TargetFolder)) {
        Write-Log "Error: Please select or enter a destination folder." -isError
        return $false
    }
    
    # Check if search and destination folders are the same - simplified approach
    if ($SearchFolder -eq $TargetFolder) {
        Write-Log "Error: Search and destination folders cannot be the same." -isError
        return $false
    }
    
    # All validations passed
    return $true
}
#>

# Function to test email sweep without moving emails - updated version
# Modify the Test-EmailSweepInputs function
function Test-EmailSweepInputs {
    param(
        [hashtable]$Parameters
    )
    
    # Validate dates
    if (-not $Parameters.StartDate -or -not $Parameters.EndDate) {
        Write-Log "Error: Please select both start and end dates." -isError
        return $false
    }
    
    if ($Parameters.EndDate -lt $Parameters.StartDate) {
        Write-Log "Error: End date must be later than start date." -isError
        return $false
    }
    
    # Validate search criteria
    if ([string]::IsNullOrWhiteSpace($Parameters.SenderName) -and 
        [string]::IsNullOrWhiteSpace($Parameters.SenderEmail) -and 
        [string]::IsNullOrWhiteSpace($Parameters.Subject)) {
        Write-Log "Error: Please enter a sender name, email address, or subject text." -isError
        return $false
    }
    
    # Validate folders
    if ([string]::IsNullOrWhiteSpace($Parameters.SearchFolderName)) {
        Write-Log "Error: Please select or enter a search folder." -isError
        return $false
    }
    
    if ([string]::IsNullOrWhiteSpace($Parameters.TargetFolderName)) {
        Write-Log "Error: Please select or enter a destination folder." -isError
        return $false
    }
    
    # Check if search and destination folders are the same - simplified approach
    if ($Parameters.SearchFolderName -eq $Parameters.TargetFolderName) {
        Write-Log "Error: Search and destination folders cannot be the same." -isError
        return $false
    }
    
    # All validations passed
    return $true
}

# Add this function to your script
function Test-EmailSweep {
    # Clear the log first
    Clear-LogOutput $logOutputTextBox
    
    # Get parameters from the UI
    $parameters = Get-EmailSweepParameters
    
    # Validate all inputs
    if (-not (Test-EmailSweepInputs -Parameters $parameters)) {
        return
    }
    
    Write-Log "Starting email search (preview mode)..."
    
    # Log the search criteria
    Write-SweepParameters -Parameters $parameters
    
    # Initialize Outlook
    $outlookObjects = Initialize-OutlookConnection
    if (-not $outlookObjects) { return }
    
    try {
        # Get source folder only - we don't need target folder for testing
        $namespace = $outlookObjects.Namespace
        $searchFolder = Get-OutlookFolder -FolderPath $parameters.SearchFolderName -Namespace $namespace
        
        if (-not $searchFolder) {
            Write-Log "Error: Search folder '$($parameters.SearchFolderName)' not found." -isError
            return
        }
        
        # Create a folder objects hashtable just with the search folder
        $folderObjects = @{
            SearchFolder = $searchFolder
        }
        
        # Find matching emails with preview mode
        $items = $searchFolder.Items
        $items.Sort("[ReceivedTime]", $true)
        
        # Use Filter-EmailItems function with PreviewMode switch
        $result = Filter-EmailItems -Items $items -StartDate $parameters.StartDate -EndDate $parameters.EndDate `
                  -SenderName $parameters.SenderName -SenderEmail $parameters.SenderEmail -Subject $parameters.Subject `
                  -PreviewMode
        
        $matchingItems = $result.Items
        $totalMatches = $result.TotalCount
        
        # Show preview information
        Write-Log ""
        Write-Log "PREVIEW: Found $totalMatches emails that match your criteria" -isBold
        if ($totalMatches > 10) {
            Write-Log "Showing first 10 results:"
        }
        Write-Log ""
        
        # Format table header - moved here to appear after the preview message
        if ($totalMatches -gt 0) {
            Write-TableHeader
        }
        
        # Display matching items (limited to 10 in preview mode)
        foreach ($item in $matchingItems) {
            # Format display
            $sender = $item.SenderName
            $emailSubject = $item.Subject
            $receivedTime = $item.ReceivedTime.ToString("MM/dd/yyyy HH:mm")
            
            Write-TableRow -Column1 $sender -Column2 $emailSubject -Column3 $receivedTime
        }
        
        # Show summary
        Write-Log ""
        if ($totalMatches -gt 10) {
            Write-Log "... and $($totalMatches - 10) more emails"
        }
        Write-Log ""
        Write-Log "This would move $totalMatches emails from '$($parameters.SearchFolderName)' to '$($parameters.TargetFolderName)'"
        Write-Log "Click 'Execute' to perform the actual move operation." -isBold
    }
    catch {
        Write-Log "Error during test: $($_.Exception.Message)" -isError
        Write-Log $_.ScriptStackTrace
    }
    finally {
        # Release COM objects
        Clear-OutlookObjects -OutlookObjects $outlookObjects -FolderObjects $folderObjects -EmailResults @{Items = $matchingItems}
    }
}

# Function to load categories from Outlook - now includes progress update and color information
function Load-OutlookCategories {
    param(
        [System.Windows.Controls.ComboBox]$CategoryComboBox
    )

    # Clear existing items
    $CategoryComboBox.Items.Clear()
    
    try {
        # Create Outlook COM object
        $outlook = New-Object -ComObject Outlook.Application
        $namespace = $outlook.GetNamespace("MAPI")
        
        # Create an array to collect all category names
        $categoryNames = New-Object System.Collections.ArrayList
        
        # Collect all category names
        foreach ($category in $namespace.Categories) {
            [void]$categoryNames.Add($category.Name)
        }
        
        # Sort the category names alphabetically
        $sortedCategories = $categoryNames | Sort-Object
        
        # Add "(None)" as the first option
        $CategoryComboBox.Items.Add("(None)")
        
        # Add the sorted category names
        foreach ($name in $sortedCategories) {
            $CategoryComboBox.Items.Add($name)
        }
        
        # Select "(None)" by default
        $CategoryComboBox.SelectedIndex = 0
        
        # Log success
        Write-Log $logOutputTextBox "Categories loaded successfully."
    }
    catch {
        Write-Log $logOutputTextBox "Error loading categories: $($_.Exception.Message)" -isError
    }
    finally {
        # Release COM objects
        if ($null -ne $namespace) {
            try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($namespace) | Out-Null } catch {}
        }
        if ($null -ne $outlook) {
            try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) | Out-Null } catch {}
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}
function New-OutlookCategory {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]$CategoryName,
        
        [Parameter()]
        [int]$ColorValue = 0,
        
        [Parameter()]
        [System.Windows.Controls.TextBox]$LogTextBox
    )
    
    # Guard against empty category name
    if ([string]::IsNullOrWhiteSpace($CategoryName) -or $CategoryName -eq "(None)") {
        Write-Log -logTextBox $LogTextBox -message "No category name provided" -isError
        return $false
    }
    
    try {
        # Create Outlook COM object
        $outlook = New-Object -ComObject Outlook.Application
        $categories = $outlook.Session.Categories
        
        # Check if category already exists
        $categoryExists = $false
        foreach ($existingCategory in $categories.Names) {
            if ($existingCategory -eq $CategoryName) {
                $categoryExists = $true
                Write-Log -logTextBox $LogTextBox -message "Category '$CategoryName' already exists"
                break
            }
        }
        
        # If category doesn't exist, create it with the selected color
        if (-not $categoryExists) {
            # Use the stored color value, default to Blue (8) if not set
            $colorToUse = if ($ColorValue -gt 0 -and $ColorValue -le 25) { $ColorValue } else { 8 }
            
            # Create the category with the selected color
            # Parameters: Name, Color, ShortcutKey (0 = no shortcut)
            $categories.Add($CategoryName, $colorToUse, 0)
            
            Write-Log -logTextBox $LogTextBox -message "Created category '$CategoryName' with color value $colorToUse"
            return $true
        }
        
        return $categoryExists
    }
    catch {
        Write-Log -logTextBox $LogTextBox -message "Error creating category: $($_.Exception.Message)" -isError
        return $false
    }
    finally {
        # Release COM objects
        if ($null -ne $outlook) {
            try { 
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) | Out-Null 
            } catch {}
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}



# Replace the Load-OutlookFolders function with this enhanced version

function Load-OutlookFolders {
    if ($logOutputTextBox) {
        # Use the Clear-LogOutput function which already knows how to handle RichTextBox
        Clear-LogOutput $logOutputTextBox
        
        # Use Write-Log instead of direct AppendText
        Write-Log "Loading Outlook folders..."
    }
    
    $searchFolderComboBox.Items.Clear()
    $folderComboBox.Items.Clear()
    
    $outlook = $null
    $namespace = $null
    $inbox = $null
    
    try {
        Update-SplashProgress -PercentComplete 65 -StatusText "Loading Outlook folders..."
        $outlook = New-Object -ComObject Outlook.Application
        $namespace = $outlook.GetNamespace("MAPI")
        $inbox = $namespace.GetDefaultFolder(6)  # 6 = Inbox
        
        # Create an array to hold all folder paths
        $folderPaths = New-Object System.Collections.ArrayList
        
        # Recursive function to collect folders with proper path
        function Add-FoldersRecursively {
            param (
                [Parameter(Mandatory=$true)]
                $ParentFolder,
                
                [Parameter(Mandatory=$false)]
                [string]$PathPrefix = ""
            )
            
            foreach ($folder in $ParentFolder.Folders) {
                # Build the path with the prefix and current folder name
                $folderPath = if ([string]::IsNullOrEmpty($PathPrefix)) {
                    $folder.Name
                } else {
                    "$PathPrefix\$($folder.Name)"
                }
                
                # Add the folder path to our collection
                [void]$folderPaths.Add($folderPath)
                
                # Recursively collect subfolders
                Add-FoldersRecursively -ParentFolder $folder -PathPrefix $folderPath
            }
        }
        
        # Start recursive collection from the Inbox
        Add-FoldersRecursively -ParentFolder $inbox
        
        # Sort the folder paths alphabetically
        $sortedPaths = $folderPaths | Sort-Object
        
        # Add the inbox itself first to both dropdowns
        $searchFolderComboBox.Items.Add("Inbox")
        $folderComboBox.Items.Add("Inbox")
        
        # Then add all sorted folders
        foreach ($path in $sortedPaths) {
            $searchFolderComboBox.Items.Add($path)
            $folderComboBox.Items.Add($path)
        }
        
        # Set default selections
        $searchFolderComboBox.SelectedItem = "Inbox"
        $folderComboBox.SelectedIndex = 0
        
        if ($logOutputTextBox) {
            $logOutputTextBox.AppendText("Folders loaded successfully." + [Environment]::NewLine)
        }
    }
    catch {
        if ($logOutputTextBox) {
            $logOutputTextBox.AppendText("Error loading folders: $($_.Exception.Message)" + [Environment]::NewLine)
        }
    }
    finally {
        # Release all COM objects in reverse order of creation
        try {
            # Release each item in the collection
            if ($itemsToMove -and $itemsToMove.Count -gt 0) {
                foreach ($item in $itemsToMove) {
                    Release-ComObject $item
                }
            }
            
            # Clear the array reference
            $itemsToMove = $null
            
            # Release main COM objects
            Release-ComObject $filteredItems
            Release-ComObject $items
            Release-ComObject $targetFolder
            Release-ComObject $searchFolder
            Release-ComObject $inbox
            Release-ComObject $namespace
            Release-ComObject $outlook
            
            # Force aggressive garbage collection
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
            [System.GC]::Collect()
        }
        catch {
            # Silent cleanup errors
        }
    }
}

# Helper function for logging
# Updated Write-Log function to handle both parameter styles

# Replace the entire Write-Log function with this version:

# Helper function for logging
function Write-Log {
    param(
        [Parameter(Position=0)]
        [object]$TextBoxOrMessage,
        
        [Parameter(Position=1)]
        [string]$Message,
        
        [switch]$isError,
        
        [switch]$isBold
    )
    
    # Initialize the textbox and message variables
    $textBox = $null
    $messageText = ""
    
    # Determine if we're being called with just the message or with textbox and message
    if ($TextBoxOrMessage -is [System.Windows.Controls.TextBox]) {
        # First parameter is a TextBox object
        $textBox = $TextBoxOrMessage
        $messageText = $Message
    }
    elseif ($TextBoxOrMessage -is [string]) {
        # First parameter is the message string
        $textBox = $global:logOutputTextBox
        $messageText = $TextBoxOrMessage
    }
    else {
        # Use globals as fallback
        $textBox = $global:logOutputTextBox
        $messageText = if ($Message) { $Message } else { "Log message" }
    }
    
    # Make sure we have a valid TextBox and message
    if ($null -eq $textBox) {
        # Fallback to console output if no valid TextBox
        if ($isError) {
            Write-Host "[ERROR] $messageText" -ForegroundColor Red
        }
        elseif ($isBold) {
            # For console output, we can use asterisks to indicate boldness
            Write-Host "**$messageText**" 
        }
        else {
            Write-Host $messageText
        }
        return
    }
    
    # Now we're sure we have a TextBox object
    try {
        # Prepare the message text with appropriate prefix
        $displayText = if ($isError) {
            "[ERROR] " + $messageText
        } else {
            $messageText
        }
        
        # Create a new paragraph and run for this message
        $paragraph = New-Object System.Windows.Documents.Paragraph
        $run = New-Object System.Windows.Documents.Run
        $run.Text = $displayText
        $run.FontFamily = New-Object System.Windows.Media.FontFamily("Consolas")  # Force monospaced font

        # Apply bold formatting if requested
        if ($isBold) {
            $run.FontWeight = "Bold"
        }

        # Apply error formatting if it's an error
        if ($isError) {
            $run.Foreground = "Red"
        }

        # Add the run to the paragraph
        $paragraph.Inlines.Add($run)
        $paragraph.LineHeight = 1.2
        $paragraph.TextAlignment = "Left"  # Ensure left alignment for tables

        # If the TextBox is a RichTextBox, add the paragraph to the document
        if ($textBox -is [System.Windows.Controls.RichTextBox]) {
            $textBox.Document.Blocks.Add($paragraph)
        }
        else {
            # For regular TextBox, just append the text without formatting
            $textBox.AppendText($displayText + [Environment]::NewLine)
        }
        
        # Force scroll to end
        $textBox.ScrollToEnd()
        
        # Force UI update
        if ($textBox.Dispatcher) {
            $textBox.Dispatcher.Invoke([Action]{}, [Windows.Threading.DispatcherPriority]::Render)
        }
        [System.Windows.Forms.Application]::DoEvents()
    }
    catch {
        # Fallback to console output if TextBox methods fail
        if ($isError) {
            Write-Host "[ERROR] $messageText" -ForegroundColor Red
        }
        elseif ($isBold) {
            Write-Host "**$messageText**"
        }
        else {
            Write-Host $messageText
        }
    }
}

# Function for creating consistently formatted tables in RichTextBox
function Format-TableCell {
    param(
        [string]$Text,
        [int]$Width,
        [switch]$TruncateWithEllipsis
    )
    
    if ($TruncateWithEllipsis -and $Text.Length -gt $Width) {
        # Truncate with ellipsis
        return $Text.Substring(0, $Width-1) + "…"
    } 
    elseif ($Text.Length -gt $Width) {
        # Just truncate
        return $Text.Substring(0, $Width)
    }
    else {
        # Pad with spaces to ensure consistent width
        return $Text.PadRight($Width)
    }
}

# Improved table header with proper formatting for RichTextBox
function Write-TableHeader {
    param(
        [string]$Column1 = "From",
        [string]$Column2 = "Subject",
        [string]$Column3 = "Received Time",
        [int]$Width1 = 30,   # Changed from 20 to 30 to widen the From column
        [int]$Width2 = 45,
        [int]$Width3 = 20
    )
    
    # Create dashed lines for separator
    $dash1 = "-" * $Width1
    $dash2 = "-" * $Width2
    $dash3 = "-" * $Width3
    
    # Write header row with bold formatting
    Write-TableRow -Column1 $Column1 -Column2 $Column2 -Column3 $Column3 -Width1 $Width1 -Width2 $Width2 -Width3 $Width3 -isBold
    
    # Write separator line
    Write-TableRow -Column1 $dash1 -Column2 $dash2 -Column3 $dash3 -Width1 $Width1 -Width2 $Width2 -Width3 $Width3
}

# Improved table row with proper formatting for RichTextBox
function Write-TableRow {
    param(
        [string]$Column1 = "",
        [string]$Column2 = "",
        [string]$Column3 = "",
        [int]$Width1 = 30,   # Changed from 20 to 30 to widen the From column
        [int]$Width2 = 45,
        [int]$Width3 = 20,
        [switch]$isBold
    )
    
    # Format each cell with proper width and truncation
    $col1Text = Format-TableCell -Text $Column1 -Width $Width1 -TruncateWithEllipsis
    $col2Text = Format-TableCell -Text $Column2 -Width $Width2 -TruncateWithEllipsis
    $col3Text = Format-TableCell -Text $Column3 -Width $Width3 -TruncateWithEllipsis
    
    # Combine columns with spacing - use tabs for consistent alignment
    $rowText = "{0} {1} {2}" -f $col1Text, $col2Text, $col3Text
    
    # Write the row with optional bold formatting
    Write-Log $rowText -isBold:$isBold
}

# Function to update log window - updated for RichTextBox
function Clear-LogOutput {
    param($textBox)
    
    if ($textBox -is [System.Windows.Controls.RichTextBox]) {
        $textBox.Document = New-Object System.Windows.Documents.FlowDocument
        $textBox.Document.PageWidth = 1000  # Set a large page width to avoid unnecessary wrapping
    } else {
        $textBox.Clear()
    }
}

# Function to execute email sweep with direct string pattern matching
# Function to execute email sweep with direct string pattern matching
# Function to execute email sweep with direct string pattern matching

function Get-EmailSweepParameters {
    return @{
        StartDate = $startDatePicker.SelectedDate
        EndDate = $endDatePicker.SelectedDate
        Subject = $subjectTextBox.Text.Trim()
        SenderName = $senderNameTextBox.Text.Trim()
        SenderEmail = $senderEmailTextBox.Text.Trim()
        SearchFolderName = $searchFolderComboBox.Text.Trim()
        TargetFolderName = $folderComboBox.Text.Trim()
        Category = $categoryComboBox.Text.Trim()
        RemoveCategories = $removeCategoriesCheckBox.IsChecked
        MarkAsRead = $markAsReadCheckBox.IsChecked
    }
}

function Write-SweepParameters {
    param(
        [hashtable]$Parameters
    )
    
    Write-Log "Search criteria:"
    Write-Log "- Date range: $($Parameters.StartDate.ToString('MM/dd/yyyy')) to $($Parameters.EndDate.ToString('MM/dd/yyyy'))"
    
    if (-not [string]::IsNullOrWhiteSpace($Parameters.SenderName)) {
        Write-Log "- Sender name contains: $($Parameters.SenderName)"
    }
    if (-not [string]::IsNullOrWhiteSpace($Parameters.SenderEmail)) {
        Write-Log "- Email address equals: $($Parameters.SenderEmail)"
    }
    if (-not [string]::IsNullOrWhiteSpace($Parameters.Subject)) {
        Write-Log "- Subject contains: $($Parameters.Subject)"
    }
    if (-not [string]::IsNullOrWhiteSpace($Parameters.Category) -and $Parameters.Category -ne "(None)") {
        Write-Log "- Apply category: $($Parameters.Category)"
    }
    if ($Parameters.RemoveCategories) {
        Write-Log "- Remove existing categories: Yes"
    }
    if ($Parameters.MarkAsRead) {
        Write-Log "- Mark emails as read: Yes"
    }
}

function Initialize-OutlookConnection {
    try {
        $outlook = New-Object -ComObject Outlook.Application
        $namespace = $outlook.GetNamespace("MAPI")
        
        return @{
            Outlook = $outlook
            Namespace = $namespace
        }
    }
    catch {
        Write-Log "Error: Could not establish connection to Outlook. $($_.Exception.Message)" -isError
        return $null
    }
}

# Function to handle folder lookup and creation consistently with V5
function Get-OutlookFolder {
    param(
        [string]$FolderPath,
        $Namespace,
        [switch]$CreateIfMissing
    )
    
    try {
        # Handle Inbox as a special case
        if ($FolderPath -eq "Inbox") {
            return $Namespace.GetDefaultFolder(6) # 6 = Inbox
        }
        
        # Simple approach: try to get the folder directly
        try {
            $inbox = $Namespace.GetDefaultFolder(6)
            $targetFolder = $inbox
            
            # If folder path contains backslashes, it's a subfolder path
            if ($FolderPath.Contains('\')) {
                $folderParts = $FolderPath -split '\\'
                foreach ($part in $folderParts) {
                    $targetFolder = $targetFolder.Folders.Item($part)
                }
            } else {
                # Simple single-level folder
                $targetFolder = $inbox.Folders.Item($FolderPath)
            }
            
            return $targetFolder
        }
        catch {
            # Folder not found - continue only if CreateIfMissing is specified
            if (-not $CreateIfMissing) {
                Write-Log "Error: Folder '$FolderPath' not found." -isError
                return $null
            }
        }
        
        # If we get here, the folder wasn't found and CreateIfMissing is true
        Write-Log "Folder '$FolderPath' not found."
        
        # Show dialog but don't capture return value (V5 approach)
        $script:DialogUserChoice = $false
        Show-CustomDialog -Title "Folder Not Found" `
            -Message "Folder '$FolderPath' not found. Would you like to create this folder or cancel the sweep execution?" `
            -OkButtonText "Create" -CancelButtonText "Cancel"
        
        # Check the script-level variable immediately after dialog
        if ($script:DialogUserChoice -ne $true) {
            Write-Log "Operation cancelled. Folder not created."
            return $null
        }
        
        # User wants to create the folder
        try {
            Write-Log "Creating folder: $FolderPath..."
            $inbox = $Namespace.GetDefaultFolder(6)
            
            # Handle multi-level folders
            if ($FolderPath.Contains('\')) {
                $folderParts = $FolderPath -split '\\'
                $currentFolder = $inbox
                
                foreach ($part in $folderParts) {
                    # Try to get the folder first (it might exist)
                    try {
                        $subFolder = $currentFolder.Folders.Item($part)
                    }
                    catch {
                        # Create if it doesn't exist
                        $subFolder = $currentFolder.Folders.Add($part)
                        Write-Log "Created folder: $part"
                    }
                    
                    $currentFolder = $subFolder
                }
                
                return $currentFolder
            }
            else {
                # Simple single-level folder
                $newFolder = $inbox.Folders.Add($FolderPath)
                Write-Log "Created folder: $FolderPath"
                return $newFolder
            }
        }
        catch {
            Write-Log "Error creating folder '$FolderPath': $($_.Exception.Message)" -isError
            return $null
        }
    }
    catch {
        Write-Log "Error accessing folder '$FolderPath': $($_.Exception.Message)" -isError
        return $null
    }
}

function Process-ReadStatus {
    param([object]$Item)
    
    try {
        $Item.UnRead = $false
    }
    catch {
        Write-Log "Warning: Could not mark item as read: $($_.Exception.Message)" -isError
    }
}

function Process-Categories {
    param(
        [object]$Item,
        [string]$Category,
        [bool]$RemoveCategories
    )
    
    try {
        # Skip if no category changes needed
        if (-not $RemoveCategories -and ($Category -eq "(None)" -or [string]::IsNullOrWhiteSpace($Category))) {
            return
        }
        
        # Remove existing categories if requested
        if ($RemoveCategories) {
            $Item.Categories = ""
        }
        
        # Apply new category if specified
        if ($Category -ne "" -and $Category -ne "(None)") {
            if (-not [string]::IsNullOrWhiteSpace($Item.Categories) -and -not $RemoveCategories) {
                # Avoid duplicate categories
                $existingCategories = $Item.Categories -split ','
                if ($existingCategories -notcontains $Category) {
                    $Item.Categories = "$($Item.Categories), $Category"
                }
            } else {
                $Item.Categories = $Category
            }
        }
        
        # Save the item to apply changes
        $Item.Save()
    }
    catch {
        Write-Log "Warning: Could not apply category changes: $($_.Exception.Message)" -isError
    }
}

function Clear-OutlookObjects {
    param(
        [hashtable]$OutlookObjects,
        [hashtable]$FolderObjects,
        [object]$EmailResults
    )
    
    try {
        # Release email items first
        if ($EmailResults -and $EmailResults.Items) {
            foreach ($item in $EmailResults.Items) {
                Release-ComObject $item
            }
        }
        
        # Release folder objects
        if ($FolderObjects) {
            if ($FolderObjects.TargetFolder) { Release-ComObject $FolderObjects.TargetFolder }
            if ($FolderObjects.SearchFolder) { Release-ComObject $FolderObjects.SearchFolder }
        }
        
        # Release Outlook objects
        if ($OutlookObjects) {
            if ($OutlookObjects.Namespace) { Release-ComObject $OutlookObjects.Namespace }
            if ($OutlookObjects.Outlook) { Release-ComObject $OutlookObjects.Outlook }
        }
        
        # Force garbage collection
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
    catch {
        Write-Log "Warning: Error during COM object cleanup (non-critical)" -isError
    }
}

# Main function - now acts as a coordinator
function Execute-EmailSweep {
    param(
        [switch]$NoClearLog,
        [switch]$BatchMode
    )

    # Clear log unless NoClearLog is specified (for batch operations)
    if (-not $NoClearLog) {
        Clear-LogOutput $logOutputTextBox
    }
    
    Write-Log "Starting email sweep process..."
    
    # Get parameters from UI or function parameters
    $parameters = Get-EmailSweepParameters
    
    # Validate all inputs first
    if (-not (Test-EmailSweepInputs -Parameters $parameters)) {
        return $null
    }
    
    # Log what we're going to do
    Write-SweepParameters -Parameters $parameters
    
    $outlook = $null
    $namespace = $null
    $searchFolder = $null
    $targetFolder = $null
    $items = $null
    $itemsToMove = $null
    $movedCount = 0
    
    try {
        # Initialize Outlook
        $outlook = New-Object -ComObject Outlook.Application
        $namespace = $outlook.GetNamespace("MAPI")
        
        #---------------------------
        # CATEGORY HANDLING - Similar to V5's direct approach
        #---------------------------
        if (-not [string]::IsNullOrWhiteSpace($parameters.Category) -and $parameters.Category -ne "(None)") {
            $categories = $outlook.Session.Categories
            $categoryExists = $false
            
            foreach ($existingCategory in $categories.Names) {
                if ($existingCategory -eq $parameters.Category) {
                    $categoryExists = $true
                    break
                }
            }
            
            if (-not $categoryExists) {
                Write-Log "Category '$($parameters.Category)' does not exist. Prompting to create..."
                
                # Reset the dialog variable BEFORE showing dialog
                $script:DialogUserChoice = $false
                
                # Show dialog and immediately check result - V5 style
                Show-CustomDialog -Title "Category Not Found" `
                    -Message "Category '$($parameters.Category)' does not exist. Would you like to create this category or cancel the sweep execution?" `
                    -OkButtonText "Create" -CancelButtonText "Cancel"
                
                # If user canceled, stop execution immediately
                if ($script:DialogUserChoice -ne $true) {
                    Write-Log "Operation cancelled. Category not created."
                    return $null
                }
                
                # User wants to create the category - do it right here, not in another function
                try {
                    Write-Log "Creating new category: $($parameters.Category)..."
                    $categories.Add($parameters.Category, 8, 0)
                    Write-Log "Category created successfully."
                }
                catch {
                    Write-Log "Error creating category: $($_.Exception.Message)" -isError
                    return $null
                }
            }
        }
        
        #---------------------------
        # SOURCE FOLDER HANDLING
        #---------------------------
        try {
            if ($parameters.SearchFolderName -eq "Inbox") {
                $searchFolder = $namespace.GetDefaultFolder(6) # 6 = Inbox
            } else {
                $inbox = $namespace.GetDefaultFolder(6)
                $searchFolderParts = $parameters.SearchFolderName -split '\\'
                $searchFolder = $inbox
                
                # Navigate folder structure
                foreach ($part in $searchFolderParts) {
                    $searchFolder = $searchFolder.Folders.Item($part)
                }
            }
        }
        catch {
            Write-Log "Error: Source folder '$($parameters.SearchFolderName)' not found." -isError
            return $null
        }
        
        #---------------------------
        # TARGET FOLDER HANDLING - Direct approach like V5
        #---------------------------
        try {
            if ($parameters.TargetFolderName -eq "Inbox") {
                $targetFolder = $namespace.GetDefaultFolder(6) # 6 = Inbox
            } else {
                $inbox = $namespace.GetDefaultFolder(6)
                $folderParts = $parameters.TargetFolderName -split '\\'
                $targetFolder = $inbox
                
                # Try to navigate to the folder
                foreach ($part in $folderParts) {
                    $targetFolder = $targetFolder.Folders.Item($part)
                }
            }
        }
        catch {
            # Folder doesn't exist, ask to create - V5 style
            Write-Log "Destination folder '$($parameters.TargetFolderName)' not found."
            
            # Reset the dialog choice BEFORE showing dialog - critical
            $script:DialogUserChoice = $false
            
            # Show dialog but DON'T store return value
            Show-CustomDialog -Title "Folder Not Found" `
                -Message "Destination folder '$($parameters.TargetFolderName)' not found. Would you like to create this folder or cancel the sweep execution?" `
                -OkButtonText "Create" -CancelButtonText "Cancel"
            
            # Check the script-level variable immediately
            if ($script:DialogUserChoice -ne $true) {
                Write-Log "Operation cancelled. Folder not created."
                return $null
            }
            
            # User clicked Create, so create the folder
            Write-Log "Creating folder path: Inbox\$($parameters.TargetFolderName)..."
            
            # Create each part of the folder path
            $targetFolder = $inbox
            foreach ($part in $folderParts) {
                try {
                    # Try to get the folder first in case it exists
                    try {
                        $subFolder = $targetFolder.Folders.Item($part)
                    }
                    catch {
                        # Folder doesn't exist, create it
                        $subFolder = $targetFolder.Folders.Add($part)
                        Write-Log "Created folder: $part"
                    }
                    
                    # Move to next level
                    $targetFolder = $subFolder
                }
                catch {
                    Write-Log "Error creating folder '$part': $($_.Exception.Message)" -isError
                    return $null
                }
            }
        }
        
        #---------------------------
        # FIND MATCHING EMAILS
        #---------------------------
        # Get all items from the search folder
        $items = $searchFolder.Items
        $items.Sort("[ReceivedTime]", $true)
        
        # Use Filter-EmailItems function to find matches
        $result = Filter-EmailItems -Items $items -StartDate $parameters.StartDate -EndDate $parameters.EndDate `
                  -SenderName $parameters.SenderName -SenderEmail $parameters.SenderEmail -Subject $parameters.Subject
        
        $itemsToMove = $result.Items
        $totalMatches = $result.TotalCount
        
        Write-Log "Found $totalMatches matching emails to move..."
        
        # Display header for the results if we have matches
        if ($totalMatches -gt 0) {
            Write-TableHeader
        }
        
        #---------------------------
        # PROCESS AND MOVE EMAILS
        #---------------------------
        foreach ($item in $itemsToMove) {
            # Format display with proper columns
            $sender = $item.SenderName
            $emailSubject = $item.Subject
            $receivedTime = $item.ReceivedTime.ToString("MM/dd/yyyy HH:mm")
            
            Write-TableRow -Column1 $sender -Column2 $emailSubject -Column3 $receivedTime
            
            # Apply Read status if requested
            if ($parameters.MarkAsRead) {
                Process-ReadStatus -Item $item
            }
                        
            # Apply category changes if needed
            Process-Categories -Item $item -Category $parameters.Category -RemoveCategories $parameters.RemoveCategories
            
            # Move email
            $item.Move($targetFolder) | Out-Null
            $movedCount++
            
            # Update progress periodically
            if ($movedCount % 10 -eq 0) {
                $percentComplete = [math]::Min(100, [math]::Round(($movedCount / $totalMatches) * 100))
                Write-Log "Moving emails: $percentComplete% complete ($movedCount of $totalMatches)"
                [System.Windows.Forms.Application]::DoEvents()
            }
        }
        
        # Return result information for summary reports
        return @{
            SweepName = if ($BatchMode) { $parameters.Name } else { "Manual Sweep" }
            MovedCount = $movedCount
            SearchFolder = $parameters.SearchFolderName
            TargetFolder = $parameters.TargetFolderName
        }
    }
    catch {
        Write-Log "Error: $($_.Exception.Message)" -isError
        Write-Log $_.ScriptStackTrace
        return @{ MovedCount = $movedCount }
    }
    finally {
        # Release all COM objects in reverse order
        try {
            # Release email items
            if ($itemsToMove) {
                foreach ($item in $itemsToMove) {
                    Release-ComObject $item
                }
            }
            
            Release-ComObject $items
            Release-ComObject $targetFolder
            Release-ComObject $searchFolder
            Release-ComObject $namespace
            Release-ComObject $outlook
            
            # Force garbage collection
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
        }
        catch {
            # Silent cleanup errors
        }
    }
}


# Function to set date range to last week
function Set-LastWeekDates {
    $today = Get-Date
    $startOfLastWeek = $today.AddDays(-7 - $today.DayOfWeek.value__)
    $endOfLastWeek = $startOfLastWeek.AddDays(6)
    
    $startDatePicker.SelectedDate = $startOfLastWeek
    $endDatePicker.SelectedDate = $endOfLastWeek
    
    Write-Log "Date range set to last week: $startOfLastWeek to $endOfLastWeek"
}

Update-SplashProgress -PercentComplete 80 -StatusText "Setting up event handlers..."

# Event handlers

# Save sweep configuration
$saveSweepButton.Add_Click({
    # Prompt for a name
    $inputDialog = [System.Windows.Window]::new()
    $inputDialog.Title = "Name Your Sweep"
    $inputDialog.Width = 400
    $inputDialog.Height = 180
    $inputDialog.WindowStartupLocation = "CenterOwner"
    $inputDialog.Owner = $window
    $inputDialog.ResizeMode = "NoResize"
    $inputDialog.Background = "#F5F5F5"
    
    $inputPanel = [System.Windows.Controls.StackPanel]::new()
    $inputPanel.Margin = "20"
    
    $label = [System.Windows.Controls.TextBlock]::new()
    $label.Text = "Enter a name for this sweep configuration:"
    $label.Margin = "0,0,0,10"
    $inputPanel.Children.Add($label)
    
    $textBox = [System.Windows.Controls.TextBox]::new()
    $textBox.Margin = "0,0,0,20"
    $textBox.Height = 30
    $textBox.FontSize = 14
    $inputPanel.Children.Add($textBox)
    
    $buttonPanel = [System.Windows.Controls.StackPanel]::new()
    $buttonPanel.Orientation = "Horizontal"
    $buttonPanel.HorizontalAlignment = "Right"
    
    $cancelBtn = [System.Windows.Controls.Button]::new()
    $cancelBtn.Content = "Cancel"
    $cancelBtn.Width = 80
    $cancelBtn.Height = 30
    $cancelBtn.Margin = "5,0,5,0"
    $cancelBtn.Add_Click({ $inputDialog.DialogResult = $false })
    $buttonPanel.Children.Add($cancelBtn)
    
    $saveBtn = [System.Windows.Controls.Button]::new()
    $saveBtn.Content = "Save"
    $saveBtn.Width = 80
    $saveBtn.Height = 30
    $saveBtn.IsDefault = $true
    $saveBtn.Add_Click({ $inputDialog.DialogResult = $true })
    $buttonPanel.Children.Add($saveBtn)
    
    $inputPanel.Children.Add($buttonPanel)
    $inputDialog.Content = $inputPanel
    
    # Show dialog and process result
    if ($inputDialog.ShowDialog()) {
        $sweepName = $textBox.Text.Trim()
        if (-not [string]::IsNullOrWhiteSpace($sweepName)) {
            Save-SweepConfiguration -Name $sweepName
        }
    }
})

# Load button handler
$loadSweepButton.Add_Click({
    $selectedItem = $savedSweepsListView.SelectedItem
    if ($selectedItem) {
        $sweep = $selectedItem.Tag
        
        # First switch to Mail Sweep tab, then load configuration
        if ($null -ne $tabControl) {
            $window.Dispatcher.Invoke([Action]{
                $tabControl.SelectedIndex = 0
            }, [Windows.Threading.DispatcherPriority]::Send)
        }
        
        Load-SweepConfiguration -sweep $sweep
        Clear-LogOutput $logOutputTextBox
        Write-Log "Loaded sweep configuration: $($sweep.Name)"
    } else {
        Show-CustomDialog -Title "No Selection" `
            -Message "Please select a sweep configuration to load." `
            -InfoOnly
    }
})

# Delete button handler
$deleteSweepButton.Add_Click({
    $selectedItem = $savedSweepsListView.SelectedItem
    if ($selectedItem) {
        $sweep = $selectedItem.Tag
        Delete-SweepConfiguration -sweep $sweep
    } else {
        [System.Windows.MessageBox]::Show("Please select a sweep configuration to delete.", "No Selection", 
            [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
    }
})

# Run button handler
$runSweepButton.Add_Click({
    $selectedItems = $savedSweepsListView.SelectedItems
    if ($selectedItems -and $selectedItems.Count -gt 0) {
        $sweeps = @()
        foreach ($item in $selectedItems) {
            if ($item.Tag) {
                $sweeps += $item.Tag
            }
        }
        
        if ($sweeps.Count -eq 0) {
            Show-CustomDialog -Title "Selection Error" -Message "No valid sweeps selected." -InfoOnly
            return
        }
        
        # Confirm execution if multiple sweeps
        $confirmMessage = "Run $($sweeps.Count) selected sweep"
        if ($sweeps.Count -gt 1) { $confirmMessage += "s" }
        $confirmMessage += " with a date range of 7-35 days old?"
        
        $result = Show-CustomDialog -Title "Confirm Run Selected" -Message $confirmMessage -YesNoButtons
        
        if ($result -eq $true) {
            # First switch to Mail Sweep tab
            if ($null -ne $tabControl) {
                $window.Dispatcher.Invoke([Action]{
                    $tabControl.SelectedIndex = 0
                }, [Windows.Threading.DispatcherPriority]::Send)
            }
            
            # Set date range to 7-35 days old for selected sweeps
            $startDatePicker.SelectedDate = (Get-Date).AddDays(-35).Date
            $endDatePicker.SelectedDate = (Get-Date).AddDays(-7).Date
            
            # Show progress information - clear log only once at the beginning
            Clear-LogOutput $logOutputTextBox
            Write-Log "Starting execution of $($sweeps.Count) selected sweeps..."
            
            # Create a counter for progress tracking and array for results
            $sweepCounter = 0
            $sweepResults = @()
            
            # Process each selected sweep in sequence
            foreach ($sweep in $sweeps) {
                $sweepCounter++
                
                # Add separator between sweeps
                Write-Log ""
                Write-Log "========================================================"
                Write-Log "[$sweepCounter of $($sweeps.Count)] Processing sweep: $($sweep.Name)"
                Write-Log "========================================================"
                Write-Log ""
                
                # Load the sweep configuration
                Load-SweepConfiguration -sweep $sweep
                
                # Execute the sweep - pass NoClearLog and BatchMode parameters
                $result = Execute-EmailSweep -NoClearLog -BatchMode
                
                # Store the result for summary
                $sweepResults += $result
                
                # Add brief pause between sweeps to ensure COM objects are properly released
                Start-Sleep -Milliseconds 500
                [System.GC]::Collect()
                [System.GC]::WaitForPendingFinalizers()
            }
            
            # Add the summary report at the end
            Write-Log ""
            Write-Log "========================================================"
            Write-Log "SUMMARY OF BATCH SWEEP OPERATIONS"
            Write-Log "========================================================"
            Write-Log "Completed execution of $($sweeps.Count) selected sweeps."
            Write-Log ""
            Write-Log "RESULTS BY SWEEP:"
            Write-Log "------------------------------------------------------"

            $totalMoved = 0
            foreach ($result in $sweepResults) {
                $totalMoved += $result.MovedCount
                Write-Log "• $($result.MovedCount) emails moved from $($result.SearchFolder) to '$($result.TargetFolder)'"
            }

            Write-Log ""
            Write-Log "========================================================" -isbBold
            Write-Log "TOTAL EMAILS MOVED: $totalMoved" -isBold
            Write-Log "========================================================" -isBold
        }
    } else {
        Show-CustomDialog -Title "No Selection" -Message "Please select at least one sweep to run." -InfoOnly
    }
})

# Double-click handler for ListView
$savedSweepsListView.Add_MouseDoubleClick({
    $selectedItem = $savedSweepsListView.SelectedItem
    if ($selectedItem) {
        $sweep = $selectedItem.Tag
        Load-SweepConfiguration -sweep $sweep
    }
})

# Add this code after the other event handlers
$savedSweepsListView.Add_SelectionChanged({
    # Update the Run Selection button text with selection count
    $selectionCount = $savedSweepsListView.SelectedItems.Count
    $runSweepButton.Content = "Run Selection ($selectionCount)"
})

$testButton.Add_Click({
    Test-EmailSweep
})

$executeButton.Add_Click({
    # Execute sweep and store the result
    $result = Execute-EmailSweep
    
    # Generate summary report
    Write-Log ""
    Write-Log "========================================================"
    Write-Log "SWEEP OPERATION SUMMARY"
    Write-Log "========================================================"
    Write-Log ""
    Write-Log "RESULTS:"
    Write-Log "========================================================"
    Write-Log "• $($result.SweepName): $($result.MovedCount) emails moved from $($result.SearchFolder) to $($result.TargetFolder)"
    Write-Log "========================================================" -isBold
    Write-Log "TOTAL EMAILS MOVED: $($result.MovedCount)" -isBold
    Write-Log "========================================================" -isBold
})

$cancelButton.Add_Click({
    # Force aggressive COM object cleanup before closing
    [System.Runtime.InteropServices.Marshal]::CleanupUnusedObjectsInCurrentContext()
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    [System.GC]::Collect()
    
    $window.Close()
})

$refreshFoldersButton.Add_Click({
    Load-OutlookFolders
})

$lastWeekButton.Add_Click({
    Set-LastWeekDates
})

# Info button event handler
$infoButton = $window.FindName('InfoButton')
$infoButton.Add_Click({
    $infoWindow = [System.Windows.Window]::new()
    $infoWindow.Title = "About Email Sweep Tool"
    $infoWindow.Width = 450
    $infoWindow.Height = 350
    $infoWindow.WindowStartupLocation = "CenterOwner"
    $infoWindow.Owner = $window
    $infoWindow.ResizeMode = "NoResize"
    $infoWindow.Background = "#F5F5F5"
    
    # Create the content
    $infoPanel = [System.Windows.Controls.StackPanel]::new()
    $infoPanel.Margin = "20,20,20,20"
    
    $infoTitle = [System.Windows.Controls.TextBlock]::new()
    $infoTitle.Text = "Email Sweep Tool"
    $infoTitle.FontSize = 18
    $infoTitle.FontWeight = "Bold"
    $infoTitle.Foreground = "#2196F3"
    $infoTitle.Margin = "0,0,0,10"
    $infoPanel.Children.Add($infoTitle)
    
    $infoText = [System.Windows.Controls.TextBlock]::new()
    $infoText.Text = "Email Sweep Tool quickly searches and organizes Outlook emails using flexible criteria. Filter by sender name, exact email address, subject text, and date range. Easily move matching messages to any folder or create a new one. Perfect for cleaning inbox clutter and organizing important communications."
    $infoText.TextWrapping = "Wrap"
    $infoText.Margin = "0,0,0,20"
    $infoText.LineHeight = 22
    $infoPanel.Children.Add($infoText)
    
    $versionText = [System.Windows.Controls.TextBlock]::new()
    $versionText.Text = "Version: 2.0"
    $versionText.Margin = "0,0,0,5"
    $versionText.Foreground = "#757575"
    $infoPanel.Children.Add($versionText)
    
    $dateText = [System.Windows.Controls.TextBlock]::new()
    $dateText.Text = "Build Date: $(Get-Date -Format 'MMMM yyyy')"
    $dateText.Foreground = "#757575"
    $infoPanel.Children.Add($dateText)
    
    # Add close button at the bottom
    $closeButton = [System.Windows.Controls.Button]::new()
    $closeButton.Content = "Close"
    $closeButton.Width = 100
    $closeButton.Height = 30
    $closeButton.Margin = "0,20,0,0"
    $closeButton.HorizontalAlignment = "Right"
    $closeButton.Add_Click({ $infoWindow.Close() })
    $infoPanel.Children.Add($closeButton)
    
    $infoWindow.Content = $infoPanel
    $infoWindow.ShowDialog()
})

# Run All Sweeps button handler
$runAllSweepsButton.Add_Click({
    # Check if there are any sweeps to run
    $noSweeps = ($savedSweepsListView.Items.Count -eq 0)
    $onlyEmptyItem = ($savedSweepsListView.Items.Count -eq 1 -and $savedSweepsListView.Items[0].IsEnabled -eq $false)
    
    if ($noSweeps -or $onlyEmptyItem) {
        Show-CustomDialog -Title "No Sweeps" `
            -Message "No saved sweeps found to run." `
            -InfoOnly
        return
    }
    
    # Confirm before running all sweeps
    $result = Show-CustomDialog -Title "Confirm Run All Sweeps" `
        -Message "This will execute all saved sweeps in sequence using a date range of 7-35 days old. Continue?" `
        -YesNoButtons
    
    if ($result -eq $true) {
        # First switch to Mail Sweep tab
        if ($null -ne $tabControl) {
            $window.Dispatcher.Invoke([Action]{
                $tabControl.SelectedIndex = 0
            }, [Windows.Threading.DispatcherPriority]::Send)
        }
        
        # Set date range to 7-35 days old for all sweeps
        $startDatePicker.SelectedDate = (Get-Date).AddDays(-35).Date
        $endDatePicker.SelectedDate = (Get-Date).AddDays(-7).Date
        
        # Get all sweep configurations (skip the empty state item if present)
        $sweeps = @()
        foreach ($item in $savedSweepsListView.Items) {
            if ($item.IsEnabled -ne $false -and $item.Tag) {
                $sweeps += $item.Tag
            }
        }
        
        # Show progress information - clear log only once at the beginning
        Clear-LogOutput $logOutputTextBox
        Write-Log "Starting batch execution of $($sweeps.Count) saved sweeps..."
        
        # Create a counter for progress tracking and array for results
        $sweepCounter = 0
        $sweepResults = @()
        
        # Process each sweep in sequence
        foreach ($sweep in $sweeps) {
            $sweepCounter++
            
            # Add separator between sweeps
            Write-Log ""
            Write-Log "========================================================"
            Write-Log "[$sweepCounter of $($sweeps.Count)] Processing sweep: $($sweep.Name)"
            Write-Log "========================================================"
            Write-Log ""
            
            # Load the sweep configuration
            Load-SweepConfiguration -sweep $sweep
            
            # Execute the sweep - pass NoClearLog and BatchMode parameters
            $result = Execute-EmailSweep -NoClearLog -BatchMode
            
            # Store the result for summary
            $sweepResults += $result
            
            # Add brief pause between sweeps to ensure COM objects are properly released
            Start-Sleep -Milliseconds 500
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
        }
        
        # Add the summary report at the end
        Write-Log ""
        Write-Log "========================================================"
        Write-Log "SUMMARY OF BATCH SWEEP OPERATIONS"
        Write-Log "========================================================"
        Write-Log "Completed execution of $($sweeps.Count) selected sweeps."
        Write-Log ""
        Write-Log "RESULTS BY SWEEP:"
        Write-Log "========================================================"

        $totalMoved = 0
        foreach ($result in $sweepResults) {
            $totalMoved += $result.MovedCount
            Write-Log "• $($result.SweepName): $($result.MovedCount) emails moved from $($result.SearchFolder) to $($result.TargetFolder)"
        }

        Write-Log ""
        Write-Log "========================================================"
        Write-Log "TOTAL EMAILS MOVED: $totalMoved" -isBold
        Write-Log "========================================================"
    }
})

# Import button handler
$importSweepsButton.Add_Click({
    # Switch to the Saved Sweeps tab first
    if ($null -ne $tabControl) {
        $window.Dispatcher.Invoke([Action]{
            $tabControl.SelectedIndex = 1 # Saved Sweeps tab
        }, [Windows.Threading.DispatcherPriority]::Send)
    }
    
    # Call the import function
    $importResult = Import-SweepConfigurations
    
    # If import was successful, show success message
    if ($importResult -eq $true) {
        Show-CustomDialog -Title "Import Successful" `
            -Message "Sweep configurations were successfully imported." `
            -InfoOnly
    }
})

# Export button handler - fixed version
$exportSweepsButton.Add_Click({
    # Get selected items
    $selectedItems = $savedSweepsListView.SelectedItems
    $hasSelection = ($selectedItems -and $selectedItems.Count -gt 0)
    
    # If there's a selection, ask if user wants to export just selected or all
    $exportAll = $false
    $sweepsToExport = @()
    
    if ($hasSelection) {
        $result = Show-CustomDialog -Title "Export Scope" `
            -Message "Do you want to export only the selected sweeps or all saved sweeps?" `
            -OkButtonText "Selected Only" -CancelButtonText "All Sweeps"
        
        Write-Log "Export dialog result: $result"
        
        if ($result -eq $true) {
            # User chose "Selected Only"
            foreach ($item in $selectedItems) {
                if ($item.Tag) {
                    $sweepsToExport += $item.Tag
                }
            }
            # Call the export function with selected sweeps
            Export-SweepConfigurations -SweepsToExport $sweepsToExport -ExportAll $false
        } else {
            # User chose "All Sweeps"
            Export-SweepConfigurations -SweepsToExport @() -ExportAll $true
        }
    } else {
        # No selection, export all without prompting
        Export-SweepConfigurations -SweepsToExport @() -ExportAll $true
    }
})

Update-SplashProgress -PercentComplete 85 -StatusText "Loading Outlook data..."

# Load folders when the application starts
Load-OutlookFolders

# Load categories when the application starts
#Load-OutlookCategories

Update-SplashProgress -PercentComplete 95 -StatusText "Loading saved configurations..."

# Add this function at the end of your script before the window.ShowDialog() call
function Initialize-ComObjectCleanup {
    # Register event handler for application exit
    $window.Add_Closed({
        # Force aggressive garbage collection on exit
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        [System.GC]::Collect()
    })
}

# Call this function just before showing the window

Initialize-ComObjectCleanup

# Call this in your window close event
$window.Add_Closed({
    # Launch external cleanup process
    Clean-OutlookExternalProcess
})

# Load saved sweeps when the application starts
Load-SavedSweeps

Update-SplashProgress -PercentComplete 100 -StatusText "Launching sweep automation..."

# Add window.Add_Loaded handler
$window.Add_Loaded({
    # Clear log before loading categories
    Clear-LogOutput $logOutputTextBox
    Write-Log $logOutputTextBox "Loading categories..."
    
    # Load categories into the dropdown - pass the ComboBox directly
    Load-OutlookCategories -CategoryComboBox $categoryComboBox
    
    # Log message confirming UI is ready
    Write-Log $logOutputTextBox "Email Sweep Tool is ready."
})

# Wait a moment at 100% for user to see completion
Start-Sleep -Milliseconds 1000

# Close the splash screen before showing main window
Close-SplashScreen

# Show the window
$window.ShowDialog() | Out-Null