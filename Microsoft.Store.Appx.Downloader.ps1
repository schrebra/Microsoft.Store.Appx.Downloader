# PowerShell function to download UWP package installation files (APPX/MSIX/MSIXBUNDLE/APPXBUNDLE) with dependencies from the Microsoft Store.
# https://woshub.com/how-to-download-appx-installation-file-for-any-windows-store-app/
# https://serverfault.com/questions/1018220/how-do-i-install-an-app-from-windows-store-using-powershell

function Download-AppxPackage {
    [CmdletBinding()]
    param (
        [string]$Uri,
        [string]$Path = ".",
        [string]$Architecture = "Auto",
        [string]$AppName = ""
    )
       
    process {
        $Path = (Resolve-Path $Path).Path
        # Determine architecture filter
        if ($Architecture -eq "Auto") {
            $archFilter = "*_"+$env:PROCESSOR_ARCHITECTURE.Replace("AMD","X").Replace("IA","X")+"_*"
        } elseif ($Architecture -eq "Neutral") {
            $archFilter = "*_neutral_*"
        } else {
            $archFilter = "*_"+$Architecture.ToLower()+"_*"
        }
        
        Write-Information "Starting download process for $AppName."
        Write-Information "Using architecture filter: $archFilter"
        # Get Urls to download
        Write-Information "Connecting to Microsoft Store service to retrieve package links for $AppName..."
        $WebResponse = Invoke-WebRequest -UseBasicParsing -Method 'POST' -Uri 'https://store.rg-adguard.net/api/GetFiles' -Body "type=url&url=$Uri&ring=Retail" -ContentType 'application/x-www-form-urlencoded'
        $LinksMatch = $WebResponse.Links | where {$_ -like '*.appx*' -or $_ -like '*.appxbundle*' -or $_ -like '*.msix*' -or $_ -like '*.msixbundle*'} | where {$_ -like $archFilter -or $_ -like '*_neutral_*'} | Select-String -Pattern '(?<=a href=").+(?=" r)'
        $DownloadLinks = $LinksMatch.matches.value 

        function Resolve-NameConflict {
            # Accepts Path to a FILE and changes it so there are no name conflicts
            param(
                [string]$Path
            )
            $newPath = $Path
            if (Test-Path $Path) {
                $i = 0;
                $item = (Get-Item $Path)
                while (Test-Path $newPath) {
                    $i += 1;
                    $newPath = Join-Path $item.DirectoryName ($item.BaseName+"($i)"+$item.Extension)
                }
            }
            return $newPath
        }
        
        $totalFiles = $DownloadLinks.Count
        Write-Information "Found $totalFiles package files available from the store for $AppName. Beginning file existence checks..."
        $downloadedFiles = 0
        
        # First, collect all file names from the store
        $filesToDownload = @()
        foreach ($url in $DownloadLinks) {
            Write-Information "Processing package link: $url"
            # Get file info
            Write-Information "Retrieving metadata for package file from $url..."
            try {
                $fileSizeRequest = Invoke-WebRequest -Uri $url -Method Head -UseBasicParsing
                $fileSize = [int]$fileSizeRequest.Headers["Content-Length"]
                $fileName = ($fileSizeRequest.Headers["Content-Disposition"] | Select-String -Pattern  '(?<=filename=).+').matches.value
                Write-Information "Metadata retrieved: File name: $fileName, Size: $fileSize bytes"
            } catch {
                Write-Warning "Failed to retrieve metadata for $url. Skipping this file."
                continue
            }
            
            $FilePath = Join-Path $Path $fileName
            
            # Check if the original file already exists
            if (Test-Path $FilePath) {
                Write-Information "File $fileName already exists in $Path. Skipping download for this file to avoid redundancy."
                continue
            } else {
                Write-Information "File $fileName does not exist. Preparing to download."
                # Resolve any potential conflicts for future downloads
                $FilePath = Resolve-NameConflict($FilePath)
                $filesToDownload += [PSCustomObject]@{Url = $url; FilePath = $FilePath; FileName = $fileName}
            }
        }
        
        Write-Information "File checks completed. Found $($filesToDownload.Count) files to download for $AppName."
        # Now download only the missing files
        foreach ($file in $filesToDownload) {
            $url = $file.Url
            $FilePath = $file.FilePath
            $fileName = $file.FileName
            
            Write-Information "Initiating download for $fileName from $url to $FilePath..."
            # Download with progress tracking using Invoke-WebRequest
            Write-Information "Downloading file: $fileName (file $downloadedFiles of $($filesToDownload.Count)) for $AppName..."
            try {
                $response = Invoke-WebRequest -Uri $url -OutFile $FilePath -UseBasicParsing
                Write-Information "Download completed for $fileName."
            } catch {
                Write-Error "Error downloading $fileName for $AppName : $($_.Exception.Message)"
            }
            
            # Since Invoke-WebRequest doesn't have built-in progress, we'll simulate it
            # For simplicity, we'll just update after each file
            $downloadedFiles++
            
            # Output progress info using Write-Information
            Write-Information ([PSCustomObject]@{
                Type = "OverallProgress"
                TotalFiles = $totalFiles
                DownloadedFiles = $downloadedFiles
                PercentComplete = (($downloadedFiles / $filesToDownload.Count) * 100)
                FilePath = $FilePath
                FileName = $fileName
                AppName = $AppName
            })
            
            Write-Host "Downloaded: $FilePath"
        }
        Write-Information "Download process for $AppName completed."
    }
}

# Function to install downloaded APPX packages
function Install-AppxPackages {
    param (
        [string]$Path
    )
    
    $Path = (Resolve-Path $Path).Path
    Write-Host "Installing APPX packages from $Path..."
    
    # Install .appx files
    Get-ChildItem $Path -Filter *.appx | ForEach-Object {
        Write-Host "Installing $($_.FullName)..."
        try {
            Add-AppxPackage -Path $_.FullName
            Write-Host "Successfully installed $($_.Name)"
        } catch {
            Write-Error "Failed to install $($_.Name): $($_.Exception.Message)"
        }
    }
    
    # Install .appxbundle files
    Get-ChildItem $Path -Filter *.appxbundle | ForEach-Object {
        Write-Host "Installing $($_.FullName)..."
        try {
            Add-AppxPackage -Path $_.FullName
            Write-Host "Successfully installed $($_.Name)"
        } catch {
            Write-Error "Failed to install $($_.Name): $($_.Exception.Message)"
        }
    }
    
    # Install .msix files
    Get-ChildItem $Path -Filter *.msix | ForEach-Object {
        Write-Host "Installing $($_.FullName)..."
        try {
            Add-AppxPackage -Path $_.FullName
            Write-Host "Successfully installed $($_.Name)"
        } catch {
            Write-Error "Failed to install $($_.Name): $($_.Exception.Message)"
        }
    }
    
    # Install .msixbundle files
    Get-ChildItem $Path -Filter *.msixbundle | ForEach-Object {
        Write-Host "Installing $($_.FullName)..."
        try {
            Add-AppxPackage -Path $_.FullName
            Write-Host "Successfully installed $($_.Name)"
        } catch {
            Write-Error "Failed to install $($_.Name): $($_.Exception.Message)"
        }
    }
    
    Write-Host "Installation process completed."
}

# XAML for the WPF GUI
$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Download UWP App Packages" Height="720" Width="660" WindowStartupLocation="CenterScreen" ResizeMode="CanResize">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
        </Grid.RowDefinitions>
        
        <Label Grid.Row="0" Content="Select Apps to Download:" FontWeight="Bold" Margin="0,0,0,5"/>
        <ListBox x:Name="AppListBox" Grid.Row="1" SelectionMode="Multiple" MaxHeight="300" Margin="0,0,0,10">
            <ListBoxItem Content="Microsoft Clock"/>
            <ListBoxItem Content="Microsoft Paint"/>
            <ListBoxItem Content="Microsoft Photos"/>
            <ListBoxItem Content="Snipping Tool"/>
            <ListBoxItem Content="Windows Media Player"/>
            <ListBoxItem Content="Windows Notepad"/>
            <ListBoxItem Content="Windows Calculator"/>
            <ListBoxItem Content="Windows Terminal"/>
            <ListBoxItem Content="Microsoft Store"/>
            <ListBoxItem Content="Windows Camera"/>
            <ListBoxItem Content="MSN Weather"/>
            <ListBoxItem Content="Xbox"/>
            <ListBoxItem Content="Outlook for Windows"/>
            <ListBoxItem Content="Microsoft Copilot"/>
            <ListBoxItem Content="Microsoft To Do"/>
            <ListBoxItem Content="Microsoft News"/>
            <ListBoxItem Content="Windows Sound Recorder"/>
            <ListBoxItem Content="Microsoft Teams"/>
        </ListBox>
        
        <StackPanel Grid.Row="2" Orientation="Horizontal" Margin="0,0,0,10">
            <TextBox x:Name="UriTextBox" Width="385" Text="Example: https://apps.microsoft.com/detail/..." Margin="0,0,10,0" Foreground="Gray"/>
            <Button x:Name="PasteUrlButton" Content="Paste URL" Width="90" Margin="0,0,10,0"/>
            <TextBlock><Hyperlink x:Name="LookupHyperlink" NavigateUri="https://apps.microsoft.com/home?hl=en-US&amp;gl=US">Lookup App URLs</Hyperlink></TextBlock>
        </StackPanel>
        
        <CheckBox x:Name="SkipDownloadCheckBox" Content="Skip Download, Install from Existing Directory" Grid.Row="3" Margin="0,0,0,10"/>
        
        <Label Grid.Row="4" Content="Path:" FontWeight="Bold" Margin="0,0,0,5"/>
        <StackPanel Grid.Row="5" Orientation="Horizontal" Margin="0,0,0,10">
            <TextBox x:Name="PathTextBox" Width="385" Text="$ENV:USERPROFILE\Documents\appx.packages"/>
            <Button x:Name="BrowseButton" Content="Browse" Width="90" Margin="10,0,0,0"/>
        </StackPanel>
        
        <Label Grid.Row="6" Content="Architecture:" FontWeight="Bold" Margin="0,0,0,5"/>
        <ComboBox x:Name="ArchComboBox" Grid.Row="7" Margin="0,0,0,10" SelectedIndex="0">
            <ComboBoxItem Content="Auto" IsSelected="True"/>
            <ComboBoxItem Content="Neutral"/>
            <ComboBoxItem Content="x64"/>
            <ComboBoxItem Content="x86"/>
            <ComboBoxItem Content="ARM"/>
        </ComboBox>
        
        <Label Grid.Row="8" Content="Log:" FontWeight="Bold" Margin="0,0,0,5"/>
        <ScrollViewer Grid.Row="9" MaxHeight="200" Margin="0,0,0,10">
            <TextBlock x:Name="LogTextBlock" Text="Ready to download" TextWrapping="Wrap" FontSize="12"/>
        </ScrollViewer>
        
        <StackPanel Grid.Row="10" Orientation="Horizontal" VerticalAlignment="Bottom" Margin="0,5,0,0">
            <Button x:Name="DownloadButton" Content="Download Selected" Width="140" HorizontalAlignment="Left" ToolTip="Download the selected apps from the list to the specified path."/>
            <Button x:Name="DownloadAllButton" Content="Download All" Width="120" Margin="10,0,0,0" ToolTip="Download all apps in the list to the specified path."/>
            <Button x:Name="InstallButton" Content="Install Packages" Width="120" Margin="10,0,0,0" IsEnabled="False" ToolTip="Install the packages from the specified path or custom URL."/>
            <Button x:Name="InstallAllButton" Content="Install All" Width="120" Margin="10,0,0,0" IsEnabled="False" ToolTip="Install packages from all subfolders in the base path."/>
            <Button x:Name="StopButton" Content="Stop" Width="80" Margin="10,0,0,0" IsEnabled="False" ToolTip="Stop the current download or installation process."/>
        </StackPanel>
    </Grid>
</Window>
"@

# Load WPF and Windows Forms assemblies
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms

# Parse XAML
$reader = New-Object System.Xml.XmlNodeReader ([xml]$xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

# Get UI elements
$uriTextBox = $window.FindName("UriTextBox")
$appListBox = $window.FindName("AppListBox")
$pathTextBox = $window.FindName("PathTextBox")
$archComboBox = $window.FindName("ArchComboBox")
$browseButton = $window.FindName("BrowseButton")
$downloadButton = $window.FindName("DownloadButton")
$downloadAllButton = $window.FindName("DownloadAllButton")
$installButton = $window.FindName("InstallButton")
$installAllButton = $window.FindName("InstallAllButton")
$stopButton = $window.FindName("StopButton")
$logTextBlock = $window.FindName("LogTextBlock")
$pasteUrlButton = $window.FindName("PasteUrlButton")
$lookupHyperlink = $window.FindName("LookupHyperlink")
$skipDownloadCheckBox = $window.FindName("SkipDownloadCheckBox")

# Define app URLs
$appUrls = @{
    "Microsoft Clock" = "https://apps.microsoft.com/detail/9wzdncrfj3pr?hl=en-US&gl=US"
    "Microsoft Paint" = "https://apps.microsoft.com/detail/9pcfs5b6t72h?hl=en-US&gl=US"
    "Microsoft Photos" = "https://apps.microsoft.com/detail/9wzdncrfjbh4?hl=en-US&gl=US"
    "Snipping Tool" = "https://apps.microsoft.com/detail/9mz95kl8mr0l?hl=en-US&gl=US"
    "Windows Media Player" = "https://apps.microsoft.com/detail/9wzdncrfj3pt?hl=en-US&gl=US"
    "Windows Notepad" = "https://apps.microsoft.com/detail/9msmlrh6lzf3?hl=en-US&gl=US"
    "Windows Calculator" = "https://apps.microsoft.com/detail/9wzdncrfhvn5?hl=en-US&gl=US"
    "Windows Terminal" = "https://apps.microsoft.com/detail/9n0dx20hk701?hl=en-US&gl=US"
    "Microsoft Store" = "https://apps.microsoft.com/detail/9wzdncrfjbmp?hl=en-US&gl=US"
    "Windows Camera" = "https://apps.microsoft.com/detail/9wzdncrfjbbg?hl=en-US&gl=US"
    "MSN Weather" = "https://apps.microsoft.com/detail/9wzdncrfj3q2?hl=en-US&gl=US"
    "Xbox" = "https://apps.microsoft.com/detail/9mv0b5hzvk9z?hl=en-US&gl=US"
    "Outlook for Windows" = "https://apps.microsoft.com/detail/9nrx63209r7b?hl=en-US&gl=US"
    "Microsoft Copilot" = "https://apps.microsoft.com/detail/9nht9rb2f4hd?hl=en-US&gl=US"
    "Microsoft To Do" = "https://apps.microsoft.com/detail/9nblggh5r558?hl=en-US&gl=US"
    "Microsoft News" = "https://apps.microsoft.com/detail/9wzdncrfhvfw?hl=en-US&gl=US"
    "Windows Sound Recorder" = "https://apps.microsoft.com/detail/9wzdncrfhwkn?hl=en-US&gl=US"
    "Microsoft Teams" = "https://apps.microsoft.com/detail/xp8bt8dw290mpq?hl=en-US&gl=US"
}

# Placeholder text handling for UriTextBox
$hintText = "Example: https://apps.microsoft.com/detail/..."
$uriTextBox.Add_GotFocus({
    if ($uriTextBox.Text -eq $hintText) {
        $uriTextBox.Text = ""
        $uriTextBox.Foreground = "Black"
    }
})
$uriTextBox.Add_LostFocus({
    if ([string]::IsNullOrWhiteSpace($uriTextBox.Text)) {
        $uriTextBox.Text = $hintText
        $uriTextBox.Foreground = "Gray"
    }
})

# Paste URL button event handler
$pasteUrlButton.Add_Click({
    $clipboardText = Get-Clipboard
    if ($clipboardText -and $clipboardText -match "https://apps\.microsoft\.com/detail/") {
        $uriTextBox.Text = $clipboardText
        $uriTextBox.Foreground = "Black"
    } else {
        [System.Windows.MessageBox]::Show("Clipboard does not contain a valid Microsoft Store app URL.", "Invalid URL", "OK", "Error")
    }
})

# Browse button event handler
$browseButton.Add_Click({
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = "Select a folder to save the downloaded files or containing packages"
    if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $pathTextBox.Text = $folderBrowser.SelectedPath
    }
})

# Hyperlink event handler
$lookupHyperlink.Add_RequestNavigate({
    param($sender, $e)
    Start-Process $e.Uri.AbsoluteUri
})

# Checkbox event handler
$skipDownloadCheckBox.Add_Checked({
    $downloadButton.IsEnabled = $false
    $downloadAllButton.IsEnabled = $false
    $installButton.IsEnabled = $true
    $installAllButton.IsEnabled = $true
    $logTextBlock.Text = "Ready to install"
})
$skipDownloadCheckBox.Add_Unchecked({
    $downloadButton.IsEnabled = $true
    $downloadAllButton.IsEnabled = $true
    $installButton.IsEnabled = $false
    $installAllButton.IsEnabled = $false
    $logTextBlock.Text = "Ready to download"
})

# Global variables for runspaces
$global:downloadRunspace = $null
$global:installRunspace = $null

# Function to reset UI to default state
function Reset-ToDefault {
    $downloadButton.IsEnabled = -not $skipDownloadCheckBox.IsChecked
    $downloadAllButton.IsEnabled = -not $skipDownloadCheckBox.IsChecked
    $installButton.IsEnabled = $skipDownloadCheckBox.IsChecked
    $installAllButton.IsEnabled = $skipDownloadCheckBox.IsChecked
    $stopButton.IsEnabled = $false
    $logTextBlock.Text = if ($skipDownloadCheckBox.IsChecked) { "Ready to install" } else { "Ready to download" }
}

# Download button event handler
$downloadButton.Add_Click({
    $basePath = $pathTextBox.Text
    $arch = $archComboBox.SelectedItem.Content
    $uris = @()
    
    if (-not [string]::IsNullOrEmpty($basePath)) {
        $logTextBlock.Text = "Initializing download process..."
        $downloadButton.IsEnabled = $false
        $downloadAllButton.IsEnabled = $false
        $installButton.IsEnabled = $false
        $installAllButton.IsEnabled = $false
        $stopButton.IsEnabled = $true
        
        # Ensure base folder exists
        if (-not (Test-Path $basePath)) {
            New-Item -ItemType Directory -Path $basePath | Out-Null
        }
        
        # Check if custom URL is provided
        if ($uriTextBox.Text -ne $hintText -and -not [string]::IsNullOrWhiteSpace($uriTextBox.Text)) {
            $uris += [PSCustomObject]@{Uri = $uriTextBox.Text; Path = Join-Path $basePath "Custom"; Name = "Custom"}
        } else {
            # Collect selected apps
            $selectedApps = @()
            foreach ($item in $appListBox.SelectedItems) {
                $selectedApps += $item.Content
            }
            
            foreach ($app in $selectedApps) {
                $uris += [PSCustomObject]@{Uri = $appUrls[$app]; Path = Join-Path $basePath $app; Name = $app}
            }
        }
        
        if ($uris.Count -eq 0) {
            $logTextBlock.Text = "No apps selected or custom URL provided. Please select apps or enter a valid URL."
            $downloadButton.IsEnabled = $true
            $downloadAllButton.IsEnabled = $true
            $installButton.IsEnabled = $true
            $installAllButton.IsEnabled = $true
            $stopButton.IsEnabled = $false
            return
        }
        
        # Run download in background to avoid blocking UI
        $global:downloadRunspace = [runspacefactory]::CreateRunspace()
        $global:downloadRunspace.Open()
        $powershell = [powershell]::Create()
        $powershell.Runspace = $global:downloadRunspace
        
        $powershell.AddScript({
            param($uris, $arch)
            
            # Define the function inside the runspace
            function Download-AppxPackage {
                [CmdletBinding()]
                param (
                    [string]$Uri,
                    [string]$Path = ".",
                    [string]$Architecture = "Auto",
                    [string]$AppName = ""
                )
                   
                process {
                    $Path = (Resolve-Path $Path).Path
                    # Determine architecture filter
                    if ($Architecture -eq "Auto") {
                        $archFilter = "*_"+$env:PROCESSOR_ARCHITECTURE.Replace("AMD","X").Replace("IA","X")+"_*"
                    } elseif ($Architecture -eq "Neutral") {
                        $archFilter = "*_neutral_*"
                    } else {
                        $archFilter = "*_"+$Architecture.ToLower()+"_*"
                    }
                    
                    Write-Information "Starting download process for $AppName."
                    Write-Information "Using architecture filter: $archFilter"
                    # Get Urls to download
                    Write-Information "Connecting to Microsoft Store service to retrieve package links for $AppName..."
                    $WebResponse = Invoke-WebRequest -UseBasicParsing -Method 'POST' -Uri 'https://store.rg-adguard.net/api/GetFiles' -Body "type=url&url=$Uri&ring=Retail" -ContentType 'application/x-www-form-urlencoded'
                    $LinksMatch = $WebResponse.Links | where {$_ -like '*.appx*' -or $_ -like '*.appxbundle*' -or $_ -like '*.msix*' -or $_ -like '*.msixbundle*'} | where {$_ -like $archFilter -or $_ -like '*_neutral_*'} | Select-String -Pattern '(?<=a href=").+(?=" r)'
                    $DownloadLinks = $LinksMatch.matches.value 

                    function Resolve-NameConflict {
                        # Accepts Path to a FILE and changes it so there are no name conflicts
                        param(
                            [string]$Path
                        )
                        $newPath = $Path
                        if (Test-Path $Path) {
                            $i = 0;
                            $item = (Get-Item $Path)
                            while (Test-Path $newPath) {
                                $i += 1;
                                $newPath = Join-Path $item.DirectoryName ($item.BaseName+"($i)"+$item.Extension)
                            }
                        }
                        return $newPath
                    }
                    
                    $totalFiles = $DownloadLinks.Count
                    Write-Information "Found $totalFiles package files available from the store for $AppName. Beginning file existence checks..."
                    $downloadedFiles = 0
                    
                    # First, collect all file names from the store
                    $filesToDownload = @()
                    foreach ($url in $DownloadLinks) {
                        Write-Information "Processing package link: $url"
                        # Get file info
                        Write-Information "Retrieving metadata for package file from $url..."
                        try {
                            $fileSizeRequest = Invoke-WebRequest -Uri $url -Method Head -UseBasicParsing
                            $fileSize = [int]$fileSizeRequest.Headers["Content-Length"]
                            $fileName = ($fileSizeRequest.Headers["Content-Disposition"] | Select-String -Pattern  '(?<=filename=).+').matches.value
                            Write-Information "Metadata retrieved: File name: $fileName, Size: $fileSize bytes"
                        } catch {
                            Write-Warning "Failed to retrieve metadata for $url. Skipping this file."
                            continue
                        }
                        
                        $FilePath = Join-Path $Path $fileName
                        
                        # Check if the original file already exists
                        if (Test-Path $FilePath) {
                            Write-Information "File $fileName already exists in $Path. Skipping download for this file to avoid redundancy."
                            continue
                        } else {
                            Write-Information "File $fileName does not exist. Preparing to download."
                            # Resolve any potential conflicts for future downloads
                            $FilePath = Resolve-NameConflict($FilePath)
                            $filesToDownload += [PSCustomObject]@{Url = $url; FilePath = $FilePath; FileName = $fileName}
                        }
                    }
                    
                    Write-Information "File checks completed. Found $($filesToDownload.Count) files to download for $AppName."
                    # Now download only the missing files
                    foreach ($file in $filesToDownload) {
                        $url = $file.Url
                        $FilePath = $file.FilePath
                        $fileName = $file.FileName
                        
                        Write-Information "Initiating download for $fileName from $url to $FilePath..."
                        # Download with progress tracking using Invoke-WebRequest
                        Write-Information "Downloading file: $fileName (file $downloadedFiles of $($filesToDownload.Count)) for $AppName..."
                        try {
                            $response = Invoke-WebRequest -Uri $url -OutFile $FilePath -UseBasicParsing
                            Write-Information "Download completed for $fileName."
                        } catch {
                            Write-Error "Error downloading $fileName for $AppName : $($_.Exception.Message)"
                        }
                        
                        # Since Invoke-WebRequest doesn't have built-in progress, we'll simulate it
                        # For simplicity, we'll just update after each file
                        $downloadedFiles++
                        
                        # Output progress info using Write-Information
                        Write-Information ([PSCustomObject]@{
                            Type = "OverallProgress"
                            TotalFiles = $totalFiles
                            DownloadedFiles = $downloadedFiles
                            PercentComplete = (($downloadedFiles / $filesToDownload.Count) * 100)
                            FilePath = $FilePath
                            FileName = $fileName
                            AppName = $AppName
                        })
                        
                        Write-Host "Downloaded: $FilePath"
                    }
                    Write-Information "Download process for $AppName completed."
                }
            }
            
            $totalApps = $uris.Count
            $currentApp = 0
            $downloadsPerformed = $false
            foreach ($item in $uris) {
                $currentApp++
                $uri = $item.Uri
                $path = $item.Path
                $appName = $item.Name
                if (-not (Test-Path $path)) {
                    New-Item -ItemType Directory -Path $path | Out-Null
                }
                Write-Information "Downloading $appName ($currentApp of $totalApps)"
                try {
                    Download-AppxPackage -Uri $uri -Path $path -Architecture $arch -AppName $appName
                    $downloadsPerformed = $true
                } catch {
                    Write-Error "Error downloading $appName : $($_.Exception.Message)"
                }
            }
            if ($downloadsPerformed) {
                $result = "Download completed successfully!"
            } else {
                $result = "No download needed, latest packages already downloaded."
            }
            return $result
        }).AddArgument($uris).AddArgument($arch)
        
        $job = $powershell.BeginInvoke()
        
        # Monitor the job and update progress
        while (-not $job.IsCompleted) {
            Start-Sleep -Milliseconds 500
            $latestProgress = $powershell.Streams.Information[-1]
            if ($latestProgress) {
                $logTextBlock.Text = $latestProgress.MessageData
            }
            [System.Windows.Forms.Application]::DoEvents()
        }
        
        $results = $powershell.EndInvoke($job)
        $global:downloadRunspace.Close()
        $global:downloadRunspace = $null
        $powershell.Dispose()
        
        $logTextBlock.Text = $results
        $downloadButton.IsEnabled = $true
        $downloadAllButton.IsEnabled = $true
        $installButton.IsEnabled = $true
        $installAllButton.IsEnabled = $true
        $stopButton.IsEnabled = $false
    } else {
        $logTextBlock.Text = "Please provide a valid Path."
    }
})

# Download All button event handler
$downloadAllButton.Add_Click({
    $basePath = $pathTextBox.Text
    $arch = $archComboBox.SelectedItem.Content
    
    if (-not [string]::IsNullOrEmpty($basePath)) {
        $logTextBlock.Text = "Initializing download for all apps..."
        $downloadButton.IsEnabled = $false
        $downloadAllButton.IsEnabled = $false
        $installButton.IsEnabled = $false
        $installAllButton.IsEnabled = $false
        $stopButton.IsEnabled = $true
        
        # Ensure base folder exists
        if (-not (Test-Path $basePath)) {
            New-Item -ItemType Directory -Path $basePath | Out-Null
        }
        
        $uris = $appUrls.Keys | ForEach-Object { [PSCustomObject]@{Uri = $appUrls[$_]; Path = Join-Path $basePath $_; Name = $_} }
        
        # Run download in background to avoid blocking UI
        $global:downloadRunspace = [runspacefactory]::CreateRunspace()
        $global:downloadRunspace.Open()
        $powershell = [powershell]::Create()
        $powershell.Runspace = $global:downloadRunspace
        
        $powershell.AddScript({
            param($uris, $arch)
            
            # Define the function inside the runspace
            function Download-AppxPackage {
                [CmdletBinding()]
                param (
                    [string]$Uri,
                    [string]$Path = ".",
                    [string]$Architecture = "Auto",
                    [string]$AppName = ""
                )
                   
                process {
                    $Path = (Resolve-Path $Path).Path
                    # Determine architecture filter
                    if ($Architecture -eq "Auto") {
                        $archFilter = "*_"+$env:PROCESSOR_ARCHITECTURE.Replace("AMD","X").Replace("IA","X")+"_*"
                    } elseif ($Architecture -eq "Neutral") {
                        $archFilter = "*_neutral_*"
                    } else {
                        $archFilter = "*_"+$Architecture.ToLower()+"_*"
                    }
                    
                    Write-Information "Starting download process for $AppName."
                    Write-Information "Using architecture filter: $archFilter"
                    # Get Urls to download
                    Write-Information "Connecting to Microsoft Store service to retrieve package links for $AppName..."
                    $WebResponse = Invoke-WebRequest -UseBasicParsing -Method 'POST' -Uri 'https://store.rg-adguard.net/api/GetFiles' -Body "type=url&url=$Uri&ring=Retail" -ContentType 'application/x-www-form-urlencoded'
                    $LinksMatch = $WebResponse.Links | where {$_ -like '*.appx*' -or $_ -like '*.appxbundle*' -or $_ -like '*.msix*' -or $_ -like '*.msixbundle*'} | where {$_ -like $archFilter -or $_ -like '*_neutral_*'} | Select-String -Pattern '(?<=a href=").+(?=" r)'
                    $DownloadLinks = $LinksMatch.matches.value 

                    function Resolve-NameConflict {
                        # Accepts Path to a FILE and changes it so there are no name conflicts
                        param(
                            [string]$Path
                        )
                        $newPath = $Path
                        if (Test-Path $Path) {
                            $i = 0;
                            $item = (Get-Item $Path)
                            while (Test-Path $newPath) {
                                $i += 1;
                                $newPath = Join-Path $item.DirectoryName ($item.BaseName+"($i)"+$item.Extension)
                            }
                        }
                        return $newPath
                    }
                    
                    $totalFiles = $DownloadLinks.Count
                    Write-Information "Found $totalFiles package files available from the store for $AppName. Beginning file existence checks..."
                    $downloadedFiles = 0
                    
                    # First, collect all file names from the store
                    $filesToDownload = @()
                    foreach ($url in $DownloadLinks) {
                        Write-Information "Processing package link: $url"
                        # Get file info
                        Write-Information "Retrieving metadata for package file from $url..."
                        try {
                            $fileSizeRequest = Invoke-WebRequest -Uri $url -Method Head -UseBasicParsing
                            $fileSize = [int]$fileSizeRequest.Headers["Content-Length"]
                            $fileName = ($fileSizeRequest.Headers["Content-Disposition"] | Select-String -Pattern  '(?<=filename=).+').matches.value
                            Write-Information "Metadata retrieved: File name: $fileName, Size: $fileSize bytes"
                        } catch {
                            Write-Warning "Failed to retrieve metadata for $url. Skipping this file."
                            continue
                        }
                        
                        $FilePath = Join-Path $Path $fileName
                        
                        # Check if the original file already exists
                        if (Test-Path $FilePath) {
                            Write-Information "File $fileName already exists in $Path. Skipping download for this file to avoid redundancy."
                            continue
                        } else {
                            Write-Information "File $fileName does not exist. Preparing to download."
                            # Resolve any potential conflicts for future downloads
                            $FilePath = Resolve-NameConflict($FilePath)
                            $filesToDownload += [PSCustomObject]@{Url = $url; FilePath = $FilePath; FileName = $fileName}
                        }
                    }
                    
                    Write-Information "File checks completed. Found $($filesToDownload.Count) files to download for $AppName."
                    # Now download only the missing files
                    foreach ($file in $filesToDownload) {
                        $url = $file.Url
                        $FilePath = $file.FilePath
                        $fileName = $file.FileName
                        
                        Write-Information "Initiating download for $fileName from $url to $FilePath..."
                        # Download with progress tracking using Invoke-WebRequest
                        Write-Information "Downloading file: $fileName (file $downloadedFiles of $($filesToDownload.Count)) for $AppName..."
                        try {
                            $response = Invoke-WebRequest -Uri $url -OutFile $FilePath -UseBasicParsing
                            Write-Information "Download completed for $fileName."
                        } catch {
                            Write-Error "Error downloading $fileName for $AppName : $($_.Exception.Message)"
                        }
                        
                        # Since Invoke-WebRequest doesn't have built-in progress, we'll simulate it
                        # For simplicity, we'll just update after each file
                        $downloadedFiles++
                        
                        # Output progress info using Write-Information
                        Write-Information ([PSCustomObject]@{
                            Type = "OverallProgress"
                            TotalFiles = $totalFiles
                            DownloadedFiles = $downloadedFiles
                            PercentComplete = (($downloadedFiles / $filesToDownload.Count) * 100)
                            FilePath = $FilePath
                            FileName = $fileName
                            AppName = $AppName
                        })
                        
                        Write-Host "Downloaded: $FilePath"
                    }
                    Write-Information "Download process for $AppName completed."
                }
            }
            
            $totalApps = $uris.Count
            $currentApp = 0
            $downloadsPerformed = $false
            foreach ($item in $uris) {
                $currentApp++
                $uri = $item.Uri
                $path = $item.Path
                $appName = $item.Name
                if (-not (Test-Path $path)) {
                    New-Item -ItemType Directory -Path $path | Out-Null
                }
                Write-Information "Downloading $appName ($currentApp of $totalApps)"
                try {
                    Download-AppxPackage -Uri $uri -Path $path -Architecture $arch -AppName $appName
                    $downloadsPerformed = $true
                } catch {
                    Write-Error "Error downloading $appName : $($_.Exception.Message)"
                }
            }
            if ($downloadsPerformed) {
                $result = "All downloads completed successfully!"
            } else {
                $result = "No download needed, latest packages already downloaded."
            }
            return $result
        }).AddArgument($uris).AddArgument($arch)
        
        $job = $powershell.BeginInvoke()
        
        # Monitor the job and update progress
        while (-not $job.IsCompleted) {
            Start-Sleep -Milliseconds 500
            $latestProgress = $powershell.Streams.Information[-1]
            if ($latestProgress) {
                $logTextBlock.Text = $latestProgress.MessageData
            }
            [System.Windows.Forms.Application]::DoEvents()
        }
        
        $results = $powershell.EndInvoke($job)
        $global:downloadRunspace.Close()
        $global:downloadRunspace = $null
        $powershell.Dispose()
        
        $logTextBlock.Text = $results
        $downloadButton.IsEnabled = $true
        $downloadAllButton.IsEnabled = $true
        $installButton.IsEnabled = $true
        $installAllButton.IsEnabled = $true
        $stopButton.IsEnabled = $false
    } else {
        $logTextBlock.Text = "Please provide a valid Path."
    }
})

# Install button event handler
$installButton.Add_Click({
    $basePath = $pathTextBox.Text
    
    # Determine the path to install from
    if ($uriTextBox.Text -ne $hintText -and -not [string]::IsNullOrWhiteSpace($uriTextBox.Text)) {
        $path = Join-Path $basePath "Custom"
    } else {
        $path = $basePath
    }
    
    if (-not [string]::IsNullOrEmpty($path) -and (Test-Path $path)) {
        $logTextBlock.Text = "Starting installation process..."
        $installButton.IsEnabled = $false
        $installAllButton.IsEnabled = $false
        $downloadButton.IsEnabled = $false
        $downloadAllButton.IsEnabled = $false
        $progressTextBlock.Text = "Installing packages from $path"
        
        # Run install in background to avoid blocking UI
        $global:installRunspace = [runspacefactory]::CreateRunspace()
        $global:installRunspace.Open()
        $powershell = [powershell]::Create()
        $powershell.Runspace = $global:installRunspace
        
        $powershell.AddScript({
            param($path)
            
            # Define the function inside the runspace
            function Install-AppxPackages {
                param (
                    [string]$Path
                )
                
                $Path = (Resolve-Path $Path).Path
                Write-Host "Installing APPX packages from $Path..."
                
                # Install .appx files
                Get-ChildItem $Path -Filter *.appx | ForEach-Object {
                    Write-Host "Installing $($_.FullName)..."
                    try {
                        Add-AppxPackage -Path $_.FullName
                        Write-Host "Successfully installed $($_.Name)"
                    } catch {
                        Write-Error "Failed to install $($_.Name): $($_.Exception.Message)"
                    }
                }
                
                # Install .appxbundle files
                Get-ChildItem $Path -Filter *.appxbundle | ForEach-Object {
                    Write-Host "Installing $($_.FullName)..."
                    try {
                        Add-AppxPackage -Path $_.FullName
                        Write-Host "Successfully installed $($_.Name)"
                    } catch {
                        Write-Error "Failed to install $($_.Name): $($_.Exception.Message)"
                    }
                }
                
                # Install .msix files
                Get-ChildItem $Path -Filter *.msix | ForEach-Object {
                    Write-Host "Installing $($_.FullName)..."
                    try {
                        Add-AppxPackage -Path $_.FullName
                        Write-Host "Successfully installed $($_.Name)"
                    } catch {
                        Write-Error "Failed to install $($_.Name): $($_.Exception.Message)"
                    }
                }
                
                # Install .msixbundle files
                Get-ChildItem $Path -Filter *.msixbundle | ForEach-Object {
                    Write-Host "Installing $($_.FullName)..."
                    try {
                        Add-AppxPackage -Path $_.FullName
                        Write-Host "Successfully installed $($_.Name)"
                    } catch {
                        Write-Error "Failed to install $($_.Name): $($_.Exception.Message)"
                    }
                }
                
                Write-Host "Installation process completed."
            }
            
            try {
                Install-AppxPackages -Path $path
                $result = "Installation completed successfully!"
            } catch {
                $result = "Error: $($_.Exception.Message)"
            }
            return $result
        }).AddArgument($path)
        
        $job = $powershell.BeginInvoke()
        
        # Monitor the job and update progress
        while (-not $job.IsCompleted) {
            Start-Sleep -Milliseconds 500
            [System.Windows.Forms.Application]::DoEvents()
        }
        
        $results = $powershell.EndInvoke($job)
        $global:installRunspace.Close()
        $global:installRunspace = $null
        $powershell.Dispose()
        
        $logTextBlock.Text = $results
        $installButton.IsEnabled = $true
        $installAllButton.IsEnabled = $true
        $downloadButton.IsEnabled = -not $skipDownloadCheckBox.IsChecked
        $downloadAllButton.IsEnabled = -not $skipDownloadCheckBox.IsChecked
        
        # Reset to default for reuse
        Reset-ToDefault
    } else {
        $logTextBlock.Text = "Please provide a valid path containing packages."
    }
})

# Install All button event handler
$installAllButton.Add_Click({
    $basePath = $pathTextBox.Text
    
    if (-not [string]::IsNullOrEmpty($basePath) -and (Test-Path $basePath)) {
        $logTextBlock.Text = "Starting installation for all apps..."
        $installButton.IsEnabled = $false
        $installAllButton.IsEnabled = $false
        $downloadButton.IsEnabled = $false
        $downloadAllButton.IsEnabled = $false
        
        # Run install in background to avoid blocking UI
        $global:installRunspace = [runspacefactory]::CreateRunspace()
        $global:installRunspace.Open()
        $powershell = [powershell]::Create()
        $powershell.Runspace = $global:installRunspace
        
        $powershell.AddScript({
            param($basePath)
            
            # Define the function inside the runspace
            function Install-AppxPackages {
                param (
                    [string]$Path
                )
                
                $Path = (Resolve-Path $Path).Path
                Write-Host "Installing APPX packages from $Path..."
                
                # Install .appx files
                Get-ChildItem $Path -Filter *.appx | ForEach-Object {
                    Write-Host "Installing $($_.FullName)..."
                    try {
                        Add-AppxPackage -Path $_.FullName
                        Write-Host "Successfully installed $($_.Name)"
                    } catch {
                        Write-Error "Failed to install $($_.Name): $($_.Exception.Message)"
                    }
                }
                
                # Install .appxbundle files
                Get-ChildItem $Path -Filter *.appxbundle | ForEach-Object {
                    Write-Host "Installing $($_.FullName)..."
                    try {
                        Add-AppxPackage -Path $_.FullName
                        Write-Host "Successfully installed $($_.Name)"
                    } catch {
                        Write-Error "Failed to install $($_.Name): $($_.Exception.Message)"
                    }
                }
                
                # Install .msix files
                Get-ChildItem $Path -Filter *.msix | ForEach-Object {
                    Write-Host "Installing $($_.FullName)..."
                    try {
                        Add-AppxPackage -Path $_.FullName
                        Write-Host "Successfully installed $($_.Name)"
                    } catch {
                        Write-Error "Failed to install $($_.Name): $($_.Exception.Message)"
                    }
                }
                
                # Install .msixbundle files
                Get-ChildItem $Path -Filter *.msixbundle | ForEach-Object {
                    Write-Host "Installing $($_.FullName)..."
                    try {
                        Add-AppxPackage -Path $_.FullName
                        Write-Host "Successfully installed $($_.Name)"
                    } catch {
                        Write-Error "Failed to install $($_.Name): $($_.Exception.Message)"
                    }
                }
                
                Write-Host "Installation process completed."
            }
            
            $subfolders = Get-ChildItem -Directory $basePath
            $totalSubfolders = $subfolders.Count
            $currentSubfolder = 0
            foreach ($folder in $subfolders) {
                $currentSubfolder++
                Write-Information "Installing from $($folder.Name) ($currentSubfolder of $totalSubfolders)"
                try {
                    Install-AppxPackages -Path $folder.FullName
                } catch {
                    Write-Error "Error installing from $($folder.Name): $($_.Exception.Message)"
                }
            }
            $result = "All installations completed successfully!"
            return $result
        }).AddArgument($basePath)
        
        $job = $powershell.BeginInvoke()
        
        # Monitor the job and update progress
        while (-not $job.IsCompleted) {
            Start-Sleep -Milliseconds 500
            $latestProgress = $powershell.Streams.Information[-1]
            if ($latestProgress) {
                $logTextBlock.Text = $latestProgress.MessageData
            }
            [System.Windows.Forms.Application]::DoEvents()
        }
        
        $results = $powershell.EndInvoke($job)
        $global:installRunspace.Close()
        $global:installRunspace = $null
        $powershell.Dispose()
        
        $logTextBlock.Text = $results
        $installButton.IsEnabled = $true
        $installAllButton.IsEnabled = $true
        $downloadButton.IsEnabled = -not $skipDownloadCheckBox.IsChecked
        $downloadAllButton.IsEnabled = -not $skipDownloadCheckBox.IsChecked
        
        # Reset to default for reuse
        Reset-ToDefault
    } else {
        $logTextBlock.Text = "Please provide a valid base path containing subfolders with packages."
    }
})

# Stop button event handler
$stopButton.Add_Click({
    if ($global:downloadRunspace -and $global:downloadRunspace.RunspaceStateInfo.State -eq "Opened") {
        $global:downloadRunspace.Close()
        $global:downloadRunspace = $null
        $logTextBlock.Text = "Download stopped."
        $downloadButton.IsEnabled = $true
        $downloadAllButton.IsEnabled = $true
        $installButton.IsEnabled = $false
        $installAllButton.IsEnabled = $false
        $stopButton.IsEnabled = $false
    }
    if ($global:installRunspace -and $global:installRunspace.RunspaceStateInfo.State -eq "Opened") {
        $global:installRunspace.Close()
        $global:installRunspace = $null
        $logTextBlock.Text = "Installation stopped."
        $installButton.IsEnabled = $true
        $installAllButton.IsEnabled = $true
        $downloadButton.IsEnabled = -not $skipDownloadCheckBox.IsChecked
        $downloadAllButton.IsEnabled = -not $skipDownloadCheckBox.IsChecked
        $stopButton.IsEnabled = $false
    }
})

# Show the window
$window.ShowDialog() | Out-Null
