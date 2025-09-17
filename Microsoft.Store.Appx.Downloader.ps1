# PowerShell function to download UWP package installation files (APPX/MSIX/MSIXBUNDLE/APPXBUNDLE) with dependencies from the Microsoft Store.
# https://woshub.com/how-to-download-appx-installation-file-for-any-windows-store-app/
# https://serverfault.com/questions/1018220/how-do-i-install-an-app-from-windows-store-using-powershell

function Download-AppxPackage {
    [CmdletBinding()]
    param (
        [string]$Uri,
        [string]$Path = ".",
        [string]$Architecture = "Auto"
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
        
        # Get Urls to download
        Write-Information "Getting download links from Microsoft Store..."
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
        Write-Information "Found $totalFiles files to download. Starting downloads..."
        $downloadedFiles = 0
        
        # Download Urls
        foreach ($url in $DownloadLinks) {
            # Get file info
            Write-Information "Getting file info for download $($downloadedFiles + 1) of $totalFiles..."
            try {
                $fileSizeRequest = Invoke-WebRequest -Uri $url -Method Head -UseBasicParsing
                $fileSize = [int]$fileSizeRequest.Headers["Content-Length"]
                $fileName = ($fileSizeRequest.Headers["Content-Disposition"] | Select-String -Pattern  '(?<=filename=).+').matches.value
            } catch {
                Write-Warning "Could not determine file size for $url. Download progress may be inaccurate."
                $fileSize = 0
                $fileName = "unknown.appx"
            }
            
            $FilePath = Join-Path $Path $fileName; $FilePath = Resolve-NameConflict($FilePath)
            
            # Download with progress tracking using Invoke-WebRequest
            Write-Information "Downloading $fileName..."
            try {
                $response = Invoke-WebRequest -Uri $url -OutFile $FilePath -UseBasicParsing
            } catch {
                Write-Error "Error downloading $($fileName): $($_.Exception.Message)"
            }
            
            # Since Invoke-WebRequest doesn't have built-in progress, we'll simulate it
            # For simplicity, we'll just update after each file
            $downloadedFiles++
            
            # Output progress info using Write-Information
            Write-Information ([PSCustomObject]@{
                Type = "OverallProgress"
                TotalFiles = $totalFiles
                DownloadedFiles = $downloadedFiles
                PercentComplete = (($downloadedFiles / $totalFiles) * 100)
                FilePath = $FilePath
                FileName = $fileName
            })
            
            Write-Host "Downloaded: $FilePath"
        }
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
        Title="Download UWP App Packages" Height="418" Width="630" WindowStartupLocation="CenterScreen" ResizeMode="CanResize">
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
        
        <Label Grid.Row="0" Content="Select Common App or Enter Microsoft Store App URL:" FontWeight="Bold" Margin="0,0,0,5"/>
        <ComboBox x:Name="AppComboBox" Grid.Row="1" Margin="0,0,0,10" SelectedIndex="0">
            <ComboBoxItem Content="Custom URL" IsSelected="True"/>
            <ComboBoxItem Content="Microsoft Clock"/>
            <ComboBoxItem Content="Microsoft Paint"/>
            <ComboBoxItem Content="Microsoft Photos"/>
            <ComboBoxItem Content="Snipping Tool"/>
            <ComboBoxItem Content="Windows Media Player"/>
            <ComboBoxItem Content="Windows Notepad"/>
            <ComboBoxItem Content="Windows Calculator"/>
            <ComboBoxItem Content="Windows Terminal"/>
            <ComboBoxItem Content="Microsoft Store"/>
        </ComboBox>
        
        <StackPanel Grid.Row="2" Orientation="Horizontal" Margin="0,0,0,10">
            <TextBox x:Name="UriTextBox" Width="385" Text="Example: https://apps.microsoft.com/detail/..." Margin="0,0,10,0" Foreground="Gray"/>
            <Button x:Name="PasteUrlButton" Content="Paste URL" Width="90" Margin="0,0,10,0"/>
            <TextBlock><Hyperlink x:Name="LookupHyperlink" NavigateUri="https://apps.microsoft.com/home?hl=en-US&amp;gl=US">Lookup App URLs</Hyperlink></TextBlock>
        </StackPanel>
        
        <Label Grid.Row="3" Content="Download Path:" FontWeight="Bold" Margin="0,0,0,5"/>
        <StackPanel Grid.Row="4" Orientation="Horizontal" Margin="0,0,0,10">
            <TextBox x:Name="PathTextBox" Width="385" Text="$ENV:USERPROFILE\Desktop"/>
            <Button x:Name="BrowseButton" Content="Browse" Width="90" Margin="10,0,0,0"/>
        </StackPanel>
        
        <Label Grid.Row="5" Content="Architecture:" FontWeight="Bold" Margin="0,0,0,5"/>
        <ComboBox x:Name="ArchComboBox" Grid.Row="6" Margin="0,0,0,10" SelectedIndex="0">
            <ComboBoxItem Content="Auto" IsSelected="True"/>
            <ComboBoxItem Content="Neutral"/>
            <ComboBoxItem Content="x64"/>
            <ComboBoxItem Content="x86"/>
            <ComboBoxItem Content="ARM"/>
        </ComboBox>
        
        <Label Grid.Row="7" Content="Progress:" FontWeight="Bold" Margin="0,0,0,5"/>
        <TextBlock x:Name="ProgressTextBlock" Grid.Row="8" Margin="0,0,0,5" Text="Ready to download" VerticalAlignment="Top" TextWrapping="Wrap" FontStyle="Italic"/>
        
        <Label Grid.Row="9" Content="Status:" FontWeight="Bold" Margin="0,0,0,5"/>
        <TextBlock x:Name="StatusTextBlock" Grid.Row="10" Margin="0,0,0,5" VerticalAlignment="Center" TextWrapping="Wrap" FontStyle="Italic"/>
        
        <StackPanel Grid.Row="11" Orientation="Horizontal" VerticalAlignment="Bottom" Margin="0,5,0,0">
            <Button x:Name="DownloadButton" Content="Download" Width="100" HorizontalAlignment="Left"/>
            <Button x:Name="InstallButton" Content="Install Packages" Width="120" Margin="10,0,0,0" IsEnabled="False"/>
            <Button x:Name="StopButton" Content="Stop" Width="80" Margin="10,0,0,0" IsEnabled="False"/>
        </StackPanel>
    </Grid>
</Window>
"@

# Load WPF assemblies
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

# Parse XAML
$reader = New-Object System.Xml.XmlNodeReader ([xml]$xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

# Get UI elements
$uriTextBox = $window.FindName("UriTextBox")
$appComboBox = $window.FindName("AppComboBox")
$pathTextBox = $window.FindName("PathTextBox")
$archComboBox = $window.FindName("ArchComboBox")
$browseButton = $window.FindName("BrowseButton")
$downloadButton = $window.FindName("DownloadButton")
$installButton = $window.FindName("InstallButton")
$stopButton = $window.FindName("StopButton")
$statusTextBlock = $window.FindName("StatusTextBlock")
$progressTextBlock = $window.FindName("ProgressTextBlock")
$pasteUrlButton = $window.FindName("PasteUrlButton")
$lookupHyperlink = $window.FindName("LookupHyperlink")

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

# ComboBox selection changed event handler
$appComboBox.Add_SelectionChanged({
    $selectedItem = $appComboBox.SelectedItem.Content
    if ($selectedItem -eq "Custom URL") {
        $uriTextBox.IsEnabled = $true
        $uriTextBox.Text = $hintText
        $uriTextBox.Foreground = "Gray"
    } else {
        $uriTextBox.IsEnabled = $false
        $uriTextBox.Text = $appUrls[$selectedItem]
        $uriTextBox.Foreground = "Black"
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
    $folderBrowser.Description = "Select a folder to save the downloaded files"
    if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $pathTextBox.Text = $folderBrowser.SelectedPath
    }
})

# Hyperlink event handler
$lookupHyperlink.Add_RequestNavigate({
    param($sender, $e)
    Start-Process $e.Uri.AbsoluteUri
})

# Global variables for runspaces
$global:downloadRunspace = $null
$global:installRunspace = $null

# Function to reset UI to default state
function Reset-ToDefault {
    $downloadButton.IsEnabled = $true
    $installButton.IsEnabled = $false
    $stopButton.IsEnabled = $false
    $statusTextBlock.Text = ""
    $progressTextBlock.Text = "Ready to download"
    # Optionally reset other fields if needed, e.g., $pathTextBox.Text = "$ENV:USERPROFILE\Desktop"
}

# Download button event handler
$downloadButton.Add_Click({
    $uri = $uriTextBox.Text
    $path = $pathTextBox.Text
    $arch = $archComboBox.SelectedItem.Content
    
    if (-not [string]::IsNullOrEmpty($uri) -and -not [string]::IsNullOrEmpty($path) -and $uri -ne $hintText) {
        $statusTextBlock.Text = "Initializing download..."
        $downloadButton.IsEnabled = $false
        $installButton.IsEnabled = $false
        $stopButton.IsEnabled = $true
        $progressTextBlock.Text = "Preparing to download"
        
        # Run download in background to avoid blocking UI
        $global:downloadRunspace = [runspacefactory]::CreateRunspace()
        $global:downloadRunspace.Open()
        $powershell = [powershell]::Create()
        $powershell.Runspace = $global:downloadRunspace
        
        $powershell.AddScript({
            param($uri, $path, $arch)
            
            # Define the function inside the runspace
            function Download-AppxPackage {
                [CmdletBinding()]
                param (
                    [string]$Uri,
                    [string]$Path = ".",
                    [string]$Architecture = "Auto"
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
                    
                    # Get Urls to download
                    Write-Information "Getting download links from Microsoft Store..."
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
                    Write-Information "Found $totalFiles files to download. Starting downloads..."
                    $downloadedFiles = 0
                    
                    # Download Urls
                    foreach ($url in $DownloadLinks) {
                        # Get file info
                        Write-Information "Getting file info for download $($downloadedFiles + 1) of $totalFiles..."
                        try {
                            $fileSizeRequest = Invoke-WebRequest -Uri $url -Method Head -UseBasicParsing
                            $fileSize = [int]$fileSizeRequest.Headers["Content-Length"]
                            $fileName = ($fileSizeRequest.Headers["Content-Disposition"] | Select-String -Pattern  '(?<=filename=).+').matches.value
                        } catch {
                            Write-Warning "Could not determine file size for $url. Download progress may be inaccurate."
                            $fileSize = 0
                            $fileName = "unknown.appx"
                        }
                        
                        $FilePath = Join-Path $Path $fileName; $FilePath = Resolve-NameConflict($FilePath)
                        
                        # Download with progress tracking using Invoke-WebRequest
                        Write-Information "Downloading $fileName..."
                        try {
                            $response = Invoke-WebRequest -Uri $url -OutFile $FilePath -UseBasicParsing
                        } catch {
                            Write-Error "Error downloading $($fileName): $($_.Exception.Message)"
                        }
                        
                        # Since Invoke-WebRequest doesn't have built-in progress, we'll simulate it
                        # For simplicity, we'll just update after each file
                        $downloadedFiles++
                        
                        # Output progress info using Write-Information
                        Write-Information ([PSCustomObject]@{
                            Type = "OverallProgress"
                            TotalFiles = $totalFiles
                            DownloadedFiles = $downloadedFiles
                            PercentComplete = (($downloadedFiles / $totalFiles) * 100)
                            FilePath = $FilePath
                            FileName = $fileName
                        })
                        
                        Write-Host "Downloaded: $FilePath"
                    }
                }
            }
            
            try {
                $progressData = Download-AppxPackage -Uri $uri -Path $path -Architecture $arch
                $result = "Download completed successfully!"
            } catch {
                $result = "Error: $($_.Exception.Message)"
            }
            return $result, $progressData
        }).AddArgument($uri).AddArgument($path).AddArgument($arch)
        
        $job = $powershell.BeginInvoke()
        
        # Monitor the job and update progress
        while (-not $job.IsCompleted) {
            Start-Sleep -Milliseconds 500
            $latestProgress = $powershell.Streams.Information[-1]
            if ($latestProgress -and $latestProgress.MessageData -is [string]) {
                $progressTextBlock.Text = $latestProgress.MessageData
            } elseif ($latestProgress -and $latestProgress.MessageData.Type -eq "OverallProgress") {
                $downloaded = $latestProgress.MessageData.DownloadedFiles
                $total = $latestProgress.MessageData.TotalFiles
                $fileName = $latestProgress.MessageData.FileName
                $progressTextBlock.Text = "Downloading $downloaded of $total`nDownloading: $fileName"
            }
            [System.Windows.Forms.Application]::DoEvents()
        }
        
        $results = $powershell.EndInvoke($job)
        $global:downloadRunspace.Close()
        $global:downloadRunspace = $null
        $powershell.Dispose()
        
        $statusTextBlock.Text = $results[0]
        $downloadButton.IsEnabled = $true
        $installButton.IsEnabled = $true
        $stopButton.IsEnabled = $false
        $progressTextBlock.Text = "Download complete"
    } else {
        $statusTextBlock.Text = "Please provide a valid URI and Path."
    }
})

# Install button event handler
$installButton.Add_Click({
    $path = $pathTextBox.Text
    
    if (-not [string]::IsNullOrEmpty($path) -and (Test-Path $path)) {
        $statusTextBlock.Text = "Installing packages..."
        $installButton.IsEnabled = $false
        $downloadButton.IsEnabled = $false
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
        
        $statusTextBlock.Text = $results
        $installButton.IsEnabled = $true
        $downloadButton.IsEnabled = $true
        $progressTextBlock.Text = "Installation complete"
        
        # Reset to default for reuse
        Reset-ToDefault
    } else {
        $statusTextBlock.Text = "Please provide a valid download path containing packages."
    }
})

# Stop button event handler
$stopButton.Add_Click({
    if ($global:downloadRunspace -and $global:downloadRunspace.RunspaceStateInfo.State -eq "Opened") {
        $global:downloadRunspace.Close()
        $global:downloadRunspace = $null
        $statusTextBlock.Text = "Download stopped."
        $downloadButton.IsEnabled = $true
        $installButton.IsEnabled = $false
        $stopButton.IsEnabled = $false
        $progressTextBlock.Text = "Ready to download"
    }
    if ($global:installRunspace -and $global:installRunspace.RunspaceStateInfo.State -eq "Opened") {
        $global:installRunspace.Close()
        $global:installRunspace = $null
        $statusTextBlock.Text = "Installation stopped."
        $installButton.IsEnabled = $true
        $downloadButton.IsEnabled = $true
        $stopButton.IsEnabled = $false
        $progressTextBlock.Text = "Ready"
    }
})

# Show the window
$window.ShowDialog() | Out-Null