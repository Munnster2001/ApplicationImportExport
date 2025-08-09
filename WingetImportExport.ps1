#requires -Version 5.0
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms

# WPF XAML for the GUI
$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Application Import/Export" Height="600" Width="800" WindowStartupLocation="CenterScreen">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
        </Grid.RowDefinitions>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        
        <Label Grid.Row="0" Grid.Column="0" Content="File Path:" VerticalAlignment="Center" Margin="5"/>
        <TextBox x:Name="TextBoxFile" Grid.Row="0" Grid.Column="1" VerticalAlignment="Center" Margin="5"/>
        <Button x:Name="ButtonBrowse" Grid.Row="0" Grid.Column="2" Content="Browse" Margin="5" Width="80"/>
        
        <Grid Grid.Row="1" Grid.Column="1" Grid.ColumnSpan="2" Margin="5">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <Button x:Name="ButtonExport" Grid.Column="0" Content="Export Applications" Margin="0,0,5,0" Width="180" HorizontalAlignment="Left"/>
            <Button x:Name="ButtonImport" Grid.Column="1" Content="Import Applications" Margin="5,0,0,0" Width="180" HorizontalAlignment="Left"/>
        </Grid>
        
        <ProgressBar x:Name="ProgressBarStatus" Grid.Row="2" Grid.Column="0" Grid.ColumnSpan="3" Height="20" Margin="5" Visibility="Hidden"/>
        
        <Label Grid.Row="3" Grid.Column="0" Content="Verbose Output:" VerticalAlignment="Top" Margin="5"/>
        <ScrollViewer Grid.Row="3" Grid.Column="1" Grid.ColumnSpan="2" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto">
            <TextBox x:Name="TextBoxVerbose" FontFamily="Consolas" FontSize="12" IsReadOnly="True" AcceptsReturn="True" TextWrapping="Wrap" Margin="5"/>
        </ScrollViewer>
    </Grid>
</Window>
"@

# Load the XAML
$xmlStream = [System.IO.MemoryStream]::new()
$writer = [System.IO.StreamWriter]::new($xmlStream)
$writer.Write($xaml)
$writer.Flush()
$xmlStream.Seek(0, 'Begin')
$window = [System.Windows.Markup.XamlReader]::Load($xmlStream)
$xmlStream.Close()

# Get controls by name
$textBoxFile = $window.FindName("TextBoxFile")
$buttonBrowse = $window.FindName("ButtonBrowse")
$buttonExport = $window.FindName("ButtonExport")
$buttonImport = $window.FindName("ButtonImport")
$progressBarStatus = $window.FindName("ProgressBarStatus")
$textBoxVerbose = $window.FindName("TextBoxVerbose")

# Set default path with correct date format
$currentDate = Get-Date -Format "yyddMM"
$defaultFileName = "MyApps_$currentDate.json"
$defaultPath = "C:\temp\$defaultFileName"
$textBoxFile.Text = $defaultPath

# Event Handlers
$buttonBrowse.Add_Click({
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "JSON files (*.json)|*.json"
    $saveFileDialog.InitialDirectory = "C:\temp"
    $saveFileDialog.FileName = $defaultFileName
    if ($saveFileDialog.ShowDialog() -eq "OK") {
        $textBoxFile.Text = $saveFileDialog.FileName
    }
})

$buttonExport.Add_Click({
    $progressBarStatus.Visibility = "Visible"
    $progressBarStatus.Value = 0
    $textBoxVerbose.Text = ""
    $filePath = $textBoxFile.Text
    
    # Ensure the directory exists
    $dirPath = Split-Path -Parent $filePath
    if (-not (Test-Path $dirPath)) {
        New-Item -ItemType Directory -Path $dirPath -Force | Out-Null
    }

    $progressBarStatus.Value = 20
    $textBoxVerbose.Text += "Starting winget list to gather application data..." + "`n"
    [System.Windows.Forms.Application]::DoEvents()
    
    # Run winget list and capture the output
    $wingetOutput = & winget list 2>&1
    
    # Check if the command was successful
    if ($LASTEXITCODE -ne 0) {
        $textBoxVerbose.Text += "Error: winget list command failed." + "`n"
        $textBoxVerbose.Text += $wingetOutput + "`n"
        [System.Windows.MessageBox]::Show("Error: winget list command failed. Please check your winget installation.", "Error")
        $progressBarStatus.Visibility = "Hidden"
        return
    }

    $progressBarStatus.Value = 40
    $textBoxVerbose.Text += "Parsing and sorting data, and creating JSON and CSV files..." + "`n"
    [System.Windows.Forms.Application]::DoEvents()

    # Split output by line and remove empty and header lines
    $lines = ($wingetOutput -split "`n") | Where-Object { $_ -notmatch '^-{2,}' -and $_ -notmatch '^Name' -and $_ }

    $packages = New-Object System.Collections.Generic.List[System.Object]
    
    # This regex is more robust and will not mistake a header line for an app
    $regex = '^(?<Name>.*?)\s{2,}(?<Id>\S+)\s{2,}(?<Version>\S*)\s{2,}(?<Source>\S+)$'
    
    $lines | ForEach-Object {
        if ($_ -match $regex) {
            $packages.Add([PSCustomObject]@{
                Name = $Matches.Name.Trim()
                Id = $Matches.Id
                Version = $Matches.Version
                Source = $Matches.Source
            })
        }
    }
    
    # Check if any packages were found after parsing
    if (-not $packages) {
        $textBoxVerbose.Text += "No winget packages were found." + "`n"
        [System.Windows.MessageBox]::Show("No winget packages found to export.", "Info")
        $progressBarStatus.Visibility = "Hidden"
        return
    }

    # Sort the packages alphabetically by name
    $sortedPackages = $packages | Sort-Object -Property Name

    # Now, build the CSV content from the parsed data
    $csvContent = New-Object System.Text.StringBuilder
    [void]$csvContent.AppendLine("`"Name`",`"Id`",`"Version`",`"Source`",`"Exportable`"")

    $sortedPackages | ForEach-Object {
        $exportableStatus = if ($_.Source -eq "winget") { "Yes" } else { "No - Source unavailable" }
        [void]$csvContent.AppendLine("`"$($_.Name)`",`"$($_.Id)`",`"$($_.Version)`",`"$($_.Source)`",`"$($exportableStatus)`"")
    }

    $csvPath = [System.IO.Path]::ChangeExtension($filePath, ".csv")
    $csvContent.ToString() | Out-File -FilePath $csvPath -Encoding UTF8 -Force

    $progressBarStatus.Value = 60
    $textBoxVerbose.Text += "Creating JSON manifest file..." + "`n"
    [System.Windows.Forms.Application]::DoEvents()
    
    # Create the JSON manifest from the parsed data
    $jsonExportObject = [PSCustomObject]@{
        CreationDate = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss.ffZ")
        Sources = @(
            [PSCustomObject]@{
                Details = "https://cdn.winget.microsoft.com/cache/source.json"
                Identifier = "winget"
            }
        )
        Packages = $packages
    }
    $jsonExportObject | ConvertTo-Json -Depth 5 | Out-File -FilePath $filePath -Encoding UTF8

    $progressBarStatus.Value = 100
    $textBoxVerbose.Text += "Files created successfully." + "`n"
    $textBoxVerbose.Text += "JSON file: $filePath" + "`n"
    $textBoxVerbose.Text += "CSV file: $csvPath" + "`n"
    [System.Windows.MessageBox]::Show("Export completed. Check both files for the application status.", "Success")
    $progressBarStatus.Visibility = "Hidden"
})

$buttonImport.Add_Click({
    $progressBarStatus.Visibility = "Visible"
    $progressBarStatus.Value = 0
    $textBoxVerbose.Text = ""
    $filePath = $textBoxFile.Text
    
    $progressBarStatus.Value = 20
    $textBoxVerbose.Text += "Starting winget import..." + "`n"
    [System.Windows.Forms.Application]::DoEvents()

    $result = & winget import -i $filePath --accept-source-agreements --accept-package-agreements --disable-interactivity 2>&1
    $textBoxVerbose.Text += $result
    
    $progressBarStatus.Value = 80
    
    $progressBarStatus.Value = 100
    $textBoxVerbose.Text += "Import complete." + "`n"
    [System.Windows.MessageBox]::Show("Import completed: $result", "Success")
    $progressBarStatus.Visibility = "Hidden"
})

$window.ShowDialog()
