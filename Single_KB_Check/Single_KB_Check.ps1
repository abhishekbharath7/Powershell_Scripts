$global:reportarray = @()
$UserName = [System.Environment]::UserName

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
# Create the main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "KB Installation Checker"
$form.Size = New-Object System.Drawing.Size(800, 650)
$form.StartPosition = "CenterScreen"
# KB Number Label
$kbLabel = New-Object System.Windows.Forms.Label
$kbLabel.Location = New-Object System.Drawing.Point(10, 20)
$kbLabel.Size = New-Object System.Drawing.Size(100, 20)
$kbLabel.Text = "KB Number:"
$form.Controls.Add($kbLabel)
# KB Number TextBox
$kbTextBox = New-Object System.Windows.Forms.TextBox
$kbTextBox.Location = New-Object System.Drawing.Point(120, 20)
$kbTextBox.Size = New-Object System.Drawing.Size(150, 20)
$form.Controls.Add($kbTextBox)
# Server List Label
$serverLabel = New-Object System.Windows.Forms.Label
$serverLabel.Location = New-Object System.Drawing.Point(10, 60)
$serverLabel.Size = New-Object System.Drawing.Size(100, 20)
$serverLabel.Text = "Servers:"
$form.Controls.Add($serverLabel)
# Server List TextBox
$serverTextBox = New-Object System.Windows.Forms.TextBox
$serverTextBox.Location = New-Object System.Drawing.Point(120, 60)
$serverTextBox.Size = New-Object System.Drawing.Size(650, 20)
$serverTextBox.Multiline = $true
$serverTextBox.ScrollBars = "Vertical"
$serverTextBox.Height = 150
$form.Controls.Add($serverTextBox)
# Progress Bar
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(10, 220)
$progressBar.Size = New-Object System.Drawing.Size(760, 20)
$progressBar.Style = "Continuous"
$progressBar.Visible = $false
$form.Controls.Add($progressBar)
# Check Button
$checkButton = New-Object System.Windows.Forms.Button
$checkButton.Location = New-Object System.Drawing.Point(240, 250)
$checkButton.Size = New-Object System.Drawing.Size(100, 30)
$checkButton.Text = "Check"
$checkButton.Add_Click({
   $dataGridView.Rows.Clear()
   $kb = $kbTextBox.Text.Trim()
   if (-not $kb.StartsWith("KB") -or -not ($kb.Substring(2) -match "^\d+$")) {
       [System.Windows.Forms.MessageBox]::Show("Please enter a valid KB number (e.g., KB1234567)")
       return
   }
   $servers = $serverTextBox.Text -split "`r?`n" | Where-Object { $_ -ne "" }
   $totalServers = $servers.Count
   if ($totalServers -gt 0) {
       $progressBar.Visible = $true
       $progressBar.Value = 0
       $progressBar.Maximum = $totalServers
   }
   foreach ($server in $servers) {

       $global:HotFixID = $Null
       $global:InstalledOnDate = $Null
       $global:Status = $Null

       $server = $server.Trim()
       $row = New-Object System.Windows.Forms.DataGridViewRow
       $row.CreateCells($dataGridView)
       $row.Cells[0].Value = $server
       $row.Cells[1].Value = "Checking..."
       $row.Cells[2].Value = "Checking..."
       $row.Cells[3].Value = "Checking..."
       $row.Cells[4].Value = "In Progress"
       $dataGridView.Rows.Add($row)
       try {
           # Get OS information
           $os = Get-CimInstance -ClassName Win32_OperatingSystem -ComputerName $server -ErrorAction Stop
           $row.Cells[3].Value = $os.Caption
           # Get Hotfix information
           $hotfix = Get-CimInstance -ClassName Win32_QuickFixEngineering -ComputerName $server -ErrorAction Stop |
                     Where-Object { $_.HotFixID -eq $kb }
           if ($hotfix) {
               $row.Cells[1].Value = $hotfix.HotFixID
               $row.Cells[2].Value = $hotfix.InstalledOn
               $row.Cells[4].Value = "Installed & Complaint"

               $HotFixID = $kb
               $InstalledOnDate = $hotfix.InstalledOn
               if($InstalledOnDate -ne $Null)
               {
                    $Status = "KB Installed"
               }
               else
               {
                    $Status = "KB Installed - Restart Required"
               }
           }
           else {
               $row.Cells[1].Value = "Not Installed & Non Complaint"
               $row.Cells[2].Value = "N/A"
               $row.Cells[4].Value = "Completed"

               $HotFixID = $kb
               $InstalledOnDate = "NA"
               $Status = "Not Installed & Non Complaint"
           }
       }
       catch {
           $row.Cells[1].Value = "Error"
           $row.Cells[2].Value = "N/A"
           $row.Cells[3].Value = "N/A"
           $row.Cells[4].Value = $_.Exception.Message

           $HotFixID = $kb
           $InstalledOnDate = "NA"
           $Status = $_.Exception.Message
       }
       $progressBar.PerformStep()
       $dataGridView.Refresh()
       [System.Windows.Forms.Application]::DoEvents()

       $global:output = [ordered]@{
                ServerName = $server
                KB = $HotFixID
                InstalledOnDate = $InstalledOnDate
                Status = $Status
            }
                
         $global:obj = New-Object -TypeName PSObject -Property $output
         $global:reportarray += $obj
   }
    $TimeStamp1 = Get-Date -Format dd-MM-yyyy_HH-mm-ss
    $file1 = $UserName+"-"+$TimeStamp1
    $reportarray | ConvertTo-Csv -NoTypeInformation | Out-File "C:\Script\KB_check\Logs\$file1.csv" -Force   # CHANGE THIS TO REQUIRED OUTPUT PATH
    
   $progressBar.Visible = $false
})



$form.Controls.Add($checkButton)
# Export Button
$exportButton = New-Object System.Windows.Forms.Button
$exportButton.Location = New-Object System.Drawing.Point(400, 250)
$exportButton.Size = New-Object System.Drawing.Size(100, 30)
$exportButton.Text = "Export CSV"
$exportButton.Add_Click({
   if ($dataGridView.Rows.Count -eq 0) {
       [System.Windows.Forms.MessageBox]::Show("No data to export!")
       return
   }
   $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
   $saveFileDialog.Filter = "CSV files (*.csv)|*.csv"
   $saveFileDialog.Title = "Save Results"
   $saveFileDialog.FileName = "KB_Check_Results.csv"
   $saveFileDialog.DefaultExt = "csv"
   if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
       $exportPath = $saveFileDialog.FileName
       $exportData = foreach ($row in $dataGridView.Rows) {
           if ($row.IsNewRow) { continue }
           [PSCustomObject]@{
               Server = $row.Cells[0].Value
               'Installed KB' = $row.Cells[1].Value
               'Installation Date' = $row.Cells[2].Value
               'OS Name' = $row.Cells[3].Value
               Status = $row.Cells[4].Value
           }
       }
       $exportData | Export-Csv -Path $exportPath -NoTypeInformation -Force
       [System.Windows.Forms.MessageBox]::Show("Data exported successfully to:`n$exportPath")
   }
})
$form.Controls.Add($exportButton)
# Results Grid
$dataGridView = New-Object System.Windows.Forms.DataGridView
$dataGridView.Location = New-Object System.Drawing.Point(10, 290)
$dataGridView.Size = New-Object System.Drawing.Size(760, 300)
$dataGridView.AutoSizeColumnsMode = "Fill"
$dataGridView.ColumnCount = 5
$dataGridView.Columns[0].Name = "Server"
$dataGridView.Columns[1].Name = "Installed KB"
$dataGridView.Columns[2].Name = "Installation Date"
$dataGridView.Columns[3].Name = "OS Name"
$dataGridView.Columns[4].Name = "Status"
$dataGridView.DefaultCellStyle.WrapMode = [System.Windows.Forms.DataGridViewTriState]::True
$dataGridView.AutoSizeRowsMode = [System.Windows.Forms.DataGridViewAutoSizeRowsMode]::DisplayedCells
$form.Controls.Add($dataGridView)
# Show the form
$form.ShowDialog()