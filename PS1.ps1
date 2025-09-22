#This Powershell script will map printers locally to both VPSX spoolers via LPR
#Mappings should include LRP ports to both VPSX spoolers using port pooling for redundancy

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create the form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Printer Mapper - LPR (Pooled)"
$form.Size = New-Object System.Drawing.Size(400, 400)
$form.StartPosition = "CenterScreen"

# Label: Printer Name
$labelPrinter = New-Object System.Windows.Forms.Label
$labelPrinter.Text = "Enter Printer Name:"
$labelPrinter.Location = New-Object System.Drawing.Point(20, 20)
$labelPrinter.AutoSize = $true
$form.Controls.Add($labelPrinter)

# TextBox: Printer Name
$textBoxPrinter = New-Object System.Windows.Forms.TextBox
$textBoxPrinter.Location = New-Object System.Drawing.Point(20, 45)
$textBoxPrinter.Width = 340
$form.Controls.Add($textBoxPrinter)

# GroupBox: Printer Type
$groupBox = New-Object System.Windows.Forms.GroupBox
$groupBox.Text = "Select Printer Type"
$groupBox.Location = New-Object System.Drawing.Point(20, 85)
$groupBox.Size = New-Object System.Drawing.Size(340, 70)
$form.Controls.Add($groupBox)

# Radio Buttons
$radioXerox = New-Object System.Windows.Forms.RadioButton
$radioXerox.Text = "Xerox"
$radioXerox.Location = New-Object System.Drawing.Point(10, 30)
$groupBox.Controls.Add($radioXerox)

$radioHP = New-Object System.Windows.Forms.RadioButton
$radioHP.Text = "HP"
$radioHP.Location = New-Object System.Drawing.Point(120, 30)
$groupBox.Controls.Add($radioHP)

# Button: Create Printer
$btnCreate = New-Object System.Windows.Forms.Button
$btnCreate.Text = "Create Printer"
$btnCreate.Location = New-Object System.Drawing.Point(20, 170)
$btnCreate.Size = New-Object System.Drawing.Size(340, 35)
$form.Controls.Add($btnCreate)

# Status Output Box
$textBoxStatus = New-Object System.Windows.Forms.TextBox
$textBoxStatus.Location = New-Object System.Drawing.Point(20, 220)
$textBoxStatus.Size = New-Object System.Drawing.Size(340, 120)
$textBoxStatus.Multiline = $true
$textBoxStatus.ScrollBars = "Vertical"
$textBoxStatus.ReadOnly = $true
$form.Controls.Add($textBoxStatus)

# Function to log to status box
function Write-Status {
    param([string]$msg)
    $textBoxStatus.AppendText("`r`n$msg")
}

# Button Click Event
$btnCreate.Add_Click({
    $PrinterBaseName = $textBoxPrinter.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($PrinterBaseName)) {
        [System.Windows.Forms.MessageBox]::Show("Printer name cannot be empty.", "Input Error", "OK", "Error")
        return
    }

    $PrinterName = "$PrinterBaseName-Pooled"

    # Determine printer type
    if ($radioXerox.Checked) {
        $DriverName = "Xerox Global Print Driver PCL6"
    }
    elseif ($radioHP.Checked) {
        $DriverName = "HP Universal Printing PCL 6 (v6.6.0)"
    }
    else {
        [System.Windows.Forms.MessageBox]::Show("Please select a printer type.", "Selection Required", "OK", "Warning")
        return
    }

    # Define LPR ports
    $Port1Name = "LPR_$PrinterBaseName`_1"
    $Port2Name = "LPR_$PrinterBaseName`_2"
    $Host1 = "trsmvpsxspl1.shcsd.sharp.com"
    $Host2 = "trsmvpsxspl2.shcsd.sharp.com"

    # Function to create LPR port
    function Create-LPRPort {
        param([string]$PortName, [string]$Host, [string]$Queue)
        if (-not (Get-PrinterPort -Name $PortName -ErrorAction SilentlyContinue)) {
            Write-Status "Creating port $PortName -> $Host : $Queue"
            Add-PrinterPort -Name $PortName `
                            -PrinterHostAddress $Host `
                            -PortNumber 515 `
                            -LprQueueName $Queue `
                            -Protocol LPR `
                            -SNMP $false
        } else {
            Write-Status "Port $PortName already exists."
        }
    }

    try {
        # Create ports
        Create-LPRPort -PortName $Port1Name -Host $Host1 -Queue $PrinterBaseName
        Create-LPRPort -PortName $Port2Name -Host $Host2 -Queue $PrinterBaseName

        # Check if printer exists
        if (Get-Printer -Name $PrinterName -ErrorAction SilentlyContinue) {
            Write-Status "Printer '$PrinterName' already exists. Aborting."
            return
        }

        Write-Status "Creating printer '$PrinterName' with driver '$DriverName'..."
        Add-Printer -Name $PrinterName -DriverName $DriverName -PortName $Port1Name

        Set-Printer -Name $PrinterName -PrinterPooling $true

        # Add second port
        Set-Printer -Name $PrinterName -PortName "$Port1Name,$Port2Name"

        Write-Status "✅ Printer '$PrinterName' created successfully with LPR pooling."
    }
    catch {
        Write-Status "❌ Error: $_"
    }
})

# Show the form
[void]$form.ShowDialog()
