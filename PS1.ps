#This Powershell script will map printers locally to both VPSX spoolers via LPR
#Mappings should include LRP ports to both VPSX spoolers using port pooling for redundancy

# Global Variables
$global:PrinterCaption = ""
$global:Computer = ""
$global:COMPUTERS = ""
$global:DriverName = ""
$global:DriverPath = ""
$global:DriverInf = ""

# Driver Configuration
$DriverName = "Xerox Global Print Driver PCL6"
$DriverPath = "C:\Windows\System32\DriverStore\FileRepository\xerox_versalink_b600_b605_b610_b615_pcl6.inf_amd64_222cd7e394115a07\"
$DriverInf = "C:\Windows\System32\DriverStore\FileRepository\sfwemenu.inf_amd64_9a84be23a5070548\Xerox_VersaLink_B600_B605_B610_B615_PCL6.inf"

# Function: Create LPR Port via WMI
Function CreateLPRPort {
    param(
        [string]$ComputerName,
        [string]$PortName,
        [string]$HostName,
        [string]$QueueName
    )
    $wmi = [wmiclass]"\\$ComputerName\root\cimv2:win32_tcpipPrinterPort"
    $wmi.psbase.scope.options.enablePrivileges = $true
    $Port = $wmi.createInstance()
    $Port.Name = $PortName
    $Port.HostAddress = $HostName
    $Port.PortNumber = 515
    $Port.Protocol = 2   # LPR
    $Port.SNMPEnabled = $false
    $Port.Queue = $QueueName
    $Port.Put() | Out-Null
}

# Function: Install Printer Driver via WMI
Function InstallPrinterDriver {
    param(
        [string]$ComputerName,
        [string]$DriverName,
        [string]$DriverPath,
        [string]$DriverInf
    )
    $wmi = [wmiclass]"\\$ComputerName\Root\cimv2:Win32_PrinterDriver"
    $wmi.psbase.scope.options.enablePrivileges = $true
    $Driver = $wmi.CreateInstance()
    $Driver.Name = $DriverName
    $Driver.DriverPath = $DriverPath
    $Driver.InfName = $DriverInf
    $wmi.AddPrinterDriver($Driver)
    $wmi.Put() | out-null
}

# Function: Create Printer
Function CreatePrinter {
    param(
        [string]$ComputerName,
        [string]$PrinterName,
        [string]$DriverName,
        [string]$PrimaryPortName
    )
    $wmi = [wmiclass]"\\$ComputerName\Root\cimv2:Win32_Printer"
    $Printer = $wmi.CreateInstance()
    $Printer.Caption = $PrinterName
    $Printer.DriverName = $DriverName
    $Printer.PortName = $PrimaryPortName
    $Printer.DeviceID = $PrinterName
    $Printer.Put() | out-null
}

# Function: Add second port to existing printer (pooling)
Function EnablePrinterPooling {
    param(
        [string]$ComputerName,
        [string]$PrinterName,
        [string]$AllPorts
    )
    Invoke-Command -ComputerName $ComputerName -ScriptBlock {
        param($PrinterName, $AllPorts)
        Set-Printer -Name $PrinterName -PrinterPooling $true
        Set-Printer -Name $PrinterName -PortName $AllPorts
    } -ArgumentList $PrinterName, $AllPorts
}

# Function: Get Computer Names from User (GUI)
Function Get-ComputerName([string]$Message, [string]$WindowTitle, [string]$DefaultText) {
    Add-Type -AssemblyName System.Drawing
    Add-Type -AssemblyName System.Windows.Forms

    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Message
    $label.AutoSize = $true
    $label.Location = '10,10'

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = '10,40'
    $textBox.Size = '575,200'
    $textBox.Multiline = $true
    $textBox.ScrollBars = 'Both'
    $textBox.Text = $DefaultText

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "OK"
    $okButton.Location = '415,250'
    $okButton.Add_Click({ $form.Tag = $textBox.Text; $form.Close() })

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Cancel"
    $cancelButton.Location = '510,250'
    $cancelButton.Add_Click({ $form.Tag = $null; $form.Close() })

    $form = New-Object System.Windows.Forms.Form
    $form.Text = $WindowTitle
    $form.Size = '610,320'
    $form.StartPosition = "CenterScreen"
    $form.Topmost = $True
    $form.Controls.AddRange(@($label, $textBox, $okButton, $cancelButton))
    $form.ShowDialog() > $null

    return $form.Tag
}

# Function: Ping Test
Function Online {
	If (Test-Connection -ComputerName $Computer -Count 2 -BufferSize 16 -Quiet) {
		return $true
	}
    else {
		return $false
	}
}

# Start of Script
Clear-Host

# Ask for printer name
Write-Host "Enter the name of the printer to be installed:"
$BasePrinterName = Read-Host
$PrinterName = "$BasePrinterName-Pooled"

# Prompt for remote computer(s)
$COMPUTERS = Get-ComputerName -Message "Enter IS tags (one per line)" -WindowTitle "Computer(s) to receive printer $PrinterName" -DefaultText "Enter IS tags or scan them"
$COMPUTERS = ($COMPUTERS -split '[\r\n]') | Where-Object { $_ }

# Define spoolers and ports
$Spooler1 = "trsmvpsxspl1.shcsd.sharp.com"
$Spooler2 = "trsmvpsxspl2.shcsd.sharp.com"
$LPRQueue = $BasePrinterName
$Port1Name = "LPR_$BasePrinterName`_1"
$Port2Name = "LPR_$BasePrinterName`_2"

Clear-Host

foreach ($Computer in $COMPUTERS) {
    if (Online) {
        Write-Host "[$Computer] Creating LPR ports..." -ForegroundColor Cyan
        CreateLPRPort -ComputerName $Computer -PortName $Port1Name -HostName $Spooler1 -QueueName $LPRQueue
        CreateLPRPort -ComputerName $Computer -PortName $Port2Name -HostName $Spooler2 -QueueName $LPRQueue

        Write-Host "[$Computer] Installing driver..." -ForegroundColor Yellow
        InstallPrinterDriver -DriverName $DriverName -DriverPath $DriverPath -DriverInf $DriverInf -ComputerName $Computer

        Write-Host "[$Computer] Creating pooled printer $PrinterName..." -ForegroundColor Green
        CreatePrinter -ComputerName $Computer -PrinterName $PrinterName -DriverName $DriverName -PrimaryPortName $Port1Name

        Write-Host "[$Computer] Enabling port pooling..." -ForegroundColor Magenta
        $AllPorts = "$Port1Name,$Port2Name"
        EnablePrinterPooling -ComputerName $Computer -PrinterName $PrinterName -AllPorts $AllPorts

        Write-Host "[$Computer] ✅ Printer $PrinterName created successfully!" -ForegroundColor White
    }
    else {
        Write-Warning "[$Computer] ❌ Offline or unreachable"
    }
}

# Cleanup
Remove-Variable -Name * -ErrorAction SilentlyContinue
