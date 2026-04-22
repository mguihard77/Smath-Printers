###################################################
# SMATH PRINTERS - SCRIPT FINAL STABLE
###################################################

$ErrorActionPreference  = "SilentlyContinue"
$WarningPreference      = "SilentlyContinue"
$ProgressPreference     = "SilentlyContinue"
$ConfirmPreference      = "None"

$Summary = @()

###################################################
# UI - BARRE DE PROGRESSION
###################################################

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = "Deploiement des imprimantes"
$form.Size = New-Object System.Drawing.Size(460,120)
$form.StartPosition = "CenterScreen"
$form.TopMost = $true
$form.ControlBox = $false

$label = New-Object System.Windows.Forms.Label
$label.Text = "Initialisation..."
$label.AutoSize = $true
$label.Location = New-Object System.Drawing.Point(20,10)
$form.Controls.Add($label)

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(20,40)
$progressBar.Size = New-Object System.Drawing.Size(420,20)
$form.Controls.Add($progressBar)

$form.Show()
$form.Refresh()

function Set-Progress {
    param([int]$Percent,[string]$Text)

    if ($Percent -lt 0)   { $Percent = 0 }
    if ($Percent -gt 100){ $Percent = 100 }

    $progressBar.Value = $Percent
    $label.Text = $Text
    [System.Windows.Forms.Application]::DoEvents()
}

###################################################
# VERIFICATION PRESENCE config.csv
###################################################

Set-Progress 5 "Verification configuration"

$CsvPath = ".\config.csv"

if (-not (Test-Path $CsvPath)) {

    $null = [System.Windows.Forms.MessageBox]::Show(
        "Le fichier config.csv est introuvable.`n`nVeuillez verifier qu il est present dans le dossier.",
        "Smath Printers - Erreur",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )

    $form.Close()
    :Exit(1)
}

###################################################
# LECTURE DU CSV (lignes valides uniquement)
###################################################

Set-Progress 10 "Chargement configuration"
$config = Import-Csv $CsvPath | Where-Object {
    $_.Name -and $_.Name.Trim() -ne ""
}

###################################################
# SUPPRESSION DES IMPRIMANTES + PORTS
###################################################

Set-Progress 30 "Suppression anciennes imprimantes"

Get-Printer | Where-Object { $_.Name -like "Stratus*" } | ForEach-Object {

    $Summary += "Suppression : $($_.Name)"

    $port = $_.PortName
    Remove-Printer -Name $_.Name

    if ($port -like "IP_*") {
        Remove-PrinterPort -Name $port
    }
}

Start-Sleep -Seconds 2

###################################################
# CREATION DES IMPRIMANTES (CSV)
###################################################

Set-Progress 65 "Installation imprimantes"

foreach ($item in $config) {

    # Installation du driver via INF si fourni
    if ($item.InfPath -and -not (Get-PrinterDriver -Name $item.Driver)) {
        pnputil.exe /add-driver $item.InfPath /install | Out-Null
    }

    # Creation du port TCP/IP
    if (-not (Get-PrinterPort -Name $item.Port)) {
        Add-PrinterPort -Name $item.Port -PrinterHostAddress $item.IP
    }

    # Creation brute de l imprimante (comportement original)
    Add-Printer -Name $item.Name `
                -DriverName $item.Driver `
                -PortName $item.Port

    $Summary += "Installation : $($item.Name)"

    # Option Sharp
    if ($item.Type -eq "Sharp") {
        Get-CimInstance Win32_Printer -Filter "Name='$($item.Name)'" |
            Set-CimInstance -Property @{ EnableBidi = $false }
    }
}

###################################################
# FIN
###################################################

Set-Progress 100 "Finalisation"
$form.Close()

###################################################
# MESSAGE FINAL (UNIQUEMENT SI ACTIONS)
###################################################

if ($Summary.Count -gt 0) {
    $null = [System.Windows.Forms.MessageBox]::Show(
        ($Summary -join "`n"),
        "Deploiement des imprimantes termine",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    )
}

:Exit(0)