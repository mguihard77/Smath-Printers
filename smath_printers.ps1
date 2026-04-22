###################################################
# SMATH PRINTERS
###################################################

function Main {

    $ErrorActionPreference  = "SilentlyContinue"
    $WarningPreference      = "SilentlyContinue"
    $ProgressPreference     = "SilentlyContinue"
    $ConfirmPreference      = "None"

    $Summary = @()

    ###################################################
    # LOAD WINFORMS
    ###################################################

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    ###################################################
    # BASE PATH
    ###################################################

    $BasePath = if ($PSScriptRoot) {
        $PSScriptRoot
    } else {
        [System.AppDomain]::CurrentDomain.BaseDirectory
    }

    ###################################################
    # LANGUAGE AUTO DETECTION
    ###################################################

    $DetectedLang = (Get-Culture).TwoLetterISOLanguageName
    $SupportedLangs = @("fr","en","de","es")

    if ($SupportedLangs -contains $DetectedLang) {
        $Lang = $DetectedLang
    } else {
        $Lang = "en"
    }

    ###################################################
    # CHECK LANG FOLDER
    ###################################################

    $LangFolder = Join-Path $BasePath "lang"

    if (-not (Test-Path $LangFolder)) {
        $null = [System.Windows.Forms.MessageBox]::Show(
            "Language folder is missing.`n`nThe 'lang' folder must be present.",
            "Smath Printers - Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return
    }

    $LangFile = Join-Path $LangFolder "$Lang.lang.ps1"

    if (-not (Test-Path $LangFile)) {
        $LangFile = Join-Path $LangFolder "en.lang.ps1"
    }

    if (-not (Test-Path $LangFile)) {
        $null = [System.Windows.Forms.MessageBox]::Show(
            "No valid language file found.",
            "Smath Printers - Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return
    }

    . $LangFile

    ###################################################
    # CHECK CSV
    ###################################################

    $CsvPath = Join-Path $BasePath "config.csv"

    if (-not (Test-Path $CsvPath)) {
        $null = [System.Windows.Forms.MessageBox]::Show(
            $Strings.ConfigMissingMessage,
            $Strings.ConfigMissingTitle,
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return
    }

    ###################################################
    # UI
    ###################################################

    $form = New-Object System.Windows.Forms.Form
    $form.Text = $Strings.AppTitle
    $form.Size = New-Object System.Drawing.Size(460,120)
    $form.StartPosition = "CenterScreen"
    $form.TopMost = $true
    $form.ControlBox = $false

    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Strings.Initializing
    $label.AutoSize = $true
    $label.Location = New-Object System.Drawing.Point(20,10)
    $form.Controls.Add($label)

    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Location = New-Object System.Drawing.Point(20,40)
    $progressBar.Size = New-Object System.Drawing.Size(420,20)
    $form.Controls.Add($progressBar)

    $form.Show()
    $form.Refresh()

    function Set-Progress($Percent,$Text){
        if ($Percent -lt 0) { $Percent = 0 }
        if ($Percent -gt 100) { $Percent = 100 }
        $progressBar.Value = $Percent
        $label.Text = $Text
        [System.Windows.Forms.Application]::DoEvents()
    }

    ###################################################
    # LOAD CSV
    ###################################################

    Set-Progress 10 $Strings.LoadingConfig

    $config = Import-Csv $CsvPath | Where-Object {
        $_.Name -and $_.Name.Trim() -ne ""
    }

    ###################################################
    # REMOVE PRINTERS
    ###################################################

    Set-Progress 30 $Strings.RemovingPrinters

    Get-Printer | Where-Object { $_.Name -like "Stratus*" } | ForEach-Object {
    $Summary += "Removal : $($_.Name)"

    $port = $_.PortName
    Remove-Printer -Name $_.Name

    if ($port -like "IP_*") {
        Remove-PrinterPort -Name $port -ErrorAction SilentlyContinue
    }
}

    Start-Sleep 2

    ###################################################
    # CREATE PRINTERS
    ###################################################

    Set-Progress 65 $Strings.InstallingPrinters

    foreach ($item in $config) {
        if ($item.InfPath -and -not (Get-PrinterDriver -Name $item.Driver)) {
            pnputil.exe /add-driver $item.InfPath /install | Out-Null
        }

        if (-not (Get-PrinterPort -Name $item.Port)) {
            Add-PrinterPort -Name $item.Port -PrinterHostAddress $item.IP
        }

        Add-Printer -Name $item.Name -DriverName $item.Driver -PortName $item.Port
        $Summary += "Installation : $($item.Name)"

        if ($item.Type -eq "Sharp") {
            Get-CimInstance Win32_Printer -Filter "Name='$($item.Name)'" |
                Set-CimInstance -Property @{ EnableBidi = $false }
        }
    }

    ###################################################
    # FINAL
    ###################################################

    Set-Progress 100 $Strings.Finalizing
    $form.Close()

    if ($Summary.Count -gt 0) {
        $null = [System.Windows.Forms.MessageBox]::Show(
            ($Summary -join "`n"),
            $Strings.SummaryTitle,
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    }
}

Main
