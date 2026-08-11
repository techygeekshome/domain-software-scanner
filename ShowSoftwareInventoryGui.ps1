#Requires -Version 5.1
<#
.SYNOPSIS
    A point-and-click front end for the Domain Software Inventory tool.

.DESCRIPTION
    Show-SoftwareInventoryGui.ps1 is a self-contained Windows Forms GUI that
    wraps Get-DomainSoftwareInventory.ps1. It lets you:

        1. Discover computers  - from Active Directory, an OU, a file or a
           typed list.
        2. Check services      - see, per machine, whether it's reachable and
           whether RemoteRegistry / WinRM are Running or Stopped.
        3. Start services       - start RemoteRegistry (or WinRM) on the
           selected machines, remotely, over DCOM - even when neither service
           is running yet (starting a service uses the Service Control Manager
           over RPC, which is independent of both).
        4. Scan                 - inventory third-party software + versions on
           the selected machines and open the CSV / HTML report.

    No installation, no dependencies beyond Windows PowerShell. Keep this
    script next to Get-DomainSoftwareInventory.ps1 and
    SoftwareInventory.Services.psm1.

.NOTES
    Run as a domain admin (or an account with local-admin on the targets).
    Launch under STA for a stable UI:
        powershell.exe -STA -ExecutionPolicy Bypass -File .\Show-SoftwareInventoryGui.ps1

    TechyGeeksHome - https://techygeekshome.info
#>

[CmdletBinding()]
param([string]$OutputFolder)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    [System.Windows.Forms.MessageBox]::Show(
        "For a stable UI, relaunch under STA:`n`n" +
        "powershell.exe -STA -ExecutionPolicy Bypass -File `".\Show-SoftwareInventoryGui.ps1`"",
        "Software Inventory", 'OK', 'Warning') | Out-Null
}

# --------------------------------------------------------------- locations ---
$script:ScriptRoot = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$script:CliScript  = Join-Path $script:ScriptRoot 'Get-DomainSoftwareInventory.ps1'
$script:SvcModule  = Join-Path $script:ScriptRoot 'SoftwareInventory.Services.psm1'
if (-not $OutputFolder) { $OutputFolder = $script:ScriptRoot }
$script:OutFolder  = $OutputFolder
$script:LastHtml   = $null
$script:LastCsv    = $null
$script:ActiveJob  = $null
$script:ActiveTimer = $null
$script:OnDone     = $null

if (Test-Path -LiteralPath $script:SvcModule) {
    Import-Module $script:SvcModule -Force -ErrorAction SilentlyContinue
}

# =============================================================== build UI ====
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Domain Software Inventory  -  TechyGeeksHome'
$form.Size = New-Object System.Drawing.Size(960, 720)
$form.MinimumSize = New-Object System.Drawing.Size(820, 640)
$form.StartPosition = 'CenterScreen'
$form.Font = New-Object System.Drawing.Font('Segoe UI', 9)

# ---- Targets group ----------------------------------------------------------
$grpTargets = New-Object System.Windows.Forms.GroupBox
$grpTargets.Text = '1. Choose target computers'
$grpTargets.Location = New-Object System.Drawing.Point(12, 10)
$grpTargets.Size = New-Object System.Drawing.Size(920, 118)
$grpTargets.Anchor = 'Top,Left,Right'
$form.Controls.Add($grpTargets)

$rbAll = New-Object System.Windows.Forms.RadioButton
$rbAll.Text = 'All domain computers (Active Directory)'
$rbAll.Location = New-Object System.Drawing.Point(14, 22)
$rbAll.Size = New-Object System.Drawing.Size(320, 22)
$rbAll.Checked = $true
$grpTargets.Controls.Add($rbAll)

$rbOU = New-Object System.Windows.Forms.RadioButton
$rbOU.Text = 'AD OU (distinguished name):'
$rbOU.Location = New-Object System.Drawing.Point(14, 47)
$rbOU.Size = New-Object System.Drawing.Size(210, 22)
$grpTargets.Controls.Add($rbOU)

$txtOU = New-Object System.Windows.Forms.TextBox
$txtOU.Location = New-Object System.Drawing.Point(230, 46)
$txtOU.Size = New-Object System.Drawing.Size(430, 23)
$txtOU.Anchor = 'Top,Left,Right'
$txtOU.Text = 'OU=Servers,DC=contoso,DC=com'
$grpTargets.Controls.Add($txtOU)

$rbFile = New-Object System.Windows.Forms.RadioButton
$rbFile.Text = 'From file:'
$rbFile.Location = New-Object System.Drawing.Point(14, 72)
$rbFile.Size = New-Object System.Drawing.Size(90, 22)
$grpTargets.Controls.Add($rbFile)

$txtFile = New-Object System.Windows.Forms.TextBox
$txtFile.Location = New-Object System.Drawing.Point(110, 71)
$txtFile.Size = New-Object System.Drawing.Size(470, 23)
$txtFile.Anchor = 'Top,Left,Right'
$grpTargets.Controls.Add($txtFile)

$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Text = 'Browse...'
$btnBrowse.Location = New-Object System.Drawing.Point(586, 70)
$btnBrowse.Size = New-Object System.Drawing.Size(74, 25)
$btnBrowse.Anchor = 'Top,Right'
$grpTargets.Controls.Add($btnBrowse)

$rbManual = New-Object System.Windows.Forms.RadioButton
$rbManual.Text = 'Typed list:'
$rbManual.Location = New-Object System.Drawing.Point(14, 94)
$rbManual.Size = New-Object System.Drawing.Size(90, 22)
$grpTargets.Controls.Add($rbManual)

$txtManual = New-Object System.Windows.Forms.TextBox
$txtManual.Location = New-Object System.Drawing.Point(110, 93)
$txtManual.Size = New-Object System.Drawing.Size(470, 23)
$txtManual.Anchor = 'Top,Left,Right'
$txtManual.Text = 'PC01, PC02, SRV05'
$grpTargets.Controls.Add($txtManual)

$btnDiscover = New-Object System.Windows.Forms.Button
$btnDiscover.Text = 'Load / Discover'
$btnDiscover.Location = New-Object System.Drawing.Point(700, 20)
$btnDiscover.Size = New-Object System.Drawing.Size(210, 40)
$btnDiscover.Anchor = 'Top,Right'
$grpTargets.Controls.Add($btnDiscover)

# ---- Grid -------------------------------------------------------------------
$grid = New-Object System.Windows.Forms.DataGridView
$grid.Location = New-Object System.Drawing.Point(12, 138)
$grid.Size = New-Object System.Drawing.Size(920, 250)
$grid.Anchor = 'Top,Bottom,Left,Right'
$grid.AllowUserToAddRows = $false
$grid.AllowUserToResizeRows = $false
$grid.RowHeadersVisible = $false
$grid.SelectionMode = 'FullRowSelect'
$grid.AutoSizeColumnsMode = 'Fill'
$grid.BackgroundColor = [System.Drawing.Color]::White
$script:grid = $grid
$form.Controls.Add($grid)

$colSel = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
$colSel.HeaderText = ''
$colSel.Name = 'Sel'
$colSel.FillWeight = 8
$grid.Columns.Add($colSel) | Out-Null

foreach ($c in @(
        @{ N = 'Computer'; W = 34 },
        @{ N = 'Reachable'; W = 16 },
        @{ N = 'RemoteRegistry'; W = 21 },
        @{ N = 'WinRM'; W = 21 })) {
    $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $col.HeaderText = $c.N
    $col.Name = $c.N
    $col.FillWeight = $c.W
    $col.ReadOnly = $true
    $grid.Columns.Add($col) | Out-Null
}

# ---- Selection / service buttons -------------------------------------------
$btnY = 396
$btnCheck = New-Object System.Windows.Forms.Button
$btnCheck.Text = 'Check services'
$btnCheck.Location = New-Object System.Drawing.Point(12, $btnY)
$btnCheck.Size = New-Object System.Drawing.Size(120, 30)
$btnCheck.Anchor = 'Bottom,Left'
$form.Controls.Add($btnCheck)

$btnStartRR = New-Object System.Windows.Forms.Button
$btnStartRR.Text = 'Start Remote Registry'
$btnStartRR.Location = New-Object System.Drawing.Point(140, $btnY)
$btnStartRR.Size = New-Object System.Drawing.Size(160, 30)
$btnStartRR.Anchor = 'Bottom,Left'
$form.Controls.Add($btnStartRR)

$btnStartWinRM = New-Object System.Windows.Forms.Button
$btnStartWinRM.Text = 'Start WinRM'
$btnStartWinRM.Location = New-Object System.Drawing.Point(308, $btnY)
$btnStartWinRM.Size = New-Object System.Drawing.Size(120, 30)
$btnStartWinRM.Anchor = 'Bottom,Left'
$form.Controls.Add($btnStartWinRM)

$btnSelAll = New-Object System.Windows.Forms.Button
$btnSelAll.Text = 'Select all'
$btnSelAll.Location = New-Object System.Drawing.Point(700, $btnY)
$btnSelAll.Size = New-Object System.Drawing.Size(105, 30)
$btnSelAll.Anchor = 'Bottom,Right'
$form.Controls.Add($btnSelAll)

$btnSelNone = New-Object System.Windows.Forms.Button
$btnSelNone.Text = 'Select none'
$btnSelNone.Location = New-Object System.Drawing.Point(810, $btnY)
$btnSelNone.Size = New-Object System.Drawing.Size(105, 30)
$btnSelNone.Anchor = 'Bottom,Right'
$form.Controls.Add($btnSelNone)

# ---- Options group ----------------------------------------------------------
$grpOpts = New-Object System.Windows.Forms.GroupBox
$grpOpts.Text = '2. Scan options'
$grpOpts.Location = New-Object System.Drawing.Point(12, 434)
$grpOpts.Size = New-Object System.Drawing.Size(920, 92)
$grpOpts.Anchor = 'Bottom,Left,Right'
$form.Controls.Add($grpOpts)

$chkMs = New-Object System.Windows.Forms.CheckBox
$chkMs.Text = 'Include Microsoft software'
$chkMs.Location = New-Object System.Drawing.Point(14, 22)
$chkMs.Size = New-Object System.Drawing.Size(200, 22)
$grpOpts.Controls.Add($chkMs)

$chkUpd = New-Object System.Windows.Forms.CheckBox
$chkUpd.Text = 'Include Windows updates'
$chkUpd.Location = New-Object System.Drawing.Point(14, 46)
$chkUpd.Size = New-Object System.Drawing.Size(200, 22)
$grpOpts.Controls.Add($chkUpd)

$chkUser = New-Object System.Windows.Forms.CheckBox
$chkUser.Text = 'Include per-user installs'
$chkUser.Location = New-Object System.Drawing.Point(230, 22)
$chkUser.Size = New-Object System.Drawing.Size(200, 22)
$grpOpts.Controls.Add($chkUser)

$lblTrans = New-Object System.Windows.Forms.Label
$lblTrans.Text = 'Transport:'
$lblTrans.Location = New-Object System.Drawing.Point(230, 48)
$lblTrans.Size = New-Object System.Drawing.Size(66, 22)
$grpOpts.Controls.Add($lblTrans)

$rbRR = New-Object System.Windows.Forms.RadioButton
$rbRR.Text = 'Remote Registry'
$rbRR.Location = New-Object System.Drawing.Point(300, 47)
$rbRR.Size = New-Object System.Drawing.Size(130, 22)
$rbRR.Checked = $true
$grpOpts.Controls.Add($rbRR)

$rbWinRM = New-Object System.Windows.Forms.RadioButton
$rbWinRM.Text = 'WinRM'
$rbWinRM.Location = New-Object System.Drawing.Point(430, 47)
$rbWinRM.Size = New-Object System.Drawing.Size(80, 22)
$grpOpts.Controls.Add($rbWinRM)

$chkAuto = New-Object System.Windows.Forms.CheckBox
$chkAuto.Text = 'Auto-start service before scan'
$chkAuto.Location = New-Object System.Drawing.Point(530, 22)
$chkAuto.Size = New-Object System.Drawing.Size(220, 22)
$chkAuto.Checked = $true
$grpOpts.Controls.Add($chkAuto)

$chkRestore = New-Object System.Windows.Forms.CheckBox
$chkRestore.Text = 'Restore service state after scan'
$chkRestore.Location = New-Object System.Drawing.Point(530, 46)
$chkRestore.Size = New-Object System.Drawing.Size(230, 22)
$chkRestore.Checked = $true
$grpOpts.Controls.Add($chkRestore)

$lblThr = New-Object System.Windows.Forms.Label
$lblThr.Text = 'Parallel:'
$lblThr.Location = New-Object System.Drawing.Point(770, 22)
$lblThr.Size = New-Object System.Drawing.Size(56, 22)
$grpOpts.Controls.Add($lblThr)

$numThr = New-Object System.Windows.Forms.NumericUpDown
$numThr.Location = New-Object System.Drawing.Point(828, 20)
$numThr.Size = New-Object System.Drawing.Size(70, 23)
$numThr.Minimum = 1
$numThr.Maximum = 256
$numThr.Value = 32
$grpOpts.Controls.Add($numThr)

# ---- Scan / progress / open -------------------------------------------------
$scanY = 534
$btnScan = New-Object System.Windows.Forms.Button
$btnScan.Text = 'Scan selected computers'
$btnScan.Location = New-Object System.Drawing.Point(12, $scanY)
$btnScan.Size = New-Object System.Drawing.Size(200, 34)
$btnScan.Anchor = 'Bottom,Left'
$btnScan.BackColor = [System.Drawing.Color]::FromArgb(76, 155, 255)
$btnScan.ForeColor = [System.Drawing.Color]::White
$btnScan.FlatStyle = 'Flat'
$form.Controls.Add($btnScan)

$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Location = New-Object System.Drawing.Point(222, $scanY + 4)
$progress.Size = New-Object System.Drawing.Size(430, 24)
$progress.Anchor = 'Bottom,Left,Right'
$progress.Style = 'Blocks'
$script:progress = $progress
$form.Controls.Add($progress)

$btnOpenHtml = New-Object System.Windows.Forms.Button
$btnOpenHtml.Text = 'Open HTML'
$btnOpenHtml.Location = New-Object System.Drawing.Point(700, $scanY)
$btnOpenHtml.Size = New-Object System.Drawing.Size(105, 34)
$btnOpenHtml.Anchor = 'Bottom,Right'
$btnOpenHtml.Enabled = $false
$form.Controls.Add($btnOpenHtml)

$btnOpenCsv = New-Object System.Windows.Forms.Button
$btnOpenCsv.Text = 'Open CSV'
$btnOpenCsv.Location = New-Object System.Drawing.Point(810, $scanY)
$btnOpenCsv.Size = New-Object System.Drawing.Size(105, 34)
$btnOpenCsv.Anchor = 'Bottom,Right'
$btnOpenCsv.Enabled = $false
$form.Controls.Add($btnOpenCsv)

# ---- Log --------------------------------------------------------------------
$log = New-Object System.Windows.Forms.TextBox
$log.Location = New-Object System.Drawing.Point(12, 576)
$log.Size = New-Object System.Drawing.Size(920, 100)
$log.Anchor = 'Bottom,Left,Right'
$log.Multiline = $true
$log.ReadOnly = $true
$log.ScrollBars = 'Vertical'
$log.BackColor = [System.Drawing.Color]::FromArgb(24, 30, 38)
$log.ForeColor = [System.Drawing.Color]::FromArgb(160, 230, 160)
$script:log = $log
$form.Controls.Add($log)

$script:ActionButtons = @(
    $btnDiscover, $btnCheck, $btnStartRR, $btnStartWinRM, $btnScan, $btnBrowse
)

# =============================================================== helpers ======
function Write-Log {
    param([string]$Message)
    $line = ('[{0}] {1}{2}' -f (Get-Date -Format 'HH:mm:ss'), $Message, [Environment]::NewLine)
    $script:log.AppendText($line)
}

function Set-Busy {
    param([bool]$Busy, [string]$Status)
    if ($Busy) {
        $script:progress.Style = 'Marquee'
        $script:progress.MarqueeAnimationSpeed = 30
    } else {
        $script:progress.Style = 'Blocks'
        $script:progress.MarqueeAnimationSpeed = 0
        $script:progress.Value = 0
    }
    foreach ($b in $script:ActionButtons) { $b.Enabled = -not $Busy }
    if ($Status) { Write-Log $Status }
}

function Get-CheckedComputers {
    $script:grid.EndEdit()   # commit any half-clicked checkbox
    $list = New-Object System.Collections.Generic.List[string]
    foreach ($row in $script:grid.Rows) {
        $sel = $row.Cells['Sel'].Value
        $name = $row.Cells['Computer'].Value
        if ($sel -eq $true -and $name) { $list.Add([string]$name) }
    }
    return $list
}

function Set-ComputerRows {
    param([string[]]$Names)
    $script:grid.Rows.Clear()
    foreach ($n in ($Names | Where-Object { $_ } | Sort-Object -Unique)) {
        $idx = $script:grid.Rows.Add()
        $script:grid.Rows[$idx].Cells['Sel'].Value = $true
        $script:grid.Rows[$idx].Cells['Computer'].Value = $n
        $script:grid.Rows[$idx].Cells['Reachable'].Value = '?'
        $script:grid.Rows[$idx].Cells['RemoteRegistry'].Value = '?'
        $script:grid.Rows[$idx].Cells['WinRM'].Value = '?'
    }
}

function Set-ServiceCell {
    param($Row, [string]$Column, [string]$Value)
    $cell = $Row.Cells[$Column]
    $cell.Value = $Value
    $cell.Style.ForeColor = switch -Regex ($Value) {
        '^Running$'            { [System.Drawing.Color]::Green; break }
        '^(Stopped|Disabled)$' { [System.Drawing.Color]::Firebrick; break }
        '^Yes$'                { [System.Drawing.Color]::Green; break }
        '^No$'                 { [System.Drawing.Color]::Firebrick; break }
        default                { [System.Drawing.Color]::DimGray }
    }
}

# Generic background-work runner: runs $Work in a job, polls with a UI timer,
# then calls $OnComplete with the job's results. Keeps the UI responsive.
function Start-BackgroundWork {
    param(
        [scriptblock]$Work,
        [object[]]$Arguments,
        [scriptblock]$OnComplete,
        [string]$BusyMessage
    )
    if ($script:ActiveJob) {
        [System.Windows.Forms.MessageBox]::Show('Please wait for the current operation to finish.',
            'Busy', 'OK', 'Information') | Out-Null
        return
    }
    Set-Busy -Busy $true -Status $BusyMessage
    $script:OnDone = $OnComplete
    $script:ActiveJob = Start-Job -ScriptBlock $Work -ArgumentList $Arguments

    $script:ActiveTimer = New-Object System.Windows.Forms.Timer
    $script:ActiveTimer.Interval = 500
    $script:ActiveTimer.Add_Tick({
        if (-not $script:ActiveJob) { $script:ActiveTimer.Stop(); return }
        if ($script:ActiveJob.State -in @('Completed', 'Failed', 'Stopped')) {
            $script:ActiveTimer.Stop()
            $results = @()
            try { $results = @(Receive-Job -Job $script:ActiveJob -ErrorAction SilentlyContinue) } catch { }
            Remove-Job -Job $script:ActiveJob -Force -ErrorAction SilentlyContinue
            $script:ActiveJob = $null
            Set-Busy -Busy $false -Status $null
            if ($script:OnDone) { & $script:OnDone $results }
        }
    })
    $script:ActiveTimer.Start()
}

# =============================================================== actions ======

# --- Discover / load targets -------------------------------------------------
$btnDiscover.Add_Click({
    if ($rbManual.Checked) {
        $names = $txtManual.Text -split '[,;\s]+' | Where-Object { $_ }
        Set-ComputerRows -Names $names
        Write-Log ("Loaded {0} computer(s) from the typed list." -f $script:grid.Rows.Count)
        return
    }
    if ($rbFile.Checked) {
        if (-not (Test-Path -LiteralPath $txtFile.Text)) {
            [System.Windows.Forms.MessageBox]::Show('File not found.', 'Load', 'OK', 'Warning') | Out-Null
            return
        }
        $names = Get-Content -LiteralPath $txtFile.Text |
            ForEach-Object { $_.Trim() } |
            Where-Object { $_ -and $_ -notmatch '^\s*#' }
        Set-ComputerRows -Names $names
        Write-Log ("Loaded {0} computer(s) from file." -f $script:grid.Rows.Count)
        return
    }

    # AD (all or OU) - run in a job so the UI stays responsive.
    $searchBase = if ($rbOU.Checked) { $txtOU.Text } else { $null }
    $work = {
        param($SearchBase)
        if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
            return [pscustomobject]@{ Error = 'ActiveDirectory RSAT module not installed.'; Names = @() }
        }
        Import-Module ActiveDirectory -ErrorAction Stop
        $p = @{ Filter = 'Enabled -eq $true'; Properties = 'DNSHostName' }
        if ($SearchBase) { $p.SearchBase = $SearchBase }
        $names = Get-ADComputer @p | ForEach-Object {
            if ($_.DNSHostName) { $_.DNSHostName } else { $_.Name }
        }
        return [pscustomobject]@{ Error = $null; Names = @($names) }
    }
    Start-BackgroundWork -Work $work -Arguments @($searchBase) -BusyMessage 'Querying Active Directory...' -OnComplete {
        param($res)
        $r = $res | Select-Object -Last 1
        if ($null -eq $r) { Write-Log 'AD query returned nothing.'; return }
        if ($r.Error) {
            [System.Windows.Forms.MessageBox]::Show($r.Error, 'Active Directory', 'OK', 'Warning') | Out-Null
            Write-Log ("AD error: {0}" -f $r.Error)
            return
        }
        Set-ComputerRows -Names $r.Names
        Write-Log ("Discovered {0} computer(s) from Active Directory." -f $script:grid.Rows.Count)
    }
})

$btnBrowse.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = 'Text files (*.txt)|*.txt|All files (*.*)|*.*'
    if ($dlg.ShowDialog() -eq 'OK') { $txtFile.Text = $dlg.FileName; $rbFile.Checked = $true }
})

$btnSelAll.Add_Click({ foreach ($r in $script:grid.Rows) { $r.Cells['Sel'].Value = $true } })
$btnSelNone.Add_Click({ foreach ($r in $script:grid.Rows) { $r.Cells['Sel'].Value = $false } })

# --- Check services ----------------------------------------------------------
$btnCheck.Add_Click({
    $computers = Get-CheckedComputers
    if ($computers.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('Tick at least one computer first.', 'Check', 'OK', 'Information') | Out-Null
        return
    }
    $work = {
        param($Module, $Computers)
        Import-Module $Module -Force
        $Computers | ForEach-Object {
            Get-SIServiceState -ComputerName $_ -Name @('RemoteRegistry', 'WinRM')
        }
    }
    Start-BackgroundWork -Work $work -Arguments @($script:SvcModule, $computers.ToArray()) `
        -BusyMessage ("Checking services on {0} machine(s)..." -f $computers.Count) -OnComplete {
        param($states)
        $map = @{}
        foreach ($s in $states) { if ($s -and $s.ComputerName) { $map[$s.ComputerName] = $s } }
        foreach ($row in $script:grid.Rows) {
            $name = [string]$row.Cells['Computer'].Value
            if (-not $map.ContainsKey($name)) { continue }
            $s = $map[$name]
            Set-ServiceCell -Row $row -Column 'Reachable' -Value $(if ($s.Online) { 'Yes' } else { 'No' })
            Set-ServiceCell -Row $row -Column 'RemoteRegistry' -Value ([string]$s.RemoteRegistry)
            Set-ServiceCell -Row $row -Column 'WinRM' -Value ([string]$s.WinRM)
        }
        Write-Log ("Checked services on {0} machine(s)." -f $states.Count)
    }
})

# --- Start a service on the selected machines --------------------------------
function Invoke-StartService {
    param([string]$ServiceName)
    $computers = Get-CheckedComputers
    if ($computers.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('Tick at least one computer first.', 'Start service', 'OK', 'Information') | Out-Null
        return
    }
    if ($ServiceName -eq 'WinRM') {
        Write-Log 'Note: starting the WinRM service does not configure a listener/firewall. If PS Remoting was never enabled, use Group Policy (Enable-PSRemoting).'
    }
    $work = {
        param($Module, $Computers, $Svc)
        Import-Module $Module -Force
        $Computers | ForEach-Object {
            Set-SIService -ComputerName $_ -Name $Svc -StartMode 'Manual' -Action 'Start'
        }
    }
    Start-BackgroundWork -Work $work -Arguments @($script:SvcModule, $computers.ToArray(), $ServiceName) `
        -BusyMessage ("Starting {0} on {1} machine(s)..." -f $ServiceName, $computers.Count) -OnComplete {
        param($results)
        $ok = 0; $fail = 0
        $map = @{}
        foreach ($r in $results) { if ($r -and $r.ComputerName) { $map[$r.ComputerName] = $r } }
        foreach ($row in $script:grid.Rows) {
            $name = [string]$row.Cells['Computer'].Value
            if (-not $map.ContainsKey($name)) { continue }
            $r = $map[$name]
            if ($r.Success -and $r.State -eq 'Running') { $ok++ } else { $fail++ }
            $col = if ($ServiceName -eq 'WinRM') { 'WinRM' } else { 'RemoteRegistry' }
            Set-ServiceCell -Row $row -Column $col -Value ([string]$r.State)
            if ($r.Error) { Write-Log ("{0}: {1}" -f $name, $r.Error) }
        }
        Write-Log ("{0}: started on {1}, failed on {2}." -f $ServiceName, $ok, $fail)
    }
}
$btnStartRR.Add_Click({ Invoke-StartService -ServiceName 'RemoteRegistry' })
$btnStartWinRM.Add_Click({ Invoke-StartService -ServiceName 'WinRM' })

# --- Scan --------------------------------------------------------------------
$btnScan.Add_Click({
    if (-not (Test-Path -LiteralPath $script:CliScript)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Get-DomainSoftwareInventory.ps1 was not found next to this GUI.",
            'Scan', 'OK', 'Error') | Out-Null
        return
    }
    $computers = Get-CheckedComputers
    if ($computers.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('Tick at least one computer first.', 'Scan', 'OK', 'Information') | Out-Null
        return
    }

    $opts = @{
        OutputFolder     = $script:OutFolder
        Throttle         = [int]$numThr.Value
        Ms               = [bool]$chkMs.Checked
        Updates          = [bool]$chkUpd.Checked
        User             = [bool]$chkUser.Checked
        WinRM            = [bool]$rbWinRM.Checked
        AutoStart        = [bool]$chkAuto.Checked
        Restore          = [bool]$chkRestore.Checked
    }
    $btnOpenHtml.Enabled = $false
    $btnOpenCsv.Enabled  = $false

    $work = {
        param($Cli, $Computers, $Opts)
        try { Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force -ErrorAction SilentlyContinue } catch { }
        $p = @{
            ComputerName  = $Computers
            OutputFolder  = $Opts.OutputFolder
            ThrottleLimit = $Opts.Throttle
        }
        if ($Opts.Ms)        { $p.IncludeMicrosoft = $true }
        if ($Opts.Updates)   { $p.IncludeUpdates = $true }
        if ($Opts.User)      { $p.IncludeUserSoftware = $true }
        if ($Opts.WinRM)     { $p.UseWinRM = $true }
        if ($Opts.AutoStart) { $p.AutoStartService = $true }
        if ($Opts.Restore)   { $p.RestoreServiceState = $true }

        $rows = & $Cli @p
        # Find the freshest report files the CLI just wrote.
        $html = Get-ChildItem -LiteralPath $Opts.OutputFolder -Filter 'SoftwareInventory_*.html' -ErrorAction SilentlyContinue |
            Sort-Object LastWriteTime -Descending | Select-Object -First 1
        $csv = Get-ChildItem -LiteralPath $Opts.OutputFolder -Filter 'SoftwareInventory_*.csv' -ErrorAction SilentlyContinue |
            Sort-Object LastWriteTime -Descending | Select-Object -First 1
        [pscustomobject]@{
            Rows = @($rows).Count
            Html = if ($html) { $html.FullName } else { $null }
            Csv  = if ($csv) { $csv.FullName } else { $null }
        }
    }

    Start-BackgroundWork -Work $work -Arguments @($script:CliScript, $computers.ToArray(), $opts) `
        -BusyMessage ("Scanning {0} computer(s) - this can take a while..." -f $computers.Count) -OnComplete {
        param($res)
        $r = $res | Where-Object { $_ -and $_.PSObject.Properties['Rows'] } | Select-Object -Last 1
        if ($null -eq $r) {
            Write-Log 'Scan finished but produced no report - check permissions / connectivity.'
            return
        }
        $script:LastHtml = $r.Html
        $script:LastCsv  = $r.Csv
        $btnOpenHtml.Enabled = [bool]$r.Html
        $btnOpenCsv.Enabled  = [bool]$r.Csv
        Write-Log ("Scan complete: {0} software row(s)." -f $r.Rows)
        if ($r.Html) { Write-Log ("HTML report: {0}" -f $r.Html) }
        if ($r.Csv)  { Write-Log ("CSV report:  {0}" -f $r.Csv) }
    }
})

$btnOpenHtml.Add_Click({ if ($script:LastHtml) { Start-Process $script:LastHtml } })
$btnOpenCsv.Add_Click({ if ($script:LastCsv) { Start-Process $script:LastCsv } })

# --- radio enable/disable niceties ------------------------------------------
$updateTargetInputs = {
    $txtOU.Enabled = $rbOU.Checked
    $txtFile.Enabled = $rbFile.Checked
    $btnBrowse.Enabled = $rbFile.Checked
    $txtManual.Enabled = $rbManual.Checked
}
$rbAll.Add_CheckedChanged($updateTargetInputs)
$rbOU.Add_CheckedChanged($updateTargetInputs)
$rbFile.Add_CheckedChanged($updateTargetInputs)
$rbManual.Add_CheckedChanged($updateTargetInputs)
& $updateTargetInputs

Write-Log 'Ready. Choose targets, click Load / Discover, tick machines, then Check services or Scan.'
if (-not (Test-Path -LiteralPath $script:SvcModule)) {
    Write-Log 'Warning: SoftwareInventory.Services.psm1 not found - service check/start will not work.'
}

# --------------------------------------------------------------- clean up ----
$form.Add_FormClosing({
    if ($script:ActiveTimer) { $script:ActiveTimer.Stop() }
    if ($script:ActiveJob) {
        Stop-Job -Job $script:ActiveJob -ErrorAction SilentlyContinue
        Remove-Job -Job $script:ActiveJob -Force -ErrorAction SilentlyContinue
    }
})

[void]$form.ShowDialog()
$form.Dispose()
