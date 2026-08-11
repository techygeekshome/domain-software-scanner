#Requires -Version 5.1
<#
.SYNOPSIS
    Inventories all installed third-party software (and its version) across
    every computer in an Active Directory domain.

.DESCRIPTION
    Get-DomainSoftwareInventory queries Active Directory for computer accounts,
    then reads each machine's "Uninstall" registry keys remotely to build a
    complete list of installed applications and their versions - the same data
    Windows shows in "Apps & features" / "Programs and Features".

    It reads BOTH registry views on every target:
        HKLM\...\Uninstall                (native 64-bit apps)
        HKLM\...\WOW6432Node\...\Uninstall (32-bit apps on 64-bit Windows)
    and, optionally, per-user installs under each loaded HKU hive.

    Reading the registry directly is deliberate: it is fast and safe. The old
    Win32_Product WMI class is intentionally NOT used because querying it
    triggers a costly MSI "reconfiguration" (consistency check) of every
    installed package on the target and can take minutes per machine.

    By default the report shows only THIRD-PARTY software: Microsoft/Windows
    operating-system components, security patches, hotfixes and Windows Updates
    are filtered out. Use the switches below to include them.

    Two transport methods are supported for reading the remote registry:
        Remote Registry (default) - uses the RemoteRegistry service + RPC.
                                     No WinRM required.
        WinRM / PS Remoting        - use -UseWinRM if you have PSRemoting
                                     enabled and Remote Registry disabled.

    Results are written to a timestamped CSV and a self-contained HTML report,
    and are also returned as objects on the pipeline so you can post-process
    them however you like.

.PARAMETER SearchBase
    Optional AD distinguished name to scope the computer search, e.g.
    "OU=Servers,DC=contoso,DC=com". Defaults to the whole domain.

.PARAMETER ComputerName
    One or more explicit computer names to scan instead of querying AD.
    Handy for testing or for a subset of machines.

.PARAMETER InputFile
    Path to a text file with one computer name per line. Overrides AD lookup.

.PARAMETER IncludeMicrosoft
    Include Microsoft-published software in the report (Office, Visual C++
    redistributables, .NET, etc.). Off by default.

.PARAMETER IncludeUpdates
    Include Windows/security updates, hotfixes and KB patches. Off by default.

.PARAMETER IncludeUserSoftware
    Also inventory per-user installs (HKU hives that are currently loaded).
    Off by default; only currently logged-on / loaded profiles are visible.

.PARAMETER UseWinRM
    Read the remote registry over PowerShell Remoting (WinRM) instead of the
    Remote Registry service.

.PARAMETER AutoStartService
    Before scanning, make sure the transport service is running on each target
    (RemoteRegistry by default, or WinRM with -UseWinRM). If it is stopped or
    disabled it is started remotely over DCOM/RPC - which does not itself need
    RemoteRegistry or WinRM. Needs admin rights and WMI reachable on the target.

.PARAMETER RestoreServiceState
    After scanning, stop any service this run started and restore its original
    start mode, leaving the estate exactly as it was found. Recommended when
    using -AutoStartService.

.PARAMETER OnlyEnabled
    When querying AD, skip disabled computer accounts. On by default.

.PARAMETER ThrottleLimit
    Maximum number of computers scanned in parallel. Default 32.

.PARAMETER OutputFolder
    Where to write the CSV and HTML reports. Defaults to the script folder.

.PARAMETER Credential
    Alternate credentials (a domain admin able to read the remote registry).

.EXAMPLE
    .\Get-DomainSoftwareInventory.ps1
    Scans every enabled computer in the domain and writes CSV + HTML reports.

.EXAMPLE
    .\Get-DomainSoftwareInventory.ps1 -SearchBase "OU=Servers,DC=contoso,DC=com" -IncludeMicrosoft
    Scans only the Servers OU and includes Microsoft software.

.EXAMPLE
    .\Get-DomainSoftwareInventory.ps1 -ComputerName PC01,PC02 -UseWinRM -Verbose
    Scans two named machines over WinRM with verbose progress.

.NOTES
    Author : TechyGeeksHome  -  https://techygeekshome.info
    Run as : a domain account with local-admin rights on the targets
             (needed to read HKLM remotely).
    Requires: RSAT ActiveDirectory module when querying AD (not needed when
              you pass -ComputerName or -InputFile).

    This script is READ-ONLY. It never installs, changes or removes anything
    on the machines it scans - it only reads registry values.
#>

[CmdletBinding(DefaultParameterSetName = 'AD')]
param(
    [Parameter(ParameterSetName = 'AD')]
    [string]$SearchBase,

    [Parameter(ParameterSetName = 'Computers', Mandatory)]
    [string[]]$ComputerName,

    [Parameter(ParameterSetName = 'File', Mandatory)]
    [string]$InputFile,

    [switch]$IncludeMicrosoft,
    [switch]$IncludeUpdates,
    [switch]$IncludeUserSoftware,
    [switch]$UseWinRM,

    # Before scanning, ensure the transport service (RemoteRegistry, or WinRM
    # with -UseWinRM) is running on each target - starting it remotely over
    # DCOM if needed. Requires admin rights + WMI reachable on the targets.
    [switch]$AutoStartService,

    # After scanning, put any service we started back the way we found it
    # (stop it and restore its original start mode). Only affects machines the
    # scan actually started.
    [switch]$RestoreServiceState,

    [Parameter(ParameterSetName = 'AD')]
    [bool]$OnlyEnabled = $true,

    [ValidateRange(1, 256)]
    [int]$ThrottleLimit = 32,

    [string]$OutputFolder,

    [System.Management.Automation.PSCredential]$Credential
)

# ------------------------------------------------------------------ setup ----
$ErrorActionPreference = 'Stop'
$startedAt = Get-Date

if (-not $OutputFolder) {
    $OutputFolder = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
}
if (-not (Test-Path -LiteralPath $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
}

$stamp    = $startedAt.ToString('yyyy-MM-dd_HHmmss')
$csvPath  = Join-Path $OutputFolder "SoftwareInventory_$stamp.csv"
$htmlPath = Join-Path $OutputFolder "SoftwareInventory_$stamp.html"

Write-Host "Ultimate Settings Panel - Domain Software Inventory" -ForegroundColor Cyan
Write-Host "TechyGeeksHome  -  https://techygeekshome.info`n" -ForegroundColor DarkCyan

# Service-control helpers (used only for -AutoStartService / -RestoreServiceState).
$servicesModule = Join-Path $PSScriptRoot 'SoftwareInventory.Services.psm1'
if (($AutoStartService -or $RestoreServiceState)) {
    if (Test-Path -LiteralPath $servicesModule) {
        Import-Module $servicesModule -Force -ErrorAction Stop
    } else {
        Write-Warning ("SoftwareInventory.Services.psm1 not found next to the script - " +
                       "-AutoStartService / -RestoreServiceState will be ignored.")
        $AutoStartService = $false
        $RestoreServiceState = $false
    }
}

# ---------------------------------------------------- build the target list --
function Get-TargetComputers {
    switch ($PSCmdlet.ParameterSetName) {
        'Computers' {
            return $ComputerName | Where-Object { $_ } | Sort-Object -Unique
        }
        'File' {
            if (-not (Test-Path -LiteralPath $InputFile)) {
                throw "Input file not found: $InputFile"
            }
            return Get-Content -LiteralPath $InputFile |
                ForEach-Object { $_.Trim() } |
                Where-Object { $_ -and $_ -notmatch '^\s*#' } |
                Sort-Object -Unique
        }
        default {
            if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
                throw "The ActiveDirectory module (RSAT) is not installed. " +
                      "Install RSAT, or use -ComputerName / -InputFile instead."
            }
            Import-Module ActiveDirectory -ErrorAction Stop

            $params = @{
                Filter     = 'Enabled -eq $true'
                Properties = 'DNSHostName'
            }
            if (-not $OnlyEnabled) { $params.Filter = '*' }
            if ($SearchBase)       { $params.SearchBase = $SearchBase }

            Write-Verbose "Querying Active Directory for computer accounts..."
            return Get-ADComputer @params |
                ForEach-Object { if ($_.DNSHostName) { $_.DNSHostName } else { $_.Name } } |
                Sort-Object -Unique
        }
    }
}

# ---------------------------------------------- the remote-worker scriptblock -
# This runs once per target computer (in a parallel runspace / job). It is
# self-contained on purpose so it can be shipped to a remote session too.
$worker = {
    param(
        [string]$Computer,
        [bool]$IncludeMicrosoft,
        [bool]$IncludeUpdates,
        [bool]$IncludeUserSoftware,
        [bool]$UseWinRM,
        [System.Management.Automation.PSCredential]$Credential
    )

    # Reachability check first - avoids long RPC timeouts on dead hosts.
    if (-not (Test-Connection -ComputerName $Computer -Count 1 -Quiet -ErrorAction SilentlyContinue)) {
        return [pscustomobject]@{
            ComputerName = $Computer
            Status       = 'Offline'
            Error        = 'No ping response'
            Apps         = @()
        }
    }

    # --- parses the entries under a single "Uninstall" registry path ---------
    # Implemented as a string of code so it can run either locally (Remote
    # Registry) or inside an Invoke-Command session unchanged.
    $collector = {
        param($Computer, $IncludeUser)

        $apps = New-Object System.Collections.Generic.List[object]

        function Add-FromUninstallKey {
            param($BaseKey, $SubPath, $Scope, $ViewLabel)
            try {
                $key = $BaseKey.OpenSubKey($SubPath)
            } catch { $key = $null }
            if (-not $key) { return }

            foreach ($name in $key.GetSubKeyNames()) {
                try {
                    $app = $key.OpenSubKey($name)
                    if (-not $app) { continue }

                    $display = $app.GetValue('DisplayName')
                    if ([string]::IsNullOrWhiteSpace($display)) { continue }

                    # Skip child components of a bundle and system components.
                    if ($app.GetValue('SystemComponent') -eq 1)   { continue }
                    if ($app.GetValue('ParentKeyName'))            { continue }
                    if ($app.GetValue('WindowsInstaller') -eq 1 -and
                        $app.GetValue('DisplayName') -eq $null)    { continue }
                    if ($app.GetValue('ReleaseType') -in @('Security Update','Update Rollup','Hotfix')) { continue }

                    $rawDate = [string]$app.GetValue('InstallDate')
                    $instDate = $null
                    if ($rawDate -match '^\d{8}$') {
                        try {
                            $instDate = [datetime]::ParseExact(
                                $rawDate, 'yyyyMMdd',
                                [System.Globalization.CultureInfo]::InvariantCulture
                            ).ToString('yyyy-MM-dd')
                        } catch { $instDate = $rawDate }
                    } elseif ($rawDate) { $instDate = $rawDate }

                    $apps.Add([pscustomobject]@{
                        DisplayName    = ([string]$display).Trim()
                        DisplayVersion = ([string]$app.GetValue('DisplayVersion')).Trim()
                        Publisher      = ([string]$app.GetValue('Publisher')).Trim()
                        InstallDate    = $instDate
                        Architecture   = $ViewLabel
                        Scope          = $Scope
                        UninstallKey   = ([string]$app.GetValue('UninstallString')).Trim()
                    })
                } catch { }
                finally { if ($app) { $app.Close() } }
            }
            $key.Close()
        }

        $paths = @(
            @{ Sub = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall';               View = '64-bit' },
            @{ Sub = 'SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall';   View = '32-bit' }
        )

        # HKLM - machine-wide installs
        $hklm = [Microsoft.Win32.RegistryKey]::OpenBaseKey(
            [Microsoft.Win32.RegistryHive]::LocalMachine,
            [Microsoft.Win32.RegistryView]::Default
        )
        foreach ($p in $paths) { Add-FromUninstallKey $hklm $p.Sub 'Machine' $p.View }
        $hklm.Close()

        # HKU - per-user installs for loaded profiles (opt-in)
        if ($IncludeUser) {
            $hku = [Microsoft.Win32.RegistryKey]::OpenBaseKey(
                [Microsoft.Win32.RegistryHive]::Users,
                [Microsoft.Win32.RegistryView]::Default
            )
            foreach ($sid in $hku.GetSubKeyNames()) {
                if ($sid -notmatch '^S-1-5-21-') { continue }   # real user SIDs only
                foreach ($p in $paths) {
                    Add-FromUninstallKey $hku "$sid\$($p.Sub)" "User ($sid)" $p.View
                }
            }
            $hku.Close()
        }

        return $apps
    }

    try {
        if ($UseWinRM) {
            $icmParams = @{
                ComputerName = $Computer
                ScriptBlock  = $collector
                ArgumentList = @($Computer, $IncludeUserSoftware)
                ErrorAction  = 'Stop'
            }
            if ($Credential) { $icmParams.Credential = $Credential }
            $apps = Invoke-Command @icmParams
        }
        else {
            # Remote Registry transport. Reach across with OpenRemoteBaseKey,
            # but reuse the same parsing logic by running it against remote keys.
            $apps = & {
                $result = New-Object System.Collections.Generic.List[object]

                function Read-Remote {
                    param($Hive, $View, $Computer, $Paths, $ScopePrefix, $UserSids)
                    try {
                        $base = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($Hive, $Computer, $View)
                    } catch { return }
                    foreach ($p in $Paths) {
                        try { $key = $base.OpenSubKey($p.Sub) } catch { $key = $null }
                        if (-not $key) { continue }
                        foreach ($name in $key.GetSubKeyNames()) {
                            try {
                                $app = $key.OpenSubKey($name)
                                if (-not $app) { continue }
                                $display = $app.GetValue('DisplayName')
                                if ([string]::IsNullOrWhiteSpace($display)) { continue }
                                if ($app.GetValue('SystemComponent') -eq 1) { continue }
                                if ($app.GetValue('ParentKeyName'))         { continue }
                                if ($app.GetValue('ReleaseType') -in @('Security Update','Update Rollup','Hotfix')) { continue }

                                $rawDate = [string]$app.GetValue('InstallDate')
                                $instDate = $null
                                if ($rawDate -match '^\d{8}$') {
                                    try {
                                        $instDate = [datetime]::ParseExact($rawDate,'yyyyMMdd',[System.Globalization.CultureInfo]::InvariantCulture).ToString('yyyy-MM-dd')
                                    } catch { $instDate = $rawDate }
                                } elseif ($rawDate) { $instDate = $rawDate }

                                $viewLabel = if ($View -eq [Microsoft.Win32.RegistryView]::Registry32) { '32-bit' } else { '64-bit' }
                                $result.Add([pscustomobject]@{
                                    DisplayName    = ([string]$display).Trim()
                                    DisplayVersion = ([string]$app.GetValue('DisplayVersion')).Trim()
                                    Publisher      = ([string]$app.GetValue('Publisher')).Trim()
                                    InstallDate    = $instDate
                                    Architecture   = $viewLabel
                                    Scope          = $ScopePrefix
                                    UninstallKey   = ([string]$app.GetValue('UninstallString')).Trim()
                                })
                            } catch { } finally { if ($app) { $app.Close() } }
                        }
                        $key.Close()
                    }
                    $base.Close()
                }

                $hklmPaths = @(
                    @{ Sub = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall' }
                )
                # 64-bit view sees native keys; 32-bit view sees WOW6432Node.
                Read-Remote ([Microsoft.Win32.RegistryHive]::LocalMachine) ([Microsoft.Win32.RegistryView]::Registry64) $Computer $hklmPaths 'Machine'
                Read-Remote ([Microsoft.Win32.RegistryHive]::LocalMachine) ([Microsoft.Win32.RegistryView]::Registry32) $Computer $hklmPaths 'Machine'

                if ($IncludeUserSoftware) {
                    try {
                        $hku = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey([Microsoft.Win32.RegistryHive]::Users, $Computer)
                        $sids = $hku.GetSubKeyNames() | Where-Object { $_ -match '^S-1-5-21-' }
                        $hku.Close()
                        foreach ($sid in $sids) {
                            $userPaths = @(
                                @{ Sub = "$sid\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" },
                                @{ Sub = "$sid\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall" }
                            )
                            Read-Remote ([Microsoft.Win32.RegistryHive]::Users) ([Microsoft.Win32.RegistryView]::Default) $Computer $userPaths "User ($sid)"
                        }
                    } catch { }
                }

                $result
            }
        }

        return [pscustomobject]@{
            ComputerName = $Computer
            Status       = 'OK'
            Error        = $null
            Apps         = @($apps)
        }
    }
    catch {
        return [pscustomobject]@{
            ComputerName = $Computer
            Status       = 'Error'
            Error        = $_.Exception.Message
            Apps         = @()
        }
    }
}

# --------------------------------------------------------------- run it all --
$targets = @(Get-TargetComputers)
if (-not $targets -or $targets.Count -eq 0) {
    throw "No target computers found. Check your -SearchBase / -ComputerName / -InputFile."
}
Write-Host ("Targets to scan : {0}" -f $targets.Count) -ForegroundColor Yellow
Write-Host ("Transport       : {0}" -f ($(if ($UseWinRM) { 'WinRM / PS Remoting' } else { 'Remote Registry (RPC)' }))) -ForegroundColor Yellow
Write-Host ("Parallelism     : {0}`n" -f $ThrottleLimit) -ForegroundColor Yellow

# ------------------------------------------------- optional service pre-pass -
# Remember what we changed so -RestoreServiceState can put it back.
$transportSvc = if ($UseWinRM) { 'WinRM' } else { 'RemoteRegistry' }
$restoreMap   = @{}   # computer -> original StartMode of $transportSvc

if ($AutoStartService) {
    Write-Host ("Ensuring '{0}' is running on {1} target(s)..." -f $transportSvc, $targets.Count) -ForegroundColor Yellow
    $i = 0
    foreach ($c in $targets) {
        $i++
        Write-Progress -Activity "Preparing $transportSvc service" -Status $c `
            -PercentComplete (($i / $targets.Count) * 100)

        $state = Get-SIServiceState -ComputerName $c -Name $transportSvc -Credential $Credential
        if (-not $state.Online) { continue }
        if ($state.$transportSvc -eq 'Running') { continue }
        if ($state.$transportSvc -in @('NotFound', 'Error', 'Unknown')) { continue }

        # Disabled services can't start until the start mode is changed first.
        $origMode = $state."${transportSvc}StartMode"
        $set = Set-SIService -ComputerName $c -Name $transportSvc `
            -StartMode 'Manual' -Action 'Start' -Credential $Credential
        if ($set.Success -and $set.State -eq 'Running') {
            $restoreMap[$c] = $origMode
            Write-Verbose "Started $transportSvc on $c (was $origMode)."
        } else {
            Write-Warning ("Could not start {0} on {1}: {2}" -f $transportSvc, $c, $set.Error)
        }
    }
    Write-Progress -Activity "Preparing $transportSvc service" -Completed
    Write-Host ("Started service on {0} machine(s).`n" -f $restoreMap.Count) -ForegroundColor Yellow
}

$machineResults = New-Object System.Collections.Generic.List[object]

# PowerShell 7 has ForEach-Object -Parallel; 5.1 falls back to background jobs.
$isPS7 = $PSVersionTable.PSVersion.Major -ge 7

if ($isPS7) {
    $machineResults = $targets | ForEach-Object -ThrottleLimit $ThrottleLimit -Parallel {
        $w = [scriptblock]::Create($using:worker)
        & $w $_ $using:IncludeMicrosoft $using:IncludeUpdates $using:IncludeUserSoftware $using:UseWinRM $using:Credential
    }
}
else {
    $queue = [System.Collections.Queue]::new(@($targets))
    $running = @{}
    while ($queue.Count -gt 0 -or $running.Count -gt 0) {
        while ($queue.Count -gt 0 -and $running.Count -lt $ThrottleLimit) {
            $c = $queue.Dequeue()
            $job = Start-Job -ScriptBlock $worker -ArgumentList @(
                $c, [bool]$IncludeMicrosoft, [bool]$IncludeUpdates,
                [bool]$IncludeUserSoftware, [bool]$UseWinRM, $Credential
            )
            $running[$job.Id] = $job
        }
        Start-Sleep -Milliseconds 200
        foreach ($job in @($running.Values | Where-Object { $_.State -ne 'Running' })) {
            try { $machineResults.Add((Receive-Job -Job $job -ErrorAction SilentlyContinue)) } catch { }
            Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
            $running.Remove($job.Id)
        }
        $done = $machineResults.Count
        Write-Progress -Activity "Scanning domain computers" `
            -Status "$done of $($targets.Count) complete" `
            -PercentComplete ([math]::Min(100, ($done / $targets.Count) * 100))
    }
    Write-Progress -Activity "Scanning domain computers" -Completed
}

# ------------------------------------------------ restore service state ------
if ($RestoreServiceState -and $restoreMap.Count -gt 0) {
    Write-Host ("`nRestoring '{0}' on {1} machine(s)..." -f $transportSvc, $restoreMap.Count) -ForegroundColor Yellow
    foreach ($c in $restoreMap.Keys) {
        # Win32 StartMode 'Auto' maps to the ChangeStartMode value 'Automatic'.
        $mode = switch ($restoreMap[$c]) {
            'Auto'     { 'Automatic' }
            'Disabled' { 'Disabled' }
            default    { 'Manual' }
        }
        $r = Set-SIService -ComputerName $c -Name $transportSvc `
            -StartMode $mode -Action 'Stop' -Credential $Credential
        if (-not $r.Success) {
            Write-Warning ("Could not restore {0} on {1}: {2}" -f $transportSvc, $c, $r.Error)
        }
    }
}

# ------------------------------------------------------- filter + flatten ----
# Publishers / names we treat as first-party OS components (excluded by default)
$msPublisherPattern = '^(Microsoft|Microsoft Corporation)$'
$updateNamePattern  = '(?i)^(Security Update|Update for|Hotfix|KB\d{6,}|Windows .*(Update|Language Pack)|Definition Update)'

$rows = New-Object System.Collections.Generic.List[object]
foreach ($mr in $machineResults) {
    if (-not $mr) { continue }
    foreach ($a in $mr.Apps) {
        if (-not $IncludeMicrosoft -and $a.Publisher -match $msPublisherPattern) { continue }
        if (-not $IncludeUpdates    -and $a.DisplayName -match $updateNamePattern) { continue }

        $rows.Add([pscustomobject]@{
            ComputerName   = $mr.ComputerName
            DisplayName    = $a.DisplayName
            DisplayVersion = $a.DisplayVersion
            Publisher      = $a.Publisher
            InstallDate    = $a.InstallDate
            Architecture   = $a.Architecture
            Scope          = $a.Scope
        })
    }
}

# De-duplicate (a machine can list the same app in 32- and 64-bit views).
$inventory = $rows |
    Sort-Object ComputerName, DisplayName, DisplayVersion, Scope -Unique

# ------------------------------------------------------------- summaries -----
$okCount      = @($machineResults | Where-Object { $_.Status -eq 'OK' }).Count
$offCount     = @($machineResults | Where-Object { $_.Status -eq 'Offline' }).Count
$errCount     = @($machineResults | Where-Object { $_.Status -eq 'Error' }).Count
$failed       = @($machineResults | Where-Object { $_.Status -ne 'OK' })

Write-Host ""
Write-Host ("Scanned OK      : {0}" -f $okCount)  -ForegroundColor Green
Write-Host ("Offline         : {0}" -f $offCount) -ForegroundColor DarkYellow
Write-Host ("Errors          : {0}" -f $errCount) -ForegroundColor Red
Write-Host ("Software rows   : {0}" -f $inventory.Count) -ForegroundColor Green

# ----------------------------------------------------------------- export ----
$inventory | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding UTF8
Write-Host ("`nCSV report      : {0}" -f $csvPath) -ForegroundColor Cyan

# Build a tidy, self-contained HTML report (no external assets).
$dupSummary = $inventory |
    Group-Object DisplayName, DisplayVersion |
    ForEach-Object {
        $parts = $_.Name -split ', ', 2
        [pscustomobject]@{
            DisplayName    = $parts[0]
            DisplayVersion = if ($parts.Count -gt 1) { $parts[1] } else { '' }
            Installs       = $_.Count
        }
    } |
    Sort-Object Installs -Descending, DisplayName

$style = @'
<style>
  :root { color-scheme: light dark; }
  body { font-family: Segoe UI, system-ui, sans-serif; margin: 24px; background:#f6f8fb; color:#1b2733; }
  h1 { font-size: 22px; margin: 0 0 4px; }
  h2 { font-size: 16px; margin: 28px 0 8px; color:#2b6fd6; }
  .sub { color:#5a6b7b; margin-bottom: 20px; font-size: 13px; }
  .cards { display:flex; gap:12px; flex-wrap:wrap; margin-bottom: 8px; }
  .card { background:#fff; border:1px solid #e2e8f0; border-radius:10px; padding:12px 16px; min-width:120px; }
  .card .n { font-size:22px; font-weight:700; }
  .card .l { font-size:12px; color:#5a6b7b; }
  table { border-collapse: collapse; width:100%; background:#fff; border:1px solid #e2e8f0; border-radius:8px; overflow:hidden; }
  th, td { text-align:left; padding:7px 10px; font-size:13px; border-bottom:1px solid #eef2f6; }
  th { background:#eaf1fb; position:sticky; top:0; cursor:pointer; }
  tr:hover td { background:#f3f7fd; }
  input#q { padding:8px 10px; width:320px; max-width:100%; border:1px solid #cbd5e1; border-radius:8px; margin-bottom:10px; }
  .foot { margin-top:22px; color:#8595a4; font-size:12px; }
  @media (prefers-color-scheme: dark) {
    body{background:#0f141a;color:#dce3ea;} .card,table{background:#161d26;border-color:#263140;}
    th{background:#1c2836;} tr:hover td{background:#1a2431;} .sub,.card .l,.foot{color:#8595a4;}
    input#q{background:#161d26;color:#dce3ea;border-color:#334155;}
  }
</style>
'@

$script = @'
<script>
function filterRows(){
  var q=document.getElementById('q').value.toLowerCase();
  document.querySelectorAll('table#inv tbody tr').forEach(function(r){
    r.style.display = r.innerText.toLowerCase().indexOf(q)>-1 ? '' : 'none';
  });
}
function sortTable(n){
  var t=document.getElementById('inv'),rows=Array.from(t.tBodies[0].rows);
  var asc=t.getAttribute('data-sort')!=String(n)+'a';
  rows.sort(function(a,b){var x=a.cells[n].innerText.toLowerCase(),y=b.cells[n].innerText.toLowerCase();return (x<y?-1:x>y?1:0)*(asc?1:-1);});
  rows.forEach(function(r){t.tBodies[0].appendChild(r);});
  t.setAttribute('data-sort',String(n)+(asc?'a':'d'));
}
</script>
'@

# Needed for HTML-encoding cell values (load before we build any rows).
Add-Type -AssemblyName System.Web -ErrorAction SilentlyContinue

$genInfo = "Generated {0} | Scanned OK: {1} | Offline: {2} | Errors: {3}" -f `
    $startedAt.ToString('yyyy-MM-dd HH:mm'), $okCount, $offCount, $errCount

$cards = @"
<div class="cards">
  <div class="card"><div class="n">$($inventory.Count)</div><div class="l">Software rows</div></div>
  <div class="card"><div class="n">$(@($inventory.DisplayName | Sort-Object -Unique).Count)</div><div class="l">Unique titles</div></div>
  <div class="card"><div class="n">$okCount</div><div class="l">Computers scanned</div></div>
  <div class="card"><div class="n">$offCount</div><div class="l">Offline</div></div>
  <div class="card"><div class="n">$errCount</div><div class="l">Errors</div></div>
</div>
"@

$invRowsHtml = ($inventory | ForEach-Object {
    "<tr><td>$([System.Net.WebUtility]::HtmlEncode($_.ComputerName))</td>" +
    "<td>$([System.Net.WebUtility]::HtmlEncode($_.DisplayName))</td>" +
    "<td>$([System.Net.WebUtility]::HtmlEncode($_.DisplayVersion))</td>" +
    "<td>$([System.Net.WebUtility]::HtmlEncode($_.Publisher))</td>" +
    "<td>$([System.Net.WebUtility]::HtmlEncode([string]$_.InstallDate))</td>" +
    "<td>$([System.Net.WebUtility]::HtmlEncode($_.Architecture))</td>" +
    "<td>$([System.Net.WebUtility]::HtmlEncode($_.Scope))</td></tr>"
}) -join "`n"

$topRowsHtml = ($dupSummary | Select-Object -First 40 | ForEach-Object {
    "<tr><td>$([System.Net.WebUtility]::HtmlEncode($_.DisplayName))</td>" +
    "<td>$([System.Net.WebUtility]::HtmlEncode($_.DisplayVersion))</td>" +
    "<td>$($_.Installs)</td></tr>"
}) -join "`n"

$failRowsHtml = ($failed | ForEach-Object {
    "<tr><td>$([System.Net.WebUtility]::HtmlEncode($_.ComputerName))</td>" +
    "<td>$([System.Net.WebUtility]::HtmlEncode($_.Status))</td>" +
    "<td>$([System.Net.WebUtility]::HtmlEncode([string]$_.Error))</td></tr>"
}) -join "`n"

$html = @"
<!doctype html><html><head><meta charset="utf-8">
<title>Domain Software Inventory - TechyGeeksHome</title>$style$script</head><body>
<h1>Domain Software Inventory</h1>
<div class="sub">$genInfo &middot; TechyGeeksHome &middot; Ultimate Settings Panel</div>
$cards
<h2>Installed third-party software</h2>
<input id="q" placeholder="Filter... (computer, app, publisher)" onkeyup="filterRows()">
<table id="inv" data-sort=""><thead><tr>
<th onclick="sortTable(0)">Computer</th><th onclick="sortTable(1)">Software</th>
<th onclick="sortTable(2)">Version</th><th onclick="sortTable(3)">Publisher</th>
<th onclick="sortTable(4)">Installed</th><th onclick="sortTable(5)">Arch</th>
<th onclick="sortTable(6)">Scope</th></tr></thead><tbody>
$invRowsHtml
</tbody></table>
<h2>Most common titles</h2>
<table><thead><tr><th>Software</th><th>Version</th><th>Installs</th></tr></thead><tbody>
$topRowsHtml
</tbody></table>
$(if ($failed.Count) {
"<h2>Computers not scanned</h2>
<table><thead><tr><th>Computer</th><th>Status</th><th>Detail</th></tr></thead><tbody>
$failRowsHtml
</tbody></table>"})
<div class="foot">Read-only report. No changes were made to any machine.
Made by TechyGeeksHome - https://techygeekshome.info</div>
</body></html>
"@

$html | Out-File -LiteralPath $htmlPath -Encoding UTF8
Write-Host ("HTML report     : {0}" -f $htmlPath) -ForegroundColor Cyan

$elapsed = (Get-Date) - $startedAt
Write-Host ("`nDone in {0:mm}m {0:ss}s." -f $elapsed) -ForegroundColor Green

# Return the inventory objects for further pipeline use.
$inventory
