#Requires -Version 5.1
<#
.SYNOPSIS
    Remote service helpers for the Domain Software Inventory tools.

.DESCRIPTION
    Checks and controls the RemoteRegistry and WinRM services on remote
    computers WITHOUT needing either of them to be running first.

    The trick: starting a Windows service on a remote machine goes through the
    Service Control Manager over RPC/DCOM (the same channel WMI uses -
    endpoint mapper on TCP 135 plus dynamic RPC / SMB 445). That channel is
    independent of the RemoteRegistry and WinRM services, so we can use it to
    turn those services ON. Requirements on the target:

        * You have local-admin rights (or pass -Credential for an account that
          does).
        * The "Windows Management Instrumentation (WMI-In)" firewall rule is
          allowed (on by default inside most domains).

    If even WMI/RPC is blocked, nothing remote will work and you should enable
    RemoteRegistry fleet-wide via Group Policy instead.

    Exported functions:
        Test-SIPort         - fast TCP reachability check (default: RPC 135)
        Get-SIServiceState  - report Running/Stopped + StartMode for services
        Set-SIService       - change start mode and/or start/stop a service

    All service calls use a DCOM CIM session so they work on both Windows
    PowerShell 5.1 and PowerShell 7 (Get-Service -ComputerName was removed in 7).
#>

function Test-SIPort {
    <#
    .SYNOPSIS  Quick TCP connect test (used as a reachability gate).
    .DESCRIPTION
        ICMP ping is often blocked while RPC is open, so we test the RPC
        endpoint-mapper port (135) by default instead of pinging.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ComputerName,
        [int]$Port = 135,
        [int]$TimeoutMs = 1500
    )
    $client = [System.Net.Sockets.TcpClient]::new()
    try {
        $iar = $client.BeginConnect($ComputerName, $Port, $null, $null)
        if ($iar.AsyncWaitHandle.WaitOne($TimeoutMs, $false) -and $client.Connected) {
            $client.EndConnect($iar)
            return $true
        }
        return $false
    } catch {
        return $false
    } finally {
        $client.Close()
    }
}

function New-SIDcomSession {
    # Internal: open a DCOM CIM session (no WinRM dependency).
    param([string]$ComputerName, [System.Management.Automation.PSCredential]$Credential)
    $opt = New-CimSessionOption -Protocol Dcom
    $params = @{ ComputerName = $ComputerName; SessionOption = $opt; ErrorAction = 'Stop'; OperationTimeoutSec = 30 }
    if ($Credential) { $params.Credential = $Credential }
    New-CimSession @params
}

function Get-SIServiceState {
    <#
    .SYNOPSIS  Report the state and start mode of services on a remote machine.
    .OUTPUTS
        A PSCustomObject with: ComputerName, Online, Error, and for each
        requested service two properties, e.g. RemoteRegistry ('Running' /
        'Stopped' / 'NotFound' / 'Error') and RemoteRegistryStartMode
        ('Auto' / 'Manual' / 'Disabled').
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ComputerName,
        [string[]]$Name = @('RemoteRegistry', 'WinRM'),
        [System.Management.Automation.PSCredential]$Credential
    )

    $result = [ordered]@{ ComputerName = $ComputerName; Online = $false; Error = $null }
    foreach ($n in $Name) { $result[$n] = 'Unknown'; $result["${n}StartMode"] = '' }

    if (-not (Test-SIPort -ComputerName $ComputerName -Port 135)) {
        $result.Error = 'Unreachable (RPC 135)'
        return [pscustomobject]$result
    }
    $result.Online = $true

    $session = $null
    try {
        $session = New-SIDcomSession -ComputerName $ComputerName -Credential $Credential
        foreach ($n in $Name) {
            try {
                $svc = Get-CimInstance -CimSession $session -ClassName Win32_Service `
                    -Filter "Name='$n'" -ErrorAction Stop
                if ($svc) {
                    $result[$n] = $svc.State
                    $result["${n}StartMode"] = $svc.StartMode
                } else {
                    $result[$n] = 'NotFound'
                }
            } catch {
                $result[$n] = 'Error'
            }
        }
    } catch {
        $result.Error = $_.Exception.Message
    } finally {
        if ($session) { Remove-CimSession $session -ErrorAction SilentlyContinue }
    }
    return [pscustomobject]$result
}

function Set-SIService {
    <#
    .SYNOPSIS  Change a service's start mode and/or start/stop it remotely.
    .DESCRIPTION
        Uses Win32_Service.ChangeStartMode / StartService / StopService over
        DCOM. Returns an object describing the outcome and the resulting state.
    .PARAMETER Name
        RemoteRegistry or WinRM.
    .PARAMETER StartMode
        Optional new start mode: Automatic, Manual or Disabled.
    .PARAMETER Action
        Start, Stop or None (default). Start is a no-op if already running.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$ComputerName,
        [Parameter(Mandatory)][ValidateSet('RemoteRegistry', 'WinRM')][string]$Name,
        [ValidateSet('Automatic', 'Manual', 'Disabled')][string]$StartMode,
        [ValidateSet('Start', 'Stop', 'None')][string]$Action = 'None',
        [System.Management.Automation.PSCredential]$Credential
    )

    $out = [pscustomobject]@{
        ComputerName = $ComputerName
        Name         = $Name
        Success      = $false
        State        = $null
        StartMode    = $null
        Error        = $null
    }

    if (-not (Test-SIPort -ComputerName $ComputerName -Port 135)) {
        $out.Error = 'Unreachable (RPC 135)'
        return $out
    }

    $session = $null
    try {
        $session = New-SIDcomSession -ComputerName $ComputerName -Credential $Credential
        $svc = Get-CimInstance -CimSession $session -ClassName Win32_Service -Filter "Name='$Name'" -ErrorAction Stop
        if (-not $svc) { throw "Service '$Name' not found on $ComputerName." }

        if ($StartMode -and $PSCmdlet.ShouldProcess($ComputerName, "Set $Name start mode to $StartMode")) {
            $r = Invoke-CimMethod -InputObject $svc -MethodName ChangeStartMode `
                -Arguments @{ StartMode = $StartMode } -ErrorAction Stop
            if ($r.ReturnValue -ne 0) { throw "ChangeStartMode failed (code $($r.ReturnValue))." }
        }

        if ($Action -eq 'Start' -and $PSCmdlet.ShouldProcess($ComputerName, "Start $Name")) {
            $svc = Get-CimInstance -CimSession $session -ClassName Win32_Service -Filter "Name='$Name'"
            if ($svc.State -ne 'Running') {
                $r = Invoke-CimMethod -InputObject $svc -MethodName StartService -ErrorAction Stop
                if ($r.ReturnValue -ne 0) { throw "StartService failed (code $($r.ReturnValue))." }
            }
        }

        if ($Action -eq 'Stop' -and $PSCmdlet.ShouldProcess($ComputerName, "Stop $Name")) {
            $svc = Get-CimInstance -CimSession $session -ClassName Win32_Service -Filter "Name='$Name'"
            if ($svc.State -eq 'Running') {
                $r = Invoke-CimMethod -InputObject $svc -MethodName StopService -ErrorAction Stop
                if ($r.ReturnValue -ne 0) { throw "StopService failed (code $($r.ReturnValue))." }
            }
        }

        Start-Sleep -Milliseconds 400
        $svc = Get-CimInstance -CimSession $session -ClassName Win32_Service -Filter "Name='$Name'"
        $out.State     = $svc.State
        $out.StartMode = $svc.StartMode
        $out.Success   = $true
    } catch {
        $out.Error = $_.Exception.Message
    } finally {
        if ($session) { Remove-CimSession $session -ErrorAction SilentlyContinue }
    }
    return $out
}

Export-ModuleMember -Function Test-SIPort, Get-SIServiceState, Set-SIService
