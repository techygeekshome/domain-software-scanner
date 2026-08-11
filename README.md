# Domain Software Scanner

Scans **every computer in an Active Directory domain** and produces a list of
all installed **third-party software with its version** — the same data you see
in *Apps & features* / *Programs and Features* — as a filterable HTML report
and a CSV. Use it from a **GUI** or from the **command line**.

> Part of the **TechyGeeksHome — Ultimate Settings Panel** toolkit.
> Read-only for your software: it only *reads* registry values. The only thing
> it can *change* is starting/stopping the RemoteRegistry or WinRM service, and
> only when you ask it to (and it can put that back afterwards).

## The files

| File | What it is |
| --- | --- |
| **`Show-SoftwareInventoryGui.ps1`** | Point-and-click GUI *(start here)*. Discover machines, see/start services, scan, open the report. |
| **`Get-DomainSoftwareInventory.ps1`** | The command-line engine. Use directly for scheduled tasks / automation. |
| **`SoftwareInventory.Services.psm1`** | Shared helper module that checks and starts the RemoteRegistry / WinRM services remotely. |

Keep all three in the same folder.

## "What if Remote Registry or WinRM isn't running?"

You don't need them running to start them. Starting a Windows service on a
remote machine goes through the **Service Control Manager over RPC/DCOM** (the
same channel WMI uses), which is *independent* of the RemoteRegistry and WinRM
services — so the tool can turn them on for you:

- **GUI:** tick the machines and click **Start Remote Registry** (or **Start
  WinRM**).
- **CLI:** add **`-AutoStartService`**. Pair it with **`-RestoreServiceState`**
  to stop the service again afterwards and leave the estate exactly as found.

Requirements to start a service remotely: local-admin rights on the target and
the *Windows Management Instrumentation (WMI-In)* firewall rule allowed (on by
default in most domains). If even WMI/RPC is blocked, enable RemoteRegistry
fleet-wide via **Group Policy** instead.

> **WinRM caveat:** starting the *service* is easy, but full PowerShell
> Remoting also needs a listener + firewall rule (what `Enable-PSRemoting`
> configures). If PS Remoting was never set up, starting the service alone
> isn't enough — do that via Group Policy. This is why **Remote Registry is the
> default transport**: enabling it is just "start the service".

---

## What it does

- Queries Active Directory for computer accounts (or takes an explicit list).
- Reads each machine's `Uninstall` registry keys remotely — **both** the
  64-bit and 32-bit (`WOW6432Node`) views, plus optional per-user installs.
- Filters out Microsoft/Windows OS components and Windows Updates by default,
  so you get genuine **third-party** software (switches let you include them).
- Writes a timestamped **CSV** and a self-contained **HTML** report
  (sortable, searchable, light/dark), and returns the objects on the pipeline.

It deliberately **does not** use the `Win32_Product` WMI class: querying that
class forces an MSI *reconfiguration* of every installed package on the target
and is painfully slow. Reading the registry is fast and side-effect free.

## Requirements

| Need | Detail |
| --- | --- |
| **Rights** | Run as a domain account with **local admin** on the targets (needed to read `HKLM` remotely). |
| **AD lookup** | The **RSAT ActiveDirectory** PowerShell module — only when you let it query AD. Not needed with `-ComputerName` / `-InputFile`. |
| **Transport** | *Either* the **Remote Registry** service reachable over RPC (default), *or* **WinRM / PS Remoting** with `-UseWinRM`. |
| **PowerShell** | Windows PowerShell 5.1 or PowerShell 7+ (7 scans in parallel faster). |

> **Tip:** the Remote Registry service is often stopped by default. Let the tool
> start it for you — **Start Remote Registry** in the GUI, or `-AutoStartService`
> on the command line (see below) — start it fleet-wide via Group Policy, or use
> `-UseWinRM` if PS Remoting is already enabled in your domain.

## Usage — GUI

Run from a domain-joined machine, elevated, under STA:

```powershell
powershell.exe -STA -ExecutionPolicy Bypass -File .\Show-SoftwareInventoryGui.ps1
```

Then: **①** pick a target source (all AD / an OU / a file / a typed list) and
click **Load / Discover** → **②** tick the machines and optionally **Check
services** / **Start Remote Registry** → **③** set options and click **Scan
selected computers** → open the **HTML** or **CSV** report. The grid shows each
machine's reachability and RemoteRegistry / WinRM state at a glance.

## Usage — command line

Run it from a domain-joined machine (a management server or a DC) in an
elevated PowerShell prompt.

```powershell
# Whole domain, third-party software only -> CSV + HTML next to the script
.\Get-DomainSoftwareInventory.ps1

# Just one OU, and include Microsoft software too
.\Get-DomainSoftwareInventory.ps1 -SearchBase "OU=Servers,DC=contoso,DC=com" -IncludeMicrosoft

# A named subset of machines, over WinRM, with progress
.\Get-DomainSoftwareInventory.ps1 -ComputerName PC01,PC02,SRV05 -UseWinRM -Verbose

# From a text file (one computer per line; # comments allowed)
.\Get-DomainSoftwareInventory.ps1 -InputFile .\computers.txt

# Start Remote Registry where it's stopped, scan, then put it back
.\Get-DomainSoftwareInventory.ps1 -AutoStartService -RestoreServiceState

# Alternate admin credentials and a custom output folder
.\Get-DomainSoftwareInventory.ps1 -Credential (Get-Credential) -OutputFolder C:\Reports
```

If PowerShell blocks the script, unblock it for the current session:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

## Key parameters

| Parameter | Purpose |
| --- | --- |
| `-SearchBase` | Limit the AD search to an OU distinguished name. |
| `-ComputerName` | Scan an explicit list instead of querying AD. |
| `-InputFile` | Read computer names from a text file (one per line). |
| `-IncludeMicrosoft` | Include Microsoft-published software (off by default). |
| `-IncludeUpdates` | Include Windows/security updates & hotfixes (off by default). |
| `-IncludeUserSoftware` | Also list per-user installs from loaded profiles. |
| `-UseWinRM` | Read the registry over PS Remoting instead of Remote Registry. |
| `-AutoStartService` | Start the transport service (RemoteRegistry, or WinRM with `-UseWinRM`) where it's stopped, before scanning. |
| `-RestoreServiceState` | After scanning, stop any service this run started and restore its original start mode. |
| `-ThrottleLimit` | Machines scanned in parallel (default 32). |
| `-OutputFolder` | Where to write the reports (default: the script folder). |
| `-Credential` | Alternate admin credentials for the remote reads. |

Full help is built in:

```powershell
Get-Help .\Get-DomainSoftwareInventory.ps1 -Full
```

## Output

Two timestamped files, e.g.:

```
SoftwareInventory_2026-08-11_141530.csv
SoftwareInventory_2026-08-11_141530.html
```

Columns: **Computer, Software, Version, Publisher, Installed, Architecture,
Scope**. The HTML report also shows headline counts, the most common titles
across the estate, and any machines that couldn't be scanned (offline / error)
with the reason.

## Notes & troubleshooting

- **"RPC server is unavailable"** — the target is off, blocked by firewall, or
  the Remote Registry service is stopped. Allow *Remote Event Log Management*
  / *Windows Management Instrumentation* through the firewall, start Remote
  Registry, or switch to `-UseWinRM`.
- **Access denied** — the account running the script isn't a local admin on
  the target. Use `-Credential` with an account that is.
- **Per-user apps look incomplete** — only profiles whose `HKU` hive is
  currently loaded (logged-on users) are visible; that's a Windows limitation,
  not a bug.

---

Made by **TechyGeeksHome** · <https://techygeekshome.info>
