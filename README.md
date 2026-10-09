# ConfigMgr Task Sequence Monitor

A lightweight, standalone PowerShell and WPF-based administration tool for real-time monitoring, troubleshooting, and reporting on Microsoft Endpoint Configuration Manager (SCCM / MECM) Task Sequence executions.

Originally conceptualized by Trevor Jones and re-architected into a standalone script by Kip Bartholomew, this tool queries the ConfigMgr site database directly to provide granular, step-by-step visibility into ongoing and historical Task Sequence deployments without requiring the Configuration Manager Admin Console.

---

## Table of Contents

- [Overview](#overview)
- [Key Features](#key-features)
- [How It Works](#how-it-works)
  - [Architecture & Concurrency](#architecture--concurrency)
  - [Database Queries & Views](#database-queries--views)
- [Prerequisites](#prerequisites)
- [Installation & Initial Configuration](#installation--initial-configuration)
- [Usage Guide](#usage-guide)
  - [Live Monitoring View](#live-monitoring-view)
  - [Viewing Action Output](#viewing-action-output)
  - [Generating Reports](#generating-reports)
- [Settings & Registry Persistence](#settings--registry-persistence)
- [Troubleshooting & Common Issues](#troubleshooting--common-issues)
- [Credits](#credits)

---

## Overview

Troubleshooting Configuration Manager Task Sequences usually involves digging through `smsts.log` files on target endpoints or waiting for Status Message queries to populate in the console. 

**ConfigMgr Task Sequence Monitor** bridges this gap by connecting directly to the site's SQL Server database. It surfaces live step progress, exit codes, execution durations, and step-level standard output in an intuitive, multi-threaded GUI.

---

## Key Features

- **Zero-Install Standalone Script**: Pure PowerShell script utilizing native WPF (XAML) for modern UI rendering; requires no third-party compiled binaries or modules.
- **Asynchronous Execution (Non-Blocking UI)**: Employs `[RunspaceFactory]::CreateRunspacePool` to execute heavy SQL queries on background threads, ensuring the UI remains responsive at all times.
- **Real-Time Step Monitoring**: Displays step numbers, step names, group names, execution timestamps, and return/exit codes with visual status icons (green check / red cross).
- **Embedded Log & Output Viewer**: Click on any executed step to instantly inspect the raw `ActionOutput` log stream in the lower pane.
- **Filtering & Auto-Refresh**:
  - Filter by Task Sequence deployment.
  - Filter by target computer name (or `-All-`).
  - Filter by timeframe (past *N* hours).
  - Quick toggle to display **Errors Only** (`ExitCode <> 0`).
  - Configurable polling interval with automatic background timer refresh.
- **Automated Reporting Suite**:
  - Export executive summary and complete step audits to **HTML** or multi-file **CSV**.
  - Calculates execution durations (`FinishTime - StartTime`) and pulls computer hardware model information.
- **Time Zone Flexibility**: Toggle between local machine time and UTC timestamps.
- **MDT Web Service Support**: Capable of resolving unknown Bare-Metal computer records via SMBIOS GUID lookups against MDT database web endpoints.

---

## How It Works

### Architecture & Concurrency

```
  +-----------------------------------------------------------+
  |                   WPF UI Thread (STA)                     |
  |  - Controls & Grids (DataGrid, Splitter, TextBoxes)       |
  |  - Win32 Icon Extractor (Shell32.dll -> comres.dll)       |
  +-----------------------------+-----------------------------+
                                |
                 Dispatches Asynchronous Tasks
                                v
  +-----------------------------------------------------------+
  |              PowerShell Runspace Pool (MTA)               |
  |  - Get-TaskSequenceData      - Populate-ComputerNames     |
  |  - Populate-ActionOutput     - Generate-Report            |
  |  - Get-MDTData                                            |
  +-----------------------------+-----------------------------+
                                |
                        SQL Connection (SSPI)
                                v
  +-----------------------------------------------------------+
  |         ConfigMgr Site Database (Microsoft SQL Server)     |
  |  - Views: vSMS_TaskSequenceExecutionStatus, v_R_System,   |
  |           vDeploymentSummary, v_GS_COMPUTER_SYSTEM        |
  +-----------------------------------------------------------+
```

1. **Auto-Discovery via CIM**:
   At startup, the script queries WMI/CIM on the SMS Provider (`ROOT\SMS\site_<SiteCode>`) via class `SMS_SCI_SiteDefinition` to automatically discover the SQL Server instance and site database name.
2. **Dynamic UI Generation**:
   The graphical interface is constructed dynamically by parsing embedded XAML strings via `[Windows.Markup.XamlReader]`.
3. **Runspace Threading**:
   Rather than locking the interface during database transactions, query tasks are wrapped into separate PowerShell script blocks and invoked within an STA Runspace Pool. The results are marshaled back to the UI thread using `$hash.Window.Dispatcher.Invoke()`.
4. **Native Icon P/Invoke**:
   The script dynamically compiles an inline C# type (`System.IconExtractor`) to pull standard system status icons from `comres.dll` into temporary BMP files, which are bound directly to the WPF DataGrid.

### Database Queries & Views

The tool interacts with the following standard Configuration Manager SQL views:

| View Name | Purpose |
|---|---|
| `vDeploymentSummary` | Retrieves all active Task Sequence deployments (`FeatureType = 7`). |
| `vSMS_TaskSequenceExecutionStatus` | Core status view containing execution records, step index, action name, exit codes, and output text. |
| `v_R_System` | Joins Resource IDs to retrieve NetBIOS computer names and SMBIOS GUIDs. |
| `v_RA_System_MACAddresses` | Correlates machine records to network adapter addresses. |
| `v_TaskSequencePackage` | Maps package identifiers to friendly Task Sequence names. |
| `v_GS_COMPUTER_SYSTEM` | Gathers hardware model details for summary reporting. |


## Prerequisites

1. **Operating System**: Windows 10 / 11 or Windows Server 2016+ (64-bit recommended).
2. **PowerShell**: Windows PowerShell 5.1 (STA mode, default on Windows).
3. **Permissions**:
   - **SQL Server Rights**: `db_datareader` access to the Configuration Manager site database (`CM_<SiteCode>`).
   - **Local Administrator** *(Optional)*: Required only if you wish to persist SQL/Database settings into the registry under `HKLM:\SOFTWARE\SMSAgent`.
   - **Remote CIM/WMI Rights**: Read permissions on the SMS Provider namespace if utilizing automatic discovery.

---

## Installation & Initial Configuration

1. Clone or download the repository to your local administrative workstation:
   ```powershell
   git clone https://github.com/Averexs/ConfigMgr-Task-Sequence-Monitor.git
   cd configmgr-ts-monitor
   ```

2. Open the script in an editor or PowerShell ISE/VS Code and configure your environment variables at the top of the file:
   ```powershell
   #==========================================================================
   # GLOBAL CONFIGURATION VARIABLES
   #==========================================================================
   $SiteCode = "XYZ"                            # Replace with your 3-character ConfigMgr Site Code
   $ProviderMachineName = "cm01.corp.domain"    # FQDN of your primary site server or SMS Provider
   ```

3. Launch the script:
   ```powershell
   powershell.exe -ExecutionPolicy Bypass -File .\ConfigMgrTaskSequenceMonitor.ps1
   ```

*(Tip: If run as Administrator, settings saved through the GUI "Settings" dialog will be cached in the Windows Registry for future sessions).*

---

## Usage Guide

### Live Monitoring View

```
+-------------------------------------------------------------------------------------------------------+
| ConfigMgr Task Sequence Monitor                                                                       |
+-------------------------------------------------------------------------------------------------------+
| Task Sequence: [ Windows 11 Enterprise Deployment   v ]  Time Period (Hours): [ 24 ]  Errors Only: [ ] |
| ComputerName:  [ -All-                              v ]  Refresh (Minutes):    [  1 ]  [ Refresh Now ] |
+-------------------------------------------------------------------------------------------------------+
| Icon | ComputerName | ExecutionTime       | Step | ActionName          | GroupName      | ExitCode     |
| (/)  | PC-W11-042   | 2025-05-10 14:12:05 | 12   | Apply OS Image      | Install OS     | 0            |
| (X)  | PC-W11-019   | 2025-05-10 13:45:10 | 38   | Install Application | App Deployment | 1603         |
+-------------------------------------------------------------------------------------------------------+
| Action Output:                                                                                        |
| Command line returned 1603. Fatal error during installation. Exiting step...                          |
+-------------------------------------------------------------------------------------------------------+
```

1. **Select Task Sequence**: Select a Task Sequence from the dropdown. The application immediately pulls the list of matching computer executions.
2. **Filter by Computer**: Choose a specific machine name or retain `-All-` to observe all machines executing that Task Sequence concurrently.
3. **Filter Errors Only**: Check the **Errors Only** box to isolate failed steps across all systems.
4. **Adjust Time Window**: Change the **Time Period (Hours)** field and press `Enter` to expand or narrow the historical query window.

### Viewing Action Output

Select any step row in the data grid. The bottom pane (**Action Output**) will display the full output generated by the task sequence engine for that specific step (identical to what is recorded in `smsts.log`).

### Generating Reports

1. Click the **Generate Report** button (or access the **Summary Report** tab under **Settings**).
2. Choose your target **Task Sequence**, **Start Date**, and **End Date**.
3. Select your desired output format:
   - **HTML Report**: Generates a formatted executive summary report (`TSReport.htm`) containing execution timing metrics (Start, Finish, Duration) followed by complete step logs. Opens automatically in your default browser.
   - **CSV Report**: Exports two CSV files to the `reports\` subfolder:
     - `TSReport_Summary.csv`: High-level per-machine start, finish, duration, and hardware model.
     - `TSReport_AllSteps.csv`: Granular row-by-row step execution data including exit codes and action outputs.

---

## Settings & Registry Persistence

Clicking the **Settings** button opens a secondary dialog allowing you to override runtime parameters:

- **SQL Server & Database**: Override auto-detected SQL connection targets.
- **Display Date/Time**: Toggle between **Local Time** and **UTC**.
- **Registry Caching**: When running elevated, custom values are saved to:
  ```
  HKLM:\SOFTWARE\SMSAgent\ConfigMgr Task Sequence Monitor
  ```
  Saved keys include `SQLServer`, `Database`, `MDTURL`, and `DTFormat`.

---

## Troubleshooting & Common Issues

| Issue | Cause | Solution |
|---|---|---|
| **`[ERROR] Could not connect to SQL Server database`** | Current Windows credentials lack SQL permissions, or the SQL instance name is unreachable. | Ensure TCP/1433 is open between your workstation and SQL Server, and verify your account has `db_datareader` permissions. |
| **DataGrid shows no records** | Time period is too short or deployment summary hasn't populated status messages. | Increase **Time Period (Hours)** to 48 or 72. Ensure Task Sequence Status Message reporting is enabled in site component configuration. |
| **Settings do not persist across restarts** | PowerShell process was launched without administrative elevation. | Right-click your PowerShell console and select **Run as Administrator** prior to updating settings. |

---

## Credits

- **Original Concept & Tool**: Trevor Jones
- **Standalone Script & Modernization**: Kip Bartholomew (Version 2.0 Standalone Script Edition)
