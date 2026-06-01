# KB Installation Checker (GUI Utility)

An interactive graphical user interface (GUI) built in PowerShell designed to audit and verify the installation status of specific Microsoft Knowledge Base (KB) security updates across multiple remote Windows Servers simultaneously.

---

## ⚙️ Features

* **Multi-Server Targeting:** Input a list of servers via a multiline text field.
* **Real-time Progress Tracker:** Active progress bar and grid updates as servers are evaluated.
* **Operating System Reporting:** Automatically queries and reports the specific OS Caption (e.g., *Windows Server 2022 Datacenter*).
* **Pending Restart Identification:** Identifies if a KB is installed but lacks a valid date stamp, suggesting a pending system reboot.
* **Dual Logging Methods:** * **Automated Background Logging:** Quietly writes session records to `C:\Script\KB_check\Logs\` upon completion.
    * **On-Demand Export:** Dedicated button to manually save clean grid results to a custom CSV path via standard File Explorer dialog.

---

## 🚀 Usage Instructions

### Prerequisites
1.  **Administrative Rights:** Run your PowerShell console as an **Administrator**.
2.  **Network Permissions:** The execution account requires standard remote CIM/WMI access permissions on the targeted remote endpoints.
3.  **WinRM / CIM Enabled:** Remote management must be configured (`Enable-PSRemoting`).

### Execution
Simply navigate to the script location and run it:
```powershell
.\Single_KB_Check.ps1
