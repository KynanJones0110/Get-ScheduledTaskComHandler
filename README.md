# Get-ScheduledTaskComHandler
This PowerShell tool is designed for system auditing and forensic analysis. It identifies Windows Scheduled Tasks that utilize COM (Component Object Model) Handlers rather than standard executable paths, and resolves those handlers to the specific DLL files they execute.

In Windows, many scheduled tasks do not point directly to an .exe. Instead, they use a CLSID (Class Identifier). When the task runs, the Task Scheduler looks up this CLSID in the Windows Registry to find the associated DLL to load.

From a security perspective, this is a common area for persistence mechanisms, as it's less obvious than a standard command-line argument. This script provides full visibility into these tasks.

**Key Features**
- Recursive Scanning: Scans the entire $ENV:windir\System32\Tasks directory for task definitions.

- XML Parsing: Deep-dives into the Task XML to extract ComHandler Class IDs.

- Registry Resolution: Automatically maps the CLSID to its InprocServer32 registry key to find the actual DLL path.

- Dual Output: * GUI: Opens a searchable, sortable window via Out-GridView.

- File: Exports all findings to a ComHandlers.csv for documentation or further analysis.


**Usage**
- Open PowerShell as Administrator (required to read the Tasks directory and Registry keys).

- Copy and paste the function into your session or run the .ps1 script.

- The results will pop up in a separate window and save a CSV to your current directory.

**PowerShell**

**# Example of the data returned:**
- TaskName : ManifestSvc
- CLSID    : {AD652000-029B-4C04-95C4-82D98D55283E}
- DllPath  : C:\Windows\system32\manifsvc.dll
- 
**Use Cases**
- Security Auditing: Detecting unauthorized or suspicious DLLs registered as scheduled tasks.
- System Troubleshooting: Identifying which component is being triggered by a specific system task.
- Forensics: Investigating persistence techniques used by malware or PUPs (Potentially Unwanted Programs).
- Red Team Persistence
