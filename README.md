# App Control Tray & Policy Manager

I created this App Control System Tray Tool to facilitate more efficient changing of App Control policies. Specifically, I wanted a way to quickly switch between Enforced Mode and Audit Mode so that I could review logs and change rules in the policies as necessary. Since this has really helped benefit my application allowlisting journey, I wanted to share it so that others could also benefit. 

### Requirements:

- Windows 11 (22H2/23H2/24H2 or newer) required for usage of CiTool
- PowerShell 7 or Windows PowerShell


### Screenshots:

![AppControlPolicyManager-5 1](https://github.com/user-attachments/assets/d6f05cdb-19d8-4bde-b992-efc4b54d1705)

![WDAC3screen](https://github.com/user-attachments/assets/e3294dd1-3eb1-4b38-8899-7c935303b7b9)


### Main Files:

`AppControlTray.exe` - System tray tool which run **unelevated** at all times.

`AppControlManager.exe` - App Control for Business GUI

`AppControlHelper.exe` - Command line-only tool which runs only specific **elevated** commands from AppControlTray related to CiTool commands and Event Viewer.

`AppControlTask.exe` - Command line-only tool which runs only specific **unelevated** commands from AppControlTray related to Scheduled Tasks, Toast Notifications and policy conversion.

