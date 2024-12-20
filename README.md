# App Control Tray Tool

I created this App Control System Tray Tool to facilitate more efficient changing of App Control policies. Specifically, I wanted a way to quickly switch between Enforced Mode and Audit Mode so that I could review logs and change rules in the policies as necessary. Since this has really helped benefit my application allowlisting journey, I wanted to share it so that others could also benefit. 

### Requirements:

- Windows 11 (22H2/23H2/24H2 or newer) required for usage of CiTool
- PowerShell 7 or Windows PowerShell


### Screenshots:

![AppControlPolicyManager-5 1](https://github.com/user-attachments/assets/d6f05cdb-19d8-4bde-b992-efc4b54d1705)

![WDAC3screen](https://github.com/user-attachments/assets/e3294dd1-3eb1-4b38-8899-7c935303b7b9)

![TrayStatus](https://github.com/user-attachments/assets/2ddac0cc-cfb0-4c5c-a30b-23f0be3e7d14)

![wdactray3-blocked](https://github.com/user-attachments/assets/ce6f04dd-0dc9-443b-8a92-2ad825670b64)

![wdactray3-audit](https://github.com/user-attachments/assets/55cf14b9-707c-40b0-94c8-b0f95d01c71d)

![wdactray3-refresh](https://github.com/user-attachments/assets/2690a8bf-2a20-4a75-bbb3-bec39526443e)


### Main Files:

`AppControlTray.exe` - System tray tool which run **unelevated** at all times.

`AppControlHelper.exe` - Command line-only tool which runs only specific **elevated** commands from AppControlTray related to CiTool commands and Event Viewer.

`AppControlTask.exe` - Command line-only tool which runs only specific **unelevated** commands from AppControlTray related to Scheduled Tasks, Toast Notifications and policy conversion.

### Policy Type:

At the moment, this tray tool only supports Multiple Policy Format since that is what I have always used since inception. Although at some point it could be extended to support Single Policy Format as well.

### Usage:

This tray tool makes use of compiled policy binaries (*.cip) that you would ideally already have. There are some included just for simple testing purposes.

To add new policies or update existing policies, simply select the tray menu option `Add or Update Policies`. This will bring up a standard file selection dialog which you can use to select any number of policy files. The selection will be parsed and those policies will be applied immediately via `CiTool -up` for each policy selected.

To remove policies, select the tray menu option `Remove Policies`. You can select as many policies for removal as you want. Those selections will be parsed and the policies will be removed immediately via `CiTool -rp` for each policy selected.


### Compiling:

To compile the script, you need to use SciTE4AutoIt3 which is available here: https://www.autoitscript.com/site/autoit-script-editor/downloads/


### Testing:

The example policies included in this are just for testing purposes and should not be used other than for testing.
The policies basically allow for everything to run. There is one Deny rule for the purpose of testing this tray tool
which is `*\test\speedyfox.exe` so that you can test the tray tool going from Audit Mode to Enforced Mode and vice versa.


### Toast Notifications:

This is implemented now with the simple Enable Notifications option now on the system tray menu to enable/disable toast notifications.

Toast notifications are implemented using KDE's Snoretoast app:
https://invent.kde.org/libraries/snoretoast
