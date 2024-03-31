# WDAC Tray Tool
I created this WDAC System Tray Tool to facilitate the rapid changing of WDAC policies. Specifically, I wanted a way to quickly switch between Enforced Mode and Audit Mode so that I could review logs and change rules in the policies as necessary. Since this has really helped benefit my application allowlisting journey, I wanted to share it so that others could also benefit. 

The tray tool itself might not be great (due to using AutoIT), but the concept itself could be replicated in a tray tool in a better programming language. If somebody could recreate this concept in another programming language and even add toast notifications via `Microsoft-Windows-CodeIntegrity` provider, that would be absolutely phenomenal. 

### Screenshot:
![WDACTrayTool](https://github.com/WildByDesign/WDACTrayTool/assets/26308319/e189baf5-8bd4-4e40-b59a-bc16906e792a)

### Policy Type:

At the moment, this tray tool only supports Multiple Policy Format since that is what I have always used since inception. Although at some point it could be extended to support Single Policy Format as well.

### Concept & Methods Used:



### Directory Structure:
```batch
 WDACTrayTool.exe
│
└───policies
    ├───AllowAllMode
    │       AllowAll.xml
    │       {BD0E4FC3-D24E-43E2-BEA9-8F4C4B7165EE}.cip
    │
    ├───AuditMode
    │       WDAC-Audit.xml
    │       {BD0E4FC3-D24E-43E2-BEA9-8F4C4B7165EE}.cip
    │       {C0ECA62D-F88F-48E4-9DBD-7923DA0DA774}.cip
    │
    ├───EnforcedMode
    │       WDAC-Enforced.xml
    │       {BD0E4FC3-D24E-43E2-BEA9-8F4C4B7165EE}.cip
    │       {C0ECA62D-F88F-48E4-9DBD-7923DA0DA774}.cip
    │
    └───RefreshPolicy
            RefreshPolicy.exe
```

### Policy Refresh:

The policy refresh within this tray tool is achieved via: https://www.microsoft.com/en-us/download/details.aspx?id=102925

Originally, I was using `CiTool.exe --refresh` to accomplish to policy refresh. However, that would cause the tray tool to hang because `CiTool.exe --refresh` would wait for user input to press Enter.


### Compiling:

To compile the script, you need to use SciTE4AutoIt3 which is available here: https://www.autoitscript.com/site/autoit-script-editor/downloads/
