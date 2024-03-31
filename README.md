# WDACTrayTool
I created this WDAC System Tray Tool to facilitate the rapid changing of WDAC policies. Specifically, I wanted a way to quickly switch between Enforced Mode and Audit Mode so that I could review logs and change rules in the policies as necessary. Since this has really helped benefit my application allowlisting journey, I wanted to share it so that others could also benefit. The tray tool itself might not be great, but the concept itself could be replicated in a tray tool in a better programming language.

### Screenshot:
![WDACTrayTool](https://github.com/WildByDesign/WDACTrayTool/assets/26308319/e189baf5-8bd4-4e40-b59a-bc16906e792a)


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

### Compiling:

To compile the script, you need to use SciTE4AutoIt3 which is available here: https://www.autoitscript.com/site/autoit-script-editor/downloads/
