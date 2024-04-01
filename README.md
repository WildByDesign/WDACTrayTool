# WDAC Tray Tool
I created this WDAC System Tray Tool to facilitate the rapid changing of WDAC policies. Specifically, I wanted a way to quickly switch between Enforced Mode and Audit Mode so that I could review logs and change rules in the policies as necessary. Since this has really helped benefit my application allowlisting journey, I wanted to share it so that others could also benefit. 

The tray tool itself might not be great (due to using AutoIT), but the concept itself could be replicated in a tray tool in a better programming language. If somebody could recreate this concept in another programming language and even add toast notifications via `Microsoft-Windows-CodeIntegrity` provider, that would be absolutely phenomenal. 

### Screenshot:
![WDACTrayTool](https://github.com/WildByDesign/WDACTrayTool/assets/26308319/e189baf5-8bd4-4e40-b59a-bc16906e792a)

### Policy Type:

At the moment, this tray tool only supports Multiple Policy Format since that is what I have always used since inception. Although at some point it could be extended to support Single Policy Format as well.

### Concept & Methods Used:

The concept (at the present time) is really quite simple. You need to have 3 policies; AllowAllMode, AuditMode and EnforcedMode.

I did not want to carelessly delete all existing policies from users' machines. That is why I created it so that all policies
have the same filename (therefore, sharing the same PolicyID in the XML files prior to conversion to binary).

For example, each base policy XML file shares the same PolicyID and BasePolicyID as follows:

```xml
  <PolicyID>{BD0E4FC3-D24E-43E2-BEA9-8F4C4B7165EE}</PolicyID>
  <BasePolicyID>{BD0E4FC3-D24E-43E2-BEA9-8F4C4B7165EE}</BasePolicyID>
```

Your policy binary files need to be placed in the corresponding policy directories:
```batch
.\policies\AllowAllMode
.\policies\AuditMode
.\policies\EnforcedMode
```

The tray tool simply copies the converted policy binary files (*.cip) to `C:\Windows\System32\CodeIntegrity\CiPolicies\Active\`, overwriting policy
files of the same filename and refreshing the policy.

The overall concept here is really quite simplistic. Yet the results of using the tray tool itself is incredibly useful once set up.

Obviously, this concept can be improved upon in many, many ways to allow for more customization around policy switching.

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
    └───EnforcedMode
            WDAC-Enforced.xml
            {BD0E4FC3-D24E-43E2-BEA9-8F4C4B7165EE}.cip
            {C0ECA62D-F88F-48E4-9DBD-7923DA0DA774}.cip
```

### Policy Refresh:

The policy refresh within this tray tool is achieved via: https://www.microsoft.com/en-us/download/details.aspx?id=102925

Originally, I was using `CiTool.exe --refresh` to accomplish to policy refresh. However, that would cause the tray tool to hang because `CiTool.exe --refresh` would wait for user input to press Enter.


### Compiling:

To compile the script, you need to use SciTE4AutoIt3 which is available here: https://www.autoitscript.com/site/autoit-script-editor/downloads/
