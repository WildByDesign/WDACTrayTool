# WDAC Tray Tool
I created this WDAC System Tray Tool to facilitate the rapid changing of WDAC policies. Specifically, I wanted a way to quickly switch between Enforced Mode and Audit Mode so that I could review logs and change rules in the policies as necessary. Since this has really helped benefit my application allowlisting journey, I wanted to share it so that others could also benefit. 

The tray tool itself might not be great (due to using AutoIT), but the concept itself could be replicated in a tray tool in a better programming language. If somebody could recreate this concept in another programming language and even add toast notifications via `Microsoft-Windows-CodeIntegrity` provider, that would be absolutely phenomenal. 

### Screenshot:
![WDACTrayTool-2 0](https://github.com/user-attachments/assets/240aca6e-c1c1-467b-898e-e16325f6a632)


![WDACTrayTool-blocked](https://github.com/WildByDesign/WDACTrayTool/assets/26308319/0d1b71c9-8dc0-495b-a4cb-cb7bbf48b2b5)

![WDACTrayTool-refresh](https://github.com/WildByDesign/WDACTrayTool/assets/26308319/6ae7dfb9-8832-4fa4-b0aa-4d478609800a)



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


### Compiling:

To compile the script, you need to use SciTE4AutoIt3 which is available here: https://www.autoitscript.com/site/autoit-script-editor/downloads/


### Testing:

The example policies included in this are just for testing purposes and should not be used other than for testing.
The policies basically allow for everything to run. There is one Deny rule for the purpose of testing this tray tool
which is `*\test\speedyfox.exe` so that you can test the tray tool going from Audit Mode to Enforced Mode and vice versa.


### Toast Notifications:

I still need to finish this section on registering the scheduled tasks.

Essentially:

Toasts require running `toasts\Install-Tasks.ps1` to register scheduled tasks _(tasks based upon current working directory)_
