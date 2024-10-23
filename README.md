# App Control Tray Tool

I created this App Control System Tray Tool to facilitate more efficient changing of App Control policies. Specifically, I wanted a way to quickly switch between Enforced Mode and Audit Mode so that I could review logs and change rules in the policies as necessary. Since this has really helped benefit my application allowlisting journey, I wanted to share it so that others could also benefit. 

### Screenshots:

![WDAC3screen](https://github.com/user-attachments/assets/474bf0c3-f8e0-4b05-aee5-c3179d613dcc)

![AppControlStatus](https://github.com/user-attachments/assets/3a68d2b6-45a7-42ef-a334-64ad5c5d7544)

![wdactray3-blocked](https://github.com/user-attachments/assets/ce6f04dd-0dc9-443b-8a92-2ad825670b64)

![wdactray3-audit](https://github.com/user-attachments/assets/55cf14b9-707c-40b0-94c8-b0f95d01c71d)

![wdactray3-refresh](https://github.com/user-attachments/assets/2690a8bf-2a20-4a75-bbb3-bec39526443e)





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

This is implemented now with the simple Enable Notifications option now on the system tray menu to enable/disable toast notifications.
