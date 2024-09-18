
# Set variables for files in user Temp directory
#
$tempdir = $env:TEMP
$CiPolicy = "$tempdir\CiPolicy.txt"
$CiPolicy1 = "$tempdir\CiPolicy1.txt"
$CiPolicy2 = "$tempdir\CiPolicy2.txt"

# Get status for CodeIntegrityPolicyEnforcementStatus and UsermodeCodeIntegrityPolicyEnforcementStatus; save to CiPolicy1.txt
#
Get-CimInstance -ClassName Win32_DeviceGuard -Namespace root\Microsoft\Windows\DeviceGuard | FL *codeintegrity* | Out-File -FilePath $CiPolicy1

# Remove blank lines from CiPolicy1.txt
#
(Get-Content $CiPolicy1) | Where-Object {$_.trim() -ne "" } | Set-Content $CiPolicy1

# String replacements in the CiPolicy1.txt file
#
$string1 = 'CodeIntegrityPolicyEnforcementStatus         :'
$string2 = 'UsermodeCodeIntegrityPolicyEnforcementStatus :'
$string3 = ' App Control policy              :'
$string4 = ' App Control user mode policy    :'
$string5 = '2'
$string6 = 'Enforced Mode '
$string7 = '1'
$string8 = 'Audit Mode '
$content1 = Get-Content $CiPolicy1
$content1 -replace $string1, $string3 -replace $string2, $string4 -replace $string5, $string6 -replace $string7, $string8 | Set-Content $CiPolicy1

# Add extra lines and wording to CiPolicy1.txt to make final output pretty
#
Add-Content -Path $CiPolicy1 -value ""
Add-Content -Path $CiPolicy1 -value ""
Add-Content -Path $CiPolicy1 -value "Currently Active Policies:"
Add-Content -Path $CiPolicy1 -value ""

# Obtain list of Active Policies and save to CiPolicy2.txt
#
(CiTool -lp -json | ConvertFrom-Json).Policies | Where-Object {$_.IsEnforced -eq 'True'} | Select-Object -Property FriendlyName | Out-File -FilePath $CiPolicy2

# String replacements in the CiPolicy2.txt file
#
$string10 = 'FriendlyName'
$string11 = ''
$string12 = '------------'
$string13 = ''
$content2 = Get-Content $CiPolicy2
$content2 -replace $string10, $string11 -replace $string12, $string13 | Set-Content $CiPolicy2

# Remove blank lines from CiPolicy2.txt
#
(Get-Content $CiPolicy2) | Where-Object {$_.trim() -ne "" } | Set-Content $CiPolicy2

# Add one extra line at bottom of CiPolicy2.txt to make final output pretty
#
Add-Content -Path $CiPolicy2 -value ""

# Merge CiPolicy1.txt and CiPolicy2.txt together and save as CiPolicy.txt
# CiPolicy.txt is later read by tray tool and viewed in a msgbox
#
Get-Content $CiPolicy1, $CiPolicy2 | Set-Content $CiPolicy
