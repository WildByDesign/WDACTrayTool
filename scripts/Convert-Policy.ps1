
# .\Convert-Policy.ps1 -XmlPolicyFile "path\to\file.xml" -BinaryDir "path\to\output" -XmlOutName "filename-no-extension" -BinaryFile "policy.cip"

param
(
    [parameter(Mandatory=$true)]
    [string] $XmlPolicyFile,
    [parameter(Mandatory=$true)]
    [string] $BinaryDir,
    [parameter(Mandatory=$true)]
    [string] $XmlOutName,
    [parameter(Mandatory=$true)]
    [string] $BinaryFile

)

# Load original XML policy file
[xml] $XmlPolicy = Get-Content $XmlPolicyFile

# Obtain policy version number from original XML policy file
$PolicyVersion = $XmlPolicy.SiPolicy.VersionEx

# Split version number and increment by .1
$split_version = $PolicyVersion.split(".")
$split_version[3] = [int]$split_version[3] + 1
$PolicyVersionNew = $split_version -join "."

# Make a copy of the original XML policy file and attach the incremented version to filename
Copy-Item $XmlPolicyFile -Destination "$BinaryDir\$XmlOutName-v$PolicyVersionNew.xml"
$XmlPolicyFileNew = "$BinaryDir\$XmlOutName-v$PolicyVersionNew.xml"

# Change policy version in updated XML policy file
Set-CIPolicyVersion -FilePath $XmlPolicyFileNew -Version $PolicyVersionNew

# Use updated XML policy file with incremented version for conversion to binary CIP
ConvertFrom-CIPolicy -XmlFilePath $XmlPolicyFileNew -BinaryFilePath "$BinaryDir\$BinaryFile"