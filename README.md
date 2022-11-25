# VDS-Sample-Utilities 2022
Extended functions for VDS Vault, VDS AutoCAD and VDS Inventor 2022.
(access 2023 version here: https://github.com/koechlm/VDS-Sample-Utilities-2023)

Library providing functions to interact with each host of the Vault Data Standard extension. 
The compiled library is part of several Data Standard Sample Configurations published under this account.

Usage in VDS PowerShell scripts:
- Load assembly: [System.Reflection.Assembly]::LoadFrom($Env:ProgramData + "\Autodesk\Vault 20xx\Extensions\DataStandard\Vault.Custom\addinVault\VdsSampleUtilities.dll')
- Instantiate the class needed:
  $InvHelpers = New-Object VdsSampleUtilities.InvHelpers
  $AcadHelpers = New-Object VdsSampleUtilities.AcadHelpers
  $VltHelpers = New-Object VdsSampleUtilities.VltHelpers

DISCLAIMER:
---------------------------------
In any case, all binaries, configuration code, templates and snippets of this solution are of "work in progress" character.
This also applies to GitHub "Release" versions.
Neither Markus Koechl, nor Autodesk represents that these samples are reliable, accurate, complete, or otherwise valid.
Accordingly, those configuration samples are provided “as is” with no warranty of any kind and you use the applications at your own risk.

Sincerely,
Markus Koechl, April 2021
