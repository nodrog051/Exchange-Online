# Exchange-Online
Powershell command to run against an Exchange-Online installation
In order to get the module to load when you start powershell do the following in a powersheel console
run $profile
You should see something like this
\\server01\home\username\WindowsPowerShell\Microsoft.PowerShell_profile.ps1

Copy the module file to the above location

Create the file Microsoft.PowerShell_profile.ps1 if it doesn't exist
Add the following line to the file

Import-Module \\server01\home\username\WindowsPowerShell\


