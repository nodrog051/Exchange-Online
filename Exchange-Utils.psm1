#$UserCredential = Get-Credential#$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic �AllowRedirection#Import-PSSession $Session
Function Global:Test-Credentials {param( [System.Management.Automation.CredentialAttribute()]  $cred) $username = $cred.username $password = $cred.GetNetworkCredential().password
 # Get current domain using logged-on user's credentials $CurrentDomain = "LDAP://" + ([ADSI]"").distinguishedName $domain = New-Object System.DirectoryServices.DirectoryEntry($CurrentDomain,$UserName,$Password)
 if ($domain.name -eq $null){  write-host "Authentication failed - please verify your username and password."  Return $false  } else {  write-host "Successfully authenticated with domain" $domain.name  Return $true }}

Function Private:Test-ExchangeOnlineConnection { $ErrorActionPreference = "SilentlyContinue" $IsConnected = $false $sessions = Get-PSSession  foreach ($sess in $sessions){  If ($sess.ComputerName.ToString() -eq "outlook.office365.com"){   $IsConnected = $true  } } Return $IsConnected
}
Function Global:Connect-ExchangeOnline{ $UserCredential = Get-Credential if (Test-Credentials($UserCredential)){    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic  Import-PSSession $Session  If (Get-PSSession outlook.office365.com){   Return $true  }  else  {   Return $false  } } }

Function Get-MailboxFolderCount {param ( [Parameter(Mandatory=$true)][string] [string]$mbxalias = $( Read-Host "Please enter mailbox alias, please" ), [Parameter(Mandatory=$true)][string] [string]$outputFile = $( Read-Host "Please enter the path to the output file, please" )) If (Test-ExchangeOnlineConnection){  write-host "Connected to Exchange On-line!!" } else {  Write-host "You are not connected to Exchange Online!!"  Write-host "Connecting you now"  If (Connect-ExchangeOnline) {   Write-host "Connected to Exchange Online"  }  else {   Write-host "Failed to connect to exchange Online"   Exit  }   }  try{  $mbx = get-mailbox $mbxalias } catch {  write-host "No mailbox found with that alias, exiting function"  exit }
 $str = "mbx,fldcnt" $str | out-file $outputFile
 $fldcnt = Get-Mailbox -identity $mbx.alias | Get-MailboxFolderStatistics | Measure-Object | Select-Object -ExpandProperty Count write-host "Folder count for: "  $mbx.WindowsEmailAddress " is "  $fldcnt
 $str =  $mbx.WindowsEmailAddress  + ","  + $fldcnt $str | out-file $outputFile -append
}

