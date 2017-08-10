#$UserCredential = Get-Credential
#$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic �AllowRedirection
#Import-PSSession $Session

Function Global:Test-Credentials {
param(
	[System.Management.Automation.CredentialAttribute()] 
	$cred
)
	$username = $cred.username
	$password = $cred.GetNetworkCredential().password

	# Get current domain using logged-on user's credentials
	$CurrentDomain = "LDAP://" + ([ADSI]"").distinguishedName
	$domain = New-Object System.DirectoryServices.DirectoryEntry($CurrentDomain,$UserName,$Password)

	if ($domain.name -eq $null){
		write-host "Authentication failed - please verify your username and password."
		Return $false
	
	}
	else {
		write-host "Successfully authenticated with domain" $domain.name
		Return $true
	}
}

Function Global:Enable-ExchangeActiveSync {
param (
	[Parameter(Mandatory=$true)][string]
	[string]$alias = $( Read-Host "Please enter mailbox alias, please" )
)

	If (Test-ExchangeOnlineConnection){
		write-host "Connected to Exchange On-line!!"
	}
	else {
		Write-host "You are not connected to Exchange Online!!"
		Write-host "Connecting you now"
		If (Connect-ExchangeOnline) {
			Write-host "Connected to Exchange Online"
		}
		else {
			Write-host "Failed to connect to exchange Online"
			Exit
		}
		
	}

	try{
		$mbx = get-mailbox $alias
	}
	catch {
		write-host "No mailbox found with that alias, exiting function"
		exit
	}
	Set-CasMailbox $alias -ActiveSyncEnabled $true


}

Function Global:Disable-ExchangeActiveSync {
param (
	[Parameter(Mandatory=$true)][string]
	[string]$alias = $( Read-Host "Please enter mailbox alias, please" )
)
	If (Test-ExchangeOnlineConnection){
		write-host "Connected to Exchange On-line!!"
	}
	else {
		Write-host "You are not connected to Exchange Online!!"
		Write-host "Connecting you now"
		If (Connect-ExchangeOnline) {
			Write-host "Connected to Exchange Online"
		}
		else {
			Write-host "Failed to connect to exchange Online"
			Exit
		}
		
	}

	try{
		$mbx = get-mailbox $alias
	}
	catch {
		write-host "No mailbox found with that alias, exiting function"
		exit
	}

	Set-CasMailbox $alias -ActiveSyncEnabled $false

}



Function Private:Test-ExchangeOnlineConnection {
	$ErrorActionPreference = "SilentlyContinue"
	$IsConnected = $false
	$sessions = Get-PSSession
	
	foreach ($sess in $sessions){
		If ($sess.ComputerName.ToString() -eq "outlook.office365.com"){
			$IsConnected = $true
		}
	}
	Return $IsConnected

}

Function Global:Connect-ExchangeOnline{
	$UserCredential = Get-Credential
	if (Test-Credentials($UserCredential)){
	
	
		$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic
		Import-PSSession $Session
		If (Get-PSSession outlook.office365.com){
			Return $true
		}
		else
		{
			Return $false
		}
	}
	
}



Function Get-MailboxFolderCount {
param (
	[Parameter(Mandatory=$true)][string]
	[string]$mbxalias = $( Read-Host "Please enter mailbox alias, please" ),
	[Parameter(Mandatory=$true)][string]
	[string]$outputFile = $( Read-Host "Please enter the path to the output file, please" )
)
	If (Test-ExchangeOnlineConnection){
		write-host "Connected to Exchange On-line!!"
	}
	else {
		Write-host "You are not connected to Exchange Online!!"
		Write-host "Connecting you now"
		If (Connect-ExchangeOnline) {
			Write-host "Connected to Exchange Online"
		}
		else {
			Write-host "Failed to connect to exchange Online"
			Exit
		}
		
	}
	
	try{
		$mbx = get-mailbox $mbxalias
	}
	catch {
		write-host "No mailbox found with that alias, exiting function"
		exit
	}

	$str = "mbx,fldcnt"
	$str | out-file $outputFile

	$fldcnt = Get-Mailbox -identity $mbx.alias | Get-MailboxFolderStatistics | Measure-Object | Select-Object -ExpandProperty Count
	write-host "Folder count for: "  $mbx.WindowsEmailAddress " is "  $fldcnt


	$str = 	$mbx.WindowsEmailAddress  + ","  + $fldcnt
	$str | out-file $outputFile -append

}

Function Global:Grant-FullMailboxAccess {
param (
	[Parameter(Mandatory=$true)][string]
	[string]$alias = $( Read-Host "Please enter mailbox alias, please" ),
	[Parameter(Mandatory=$true)][string]
	[string]$user = $( Read-Host "Please enter user who requires access, please" )

)

	If (Test-ExchangeOnlineConnection){
		write-host "Connected to Exchange On-line!!"
	}
	else {
		Write-host "You are not connected to Exchange Online!!"
		Write-host "Connecting you now"
		If (Connect-ExchangeOnline) {
			Write-host "Connected to Exchange Online"
		}
		else {
			Write-host "Failed to connect to exchange Online"
			Exit
		}
		
	}

	try{
		$mbx = get-mailbox $alias
	}
	catch {
		write-host "No mailbox found with that alias, exiting function"
		exit
	}

	try{
		$user = get-user $user
	}
	catch {
		write-host "User not found, exiting function"
		exit
	}

	Add-MailboxPermission $alias -user $user -AccessRights fullaccess -AutoMapping $false


}

Function Global:Grant-FullMailboxAccessFromFile {
param (
	[Parameter(Mandatory=$true)][string]
	[string]$alias = $( Read-Host "Please enter mailbox alias, please" ),
	[Parameter(Mandatory=$true)][string]
	[string]$filepath = $( Read-Host "Please enter path to file that contains UPNs, please" )

)

	If (Test-ExchangeOnlineConnection){
		write-host "Connected to Exchange On-line!!"
	}
	else {
		Write-host "You are not connected to Exchange Online!!"
		Write-host "Connecting you now"
		If (Connect-ExchangeOnline) {
			Write-host "Connected to Exchange Online"
		}
		else {
			Write-host "Failed to connect to exchange Online"
			Exit
		}
		
	}


	try{
		Test-Path $filepath
	}
	catch {
		Write-host "There is no file at that location, exiting function"
		Exit
	}

	$inFilePath = $filepath
	$csvColumnNames = (Get-Content $inFilePath | Select-Object -First 1)
	If ($csvColumnNames -ne "UPN"){
		write-host "Column heading has to be 'UPN' exiting function"
		exit
	}

	try{
		$mbx = get-mailbox $alias
	}
	catch {
		write-host "No mailbox found with that alias, exiting function"
		exit
	}

	try{
		$user = get-user $user
	}
	catch {
		write-host "User not found, exiting function"
		exit
	}

	Import-csv $filepath | %{Add-MailboxPermission $alias -user $_.UPN -AccessRights fullaccess -AutoMapping $false}


}


Function Global:Grant-SendAsAccessFromFile {
param (
	[Parameter(Mandatory=$true)][string]
	[string]$alias = $( Read-Host "Please enter mailbox alias, please" ),
	[Parameter(Mandatory=$true)][string]
	[string]$filepath = $( Read-Host "Please enter path to file that contains UPNs that will be granted SendAs writes, please" )

)

	If (Test-ExchangeOnlineConnection){
		write-host "Connected to Exchange On-line!!"
	}
	else {
		Write-host "You are not connected to Exchange Online!!"
		Write-host "Connecting you now"
		If (Connect-ExchangeOnline) {
			Write-host "Connected to Exchange Online"
		}
		else {
			Write-host "Failed to connect to exchange Online"
			Exit
		}
		
	}


	try{
		Test-Path $filepath
	}
	catch {
		Write-host "There is no file at that location, exiting function"
		Exit
	}

	$inFilePath = $filepath
	$csvColumnNames = (Get-Content $inFilePath | Select-Object -First 1)
	If ($csvColumnNames -ne "UPN"){
		write-host "Column heading has to be 'UPN' exiting function"
		exit
	}

	try{
		$mbx = get-mailbox $alias
	}
	catch {
		write-host "No mailbox found with that alias, exiting function"
		exit
	}

	try{
		$user = get-user $user
	}
	catch {
		write-host "User not found, exiting function"
		exit
	}

	Import-csv $filepath | %{Add-RecipientPermission $alias -trustee $_.UPN -AccessRights sendas}
}


Function Global:Remove-FullMailboxAccess {
param (
	[Parameter(Mandatory=$true)][string]
	[string]$alias = $( Read-Host "Please enter mailbox alias, please" ),
	[Parameter(Mandatory=$true)][string]
	[string]$user = $( Read-Host "Please enter user who requires access, please" )

)

	If (Test-ExchangeOnlineConnection){
		write-host "Connected to Exchange On-line!!"
	}
	else {
		Write-host "You are not connected to Exchange Online!!"
		Write-host "Connecting you now"
		If (Connect-ExchangeOnline) {
			Write-host "Connected to Exchange Online"
		}
		else {
			Write-host "Failed to connect to exchange Online"
			Exit
		}
		
	}

	try{
		$mbx = get-mailbox $alias
	}
	catch {
		write-host "No mailbox found with that alias, exiting function"
		exit
	}

	try{
		$user = get-user $user
	}
	catch {
		write-host "User not found, exiting function"
		exit
	}

	Remove-MailboxPermission $alias -user $user -AccessRights fullaccess

}


Function Global:Check-ManagementScope {
param (
	[Parameter(Mandatory=$true)][string]
	[string]$alias = $( Read-Host "Please enter mailbox alias, please" ),
	[Parameter(Mandatory=$true)][string]
	[string]$user = $( Read-Host "Please enter user who requires access, please" )

)

	If (Test-ExchangeOnlineConnection){
		write-host "Connected to Exchange On-line!!"
	}
	else {
		Write-host "You are not connected to Exchange Online!!"
		Write-host "Connecting you now"
		If (Connect-ExchangeOnline) {
			Write-host "Connected to Exchange Online"
		}
		else {
			Write-host "Failed to connect to exchange Online"
			Exit
		}
		
	}

	try{
		$mbx = get-mailbox $alias
	}
	catch {
		write-host "No mailbox found with that alias, exiting function"
		exit
	}

	try{
		$user = get-user $user
	}
	catch {
		write-host "User not found, exiting function"
		exit
	}

	Remove-MailboxPermission $alias -user $user -AccessRights fullaccess

}







