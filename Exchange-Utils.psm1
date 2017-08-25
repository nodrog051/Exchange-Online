#$UserCredential = Get-Credential
#$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic �AllowRedirection
#Import-PSSession $Session


Function Private:Test-Credentials {
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

Function Private:Enable-ExchangeActiveSync {
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
		$mbx = get-casmailbox $alias
	}
	catch {
		write-host "No mailbox found with that alias, exiting function"
		exit
	}
	Set-CasMailbox $alias -ActiveSyncEnabled $true


}

Function Private:Disable-ExchangeActiveSync {
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

Function Private:Connect-ExchangeOnline{
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

Function Private:Get-MailboxFolderCount {
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

Function Private:Grant-FullMailboxAccess {
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

Function Private:Grant-FullMailboxAccessFromFile {
param (
	[Parameter(Mandatory=$true)][string]
	[string]$alias = $( Read-Host "Please enter mailbox alias, please" ),
	[Parameter(Mandatory=$true)][string]
	[string]$filepath = $( Read-Host "Please enter path to file that contains UPNs, row 1 must have the text UPN" )

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

	Import-csv $filepath | %{Add-MailboxPermission $alias -user $_.UPN -AccessRights fullaccess -AutoMapping $false}


}

Function Private:Grant-FullAccessFromDistributionGroup {
param (
	[Parameter(Mandatory=$true)][string]
	[string]$alias = $( Read-Host "Please enter mailbox alias, please" ),
	[Parameter(Mandatory=$true)][string]
	[string]$group = $( Read-Host "Please enter path to file that contains UPNs that will be granted SendAs writes, please" )

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
		$grp = Get-DistributionGroupMember $group
	}
	catch {
		Write-host "There is no distribution group with that name, exiting function"
		Exit
	}


	try{
		$mbx = get-mailbox $alias
	}
	catch {
		write-host "No mailbox found with that alias, exiting function"
		exit
	}

	Foreach ($entry in $grp){
		If($entry.recipientType -eq "UserMailbox"){
			Add-MailboxPermission $alias -user $entry -AccessRights FullAccess -autmount $false -Confirm:$false
		}
	}
}


Function Private:Grant-SendAsAccess {
param (
	[Parameter(Mandatory=$true)][string]
	[string]$alias = $( Read-Host "Please enter mailbox alias, please" ),
	[Parameter(Mandatory=$true)][string]
	[string]$user = $( Read-Host "Please enter path to file that contains UPNs that will be granted SendAs writes, please" )

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
		$usr = Get-User $user
	}
	catch {
		Write-host "Cannot find user, exiting function"
		Exit
	}


	try{
		$mbx = Get-Mailbox $alias
	}
	catch {
		write-host "Mailbox not found, exiting function"
		Exit
	}

	Add-RecipientPermission $alias -trustee $user -AccessRights sendas -Confirm:$false
}


###
Function Private:Grant-SendAsAccessFromFile {
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

	Import-csv $filepath | %{Add-RecipientPermission $alias -trustee $_.UPN -AccessRights sendas -Confirm:$false}
}


Function Private:Grant-SendAsFromDistributionGroup {
param (
	[Parameter(Mandatory=$true)][string]
	[string]$alias = $( Read-Host "Please enter mailbox alias, please" ),
	[Parameter(Mandatory=$true)][string]
	[string]$group = $( Read-Host "Please enter path to file that contains UPNs that will be granted SendAs writes, please" )

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
		$grp = Get-DistributionGroupMember $group
	}
	catch {
		Write-host "There is no distribution group with that name, exiting function"
		Exit
	}


	try{
		$mbx = get-mailbox $alias
	}
	catch {
		write-host "No mailbox found with that alias, exiting function"
		exit
	}

	Foreach ($entry in $grp){
		If($entry.recipientType -eq "UserMailbox"){
			Add-RecipientPermission $alias -trustee $entry -AccessRights sendas -Confirm:$false
		}
	}
}



Function Private:Remove-FullMailboxAccess {
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


Function Private:Enable-ApplicationImpersonation {
param (
	[Parameter(Mandatory=$true)][string]
	[string]$mbox = $( Read-Host "Please enter mailbox identity" ),
	[Parameter(Mandatory=$true)][string]
	[string]$user = $( Read-Host "Please enter user who requires impersonation rights" ),
	[Parameter(Optional=$true)][string]
	[string]$group = $( Read-Host "Please enter the group name that contains the users that the service account will be granted impersonation right too" )
)

	If ($mbox -AND $group){
		Write-Host "Cannot user both group and mailbox parameters at the same time"
		Exit
	}

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
	
	If ($mbox){
		try{
			$mbx = get-mailbox $mbox
		}
		catch {
			write-host "No mailbox found with that identity, exiting function"
			exit
		}
	}

	If ($user){
		try{
			$usr = get-user $user
		}
		catch {
			write-host "No User found with that identity, exiting function"
			exit
		}
	}

	If ($group){
		try{
			$grp = Get-DistributionGroup $group
		}
		catch{
			write-host "No group found with that identity, exiting function"
			exit		
		}
	}

	write-host "Granting FullAccess to user on the mailbox"
	Grant-FullMailboxAccess -alias $mbox  -user $user

	write-host "Granting SendAs rights to user on the mailbox"
	Grant-SendAsAccess -alias $mbox -user $user
	
	write-host "Checking to see if scope name already exists, if probably will"
	$strScope = "Scope_" + $mbx.Name.ToString()

	Try {
		Get-ManagementScope $strScope
	}
	Catch {
		Wite-Host "In progress"
	}
	$ErrorActionPreference = "SilentlyContinue"
	If (Get-ManagementScope $strScope) {
		write-host "ManagmentScope name allready in use, going to generate some random numbers and append them to the name"
		$loop = $true
		while ($loop){
			$strAppend = Get-Random -maximum 6500
			$tmpStr = $strScope + "-" + $strAppend
			if (Get-ManagmentScope $tmpStr) {
				$loop = $true
			}
			else {
				$strScope = $tmpStr
				$loop = $false
			}

		}

	}
	Write-Host "OK, we have a name for the Managment Scope, its going to be called: " + $strScope
	write-host "Checking to see if assignment role name allready exists, if probably will"
	$strRole = "Role_" + $usr.Name.ToString()
	If (Get-ManagementRoleAssignment $strRole) {
		write-host "Managment Role Assignment name already in use, going to generate some random numbers and append them to the name"
		$loop = $true
		while ($loop){
			$strAppend = Get-Random -maximum 6500
			Write-host $strAppend
			$tmpStr = $strRole + "-" + $strAppend
			if (Get-ManagmentRoleAssignment $tmpStr) {
				$loop = $true
			}
			else {
				$strRole = $tmpStr
				$loop = $false
			}

		}

	}

	Write-Host "OK, we have a name for the Managment Roel Assignment, its going to be called: " + $strRole


}