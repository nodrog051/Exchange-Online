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
		$mbx = get-casmailbox $alias  -EA Stop
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
		$mbx = get-mailbox $alias -EA Stop
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
		$mbx = get-mailbox $mbxalias -EA Stop
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
		$mbx = get-mailbox $alias -EA Stop
	}
	catch {
		write-host "No mailbox found with that alias, exiting function"
		break
	}

	try{
		$usr = get-user $user -EA Stop
	}
	catch {
		write-host "User not found, exiting function"
		break
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
		Test-Path $filepath -EA Stop
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
		$mbx = get-mailbox $alias -EA Stop
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
		$grp = Get-DistributionGroupMember $group -EA Stop
	}
	catch {
		Write-host "There is no distribution group with that name, exiting function"
		Exit
	}


	try{
		$mbx = get-mailbox $alias -EA Stop
	}
	catch {
		write-host "No mailbox found with that alias, exiting function"
		exit
	}

	Foreach ($entry in $grp){
		If($entry.recipientType -eq "UserMailbox"){
			Add-MailboxPermission $alias -user $entry.alias.ToString() -AccessRights FullAccess -automapping $false -Confirm:$false
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
		$usr = Get-User $user -EA Stop
	}
	catch {
		Write-host "Cannot find user, exiting function"
		Exit
	}


	try{
		$mbx = Get-Mailbox $alias -EA Stop
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
		Test-Path $filepath -EA Stop
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
		$mbx = get-mailbox $alias -EA Stop
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
		$grp = Get-DistributionGroupMember $group  -EA Stop
	}
	catch {
		Write-host "There is no distribution group with that name, exiting function"
		Break
	}


	try{
		$mbx = get-mailbox $alias -EA Stop
	}
	catch {
		write-host "No mailbox found with that alias, exiting function"
		Break
	}

	Foreach ($entry in $grp){
		If($entry.recipientType -eq "UserMailbox"){
			Add-RecipientPermission $alias -trustee $entry.alias.ToString() -AccessRights SendAs -Confirm:$false
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
		$mbx = get-mailbox $alias -EA Stop
	}
	catch {
		write-host "No mailbox found with that alias, exiting function"
		exit
	}

	try{
		$user = get-user $user -EA Stop
	}
	catch {
		write-host "User not found, exiting function"
		exit
	}

	Remove-MailboxPermission $alias -user $user -AccessRights fullaccess

}


Function Private:Search-ActiveDirectoryForObject{
param (
	[Parameter(Mandatory=$true)]
	[string]$strCat,
	[Parameter(Mandatory=$true)]
	[string]$strSearch,
	[Parameter(Mandatory=$true)]
	[string]$strProperty
)
	write-host "Category:" + $strCat
	write-host "Search: " + $strSearch
	write-host "Property: " + $strProperty


	$strFilter = "(&(objectCategory=$strCat)($strSearch))"

	write-host "Filter: " + $strFilter

	$objDomain = New-Object System.DirectoryServices.DirectoryEntry

	$objSearcher = New-Object System.DirectoryServices.DirectorySearcher
	$objSearcher.SearchRoot = $objDomain
	$objSearcher.PageSize = 1000
	$objSearcher.Filter = $strFilter
	$objSearcher.SearchScope = "Subtree"

	#$colProplist = "name", "distinguishedName"
	$colProplist = $strProperty
	foreach ($i in $colPropList){$objSearcher.PropertiesToLoad.Add($i)}

	$colResults = $objSearcher.FindAll()
	Write-Host "Count : " + $colResults.count

	If($colResults.count -ne "1"){
		Write-Host "Function Search-ActiveDirectoryForObject found more than 1 object or found nothing based on your search, exiting"
		Exit
	}	


	foreach ($objResult in $colResults){
#		$objItem = $objResult.Properties; $objItem.name
#		write-host "Properties: " + $objResult.Properties
#		write-host "DN: " + $objItem.distinguishedname
	#	If ($objItem.name -eq $strProperty) {
#			$objItem = $objResult.Properties
			$dn = $objItem.distinguishedname
	#	}
	}
	$dn
	Return $dn


}



Function Private:Enable-ApplicationImpersonation {
param (
	[Parameter(Mandatory=$false)]
	[string]$mbox,
	[Parameter(Mandatory=$true)]
	[string]$user,
	[Parameter(Mandatory=$false)]
	[string]$group 
)

	If ($mbox -AND $group){
		Write-Host "Cannot user both group and mailbox parameters at the same time"
		Break
	}

	If (!$mbox -AND !$group) {
		Write-Host "Must specify a mailbox or group as well as the user"
		Break
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
			Break
		}
	}
	
	If ($mbox){
		Try {
			$mbx = get-mailbox $mbox -ErrorAction Stop
		}
		Catch {
			write-host "No mailbox found with that identity, exiting function"
			Break
		}
		$Guid = $mbx.Guid
		$mainScopeName = "U_" + $mbx.PrimarySmtpAddress.ToString()
	}

	If ($user){
		Try {
			$usr = get-user $user -ErrorAction Stop
		}
		Catch {
			write-host "No User found with that identity, exiting function"
			Break
		}
		$mainAssName = $usr.UserPrincipalName.ToString()

	}

	If ($group){
		Try {
			$grp = Get-DistributionGroup $group -ErrorAction Stop
		}
		Catch {

			write-host "No group found with that identity, exiting function"
			Break		
		}
		$Guid = $grp.Guid

		$searchStr = "mail=" + $group.PrimarySmtpAddress
		$objGrp = Search-ActiveDirectoryForObject -strCat "Group" -strSearch $searchStr -strProperty "distinguishedName"
		$grpDN = $objGrp[2]

		$mainScopeName = "G_" + $grp.PrimarySmtpAddress.ToString()
	}
	If ($mbx) {
		write-host "Granting FullAccess to user on the mailbox"
		Grant-FullMailboxAccess -alias $mbox -user $usr.UserPrincipalName.ToString()
	}
	If ($mbx) {
		write-host "Granting SendAs rights to user on the mailbox"
		Grant-SendAsAccess -alias $mbox -user $usr.UserPrincipalName.ToString()
	}

	If ($grp){
		write-host "Granting Full Access rights to mailboxes in distribution group"
		Grant-FullAccessFromDistributionGroup -alias $usr.UserPrincipalName.ToString() -group $grp.alias.ToString()
	}

	If ($grp){
		write-host "Granting SendAs rights to mailboxes in distribution group"
		Grant-SendAsFromDistributionGroup -alias $usr.UserPrincipalName.ToString() -group $grp.alias.tostring()
	}

	write-host "Checkin to see if scope already exists"
	$myScope = $null
	$myGuid = $null
	$myGuid = "Guid -eq " + $Guid.ToString()


	$myScope = Get-ManagementScope | where {$_.recipientfilter -eq $myGuid}

	If ($myScope) {
		write-host "No scope found with this Guid, better create one"
		$newScope = $false
	}
	else{
		$newScope = $true
	}

	
	
	If ($newScope){
		write-host "Checking to see if scope name is in use"
		$useScopeName = $false

		$useScopeName = $false
		$myScope = $null
		While ($useScopeName -eq $false){

			$strAppend = "_" + (Get-Random -maximum 650000).tostring()
			$strScopeCheck = "Scope_" + $mainScopeName + $strAppend
			Try {
				$myScope = Get-ManagementScope $strScopeCheck -ErrorAction Stop
			}
			Catch {

				Write-Host "Scope Name not in use"
				$useScopeName = $true
				$strScopeName = $strScopeCheck
			}
		
		}
	}
	else{
		$strScopeName = $myScope.Name.ToString()
	}

	Write-Host "OK, we have a name for the Managment Scope, its going to be called: " + $strScopeName	
	#Create Scope Command here using strScoopeName
	If ($mbx){
		Try{
			New-ManagementScope -Name $strScopeName -RecipientRestrictionFilter "Name -eq '$mbx'" -EA Stop -whatif
		}
		Catch {
			Write-host "Failed to create Managment Scope, exiting now before I do any real damage"
			write-Host " Exit code: " + $error[0]
			Break
		}
	}
	else {
		Try{
			$dn = $grp.distinguishedname
			New-ManagementScope -Name $strScopeName -RecipientRestrictionFilter "MemberOfGroup -eq '$dn'" -EA Stop-whatif
		}
		Catch {
			Write-host "Failed to create Managment Scope, exiting now before I do any real damage"
			Break
		}

	}
	
	#Do the same thing with the RoleName
	write-host "Checking to see if RoleAssignment name is in use"
	$useRoleAssName = $false


	While ($useRoleAssName -eq $false){

		$strAppend = "_" + (Get-Random -maximum 650000).toString()
		$strRoleAssCheck = "Role_" + $mainAssName + $strAppend
		$Error.clear()
		$myRole = Get-ManagementRoleAssignment $strScopeCheck -ErrorAction SilentlyContinue

		If ($Error[0] -eq $null) {
	
			Write-Host "RoleAssignment name not in use: " + $strRoleAssCheck
			$useRoleAssName = $true
			$strRoleAssName = $strRoleAssCheck
		}

	}

	Write-Host "OK, we have a name for the RoleAssignment, its going to be called: " + $strRoleAssName	
	#Create RoleAssignment Command here using strRoleAssName

	Try {
#		New-ManagementRoleAssignment –Name:$strRoleAssName  –Role:ApplicationImpersonation –User: $usr –CustomRecipientWriteScope:$strScopeName
	}
	Catch {
		Write-Host "Failed creating the Managment rol Assignment, aaaaargh!!!!"
		Break
	}

}