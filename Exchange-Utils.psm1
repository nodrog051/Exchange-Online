#$UserCredential = Get-Credential
#$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic –AllowRedirection
#Import-PSSession $Session

Function Private:Test_ExchangeOnlineConnection {

}


Function Get-MailboxFolderCount {

param (
    [Parameter(Mandatory=$true)][string]
    [string]$mbxalias = $( Read-Host "Please enter mailbox alias, please" ),
    [Parameter(Mandatory=$true)][string]
    [string]$outputFile = $( Read-Host "Please enter the path to the output file, please" )
)


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
	write-host "Folder count for:" + $mbx.WindowsEmailAddress " is " + $fldcnt


	$str = 	$mbx.WindowsEmailAddress  + ","  + $fldcnt
	$str | out-file $outputFile -append

}
