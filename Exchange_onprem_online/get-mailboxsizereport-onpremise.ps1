 ####################################################
####################################################
#
#
#	name:		get-mailboxsizereport.ps1
#	Version:	0.5
#	date:		1/30/2012
#	autor:		Gregorio Parra - gregorio.parra@microsoft.com
#	
#	Description: thi script collect information from user´s mailbox,
#	like alias, displayname, quotas and OU from get-mailbox commandlet
#	and information like  itemcount  mailbox status and size in MB 
#	from get-mailboxstatistics, in only one report that is saved on 
#   C:\exrap\other_reports\User_Mailbox_Size_and_Item_Count_all_data.csv
#	to import in excel and use a lot a pivot Table to get reports.
#	
#	
#		for now this could be used on Exchange 2007 and 2010 only, 
#		any feedback is very welcome!
#	
####################################################

function namefile_date {
# genera un nombre, que va con adicional de la fecha y hora del nomento de ejecucion de la funcion
# ideal para reportar resultados
# ejemplo:  namefile_date -name "user-mailbox-size" -extension ".csv"
	param(
		$Name,  # nombre del archivo
		$extension  # extension con la que quedaria, debe ir con el punto
	)
	$filename = Get-Date -Format yyyyMMdd-HHmms
	$_outfilename = $Name+"_"+$filename+$extension
#	Write-Warning $_outfilename
	return $_outfilename
}

##main 

If (!($args.Count -eq 0)){
	$outputfile = $args[0]
}
Else{
	#If the csv file has not been specified then prompt for it
	$Outputfile = Read-Host "Enter CSV file path (eg. C:\Reports)"
}

#Remove any leading or trailing quotes
$Outputfile = $Outputfile.TrimEnd('"')
$Outputfile = $Outputfile.Trim('"')
<# added because was created to differentiate Ex2007
$2007snapin = Get-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.Admin
if ($2007snapin)
{
	$AdminSessionADSettings.ViewEntireForest = 1
}
else
{
	$2010snapin = Get-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.E2010
	if ($2010snapin)
	{
		Set-ADServerSettings -ViewEntireForest $true
	}
}
#>
Set-ADServerSettings -ViewEntireForest $true
$hashtableDB = @{}
Get-MailboxDatabase | select identity, prohibitsendreceivequota | %{ $hashtableDB.add("$($_.identity)", $_.ProhibitSendReceiveQuota)}



$result = @()
$a = Get-Mailbox -resultsize unlimited
$number = $a.count
write-host "usuarios: $number"
$i =1
$a | ForEach-Object {
    $myObject2 = New-Object System.Object  -ErrorAction SilentlyContinue
    $myObject2 | Add-Member -type NoteProperty -name alias -Value $_.alias
    $myObject2 | Add-Member -type NoteProperty -name PrimarySMTPAddress -Value $_.PrimarySMTPAddress
    $myObject2 | Add-Member -type NoteProperty -name DisplayName -Value $_.DisplayName -ErrorAction SilentlyContinue
    $myObject2 | Add-Member -type NoteProperty -name IssueWarningQuota -Value $_.IssueWarningQuota -ErrorAction SilentlyContinue
    $myObject2 | Add-Member -type NoteProperty -name ProhibitSendQuota -Value $_.ProhibitSendQuota -ErrorAction SilentlyContinue
    $myObject2 | Add-Member -type NoteProperty -name ProhibitSendReceiveQuota -Value $_.ProhibitSendReceiveQuota -ErrorAction SilentlyContinue
    $myObject2 | Add-Member -type NoteProperty -name UseDatabaseQuotaDefaults -Value $_.UseDatabaseQuotaDefaults -ErrorAction SilentlyContinue
    $myObject2 | Add-Member -type NoteProperty -name DBQuota -Value $hashtableDB["$($_.Database)"].value.toMB() -ErrorAction SilentlyContinue
#    $myObject2 | Add-Member -type NoteProperty -name DBQuota -Value $hashtableDB["$($_.Database)"].value -ErrorAction SilentlyContinue
    $myObject2 | Add-Member -type NoteProperty -name ServerName -Value $_.ServerName -ErrorAction SilentlyContinue
    $myObject2 | Add-Member -type NoteProperty -name OrganizationalUnit -Value $_.OrganizationalUnit -ErrorAction SilentlyContinue
    $myObject2 | Add-Member -type NoteProperty -name MaxSendSize -Value $_.MaxSendSize -ErrorAction SilentlyContinue
    $myObject2 | Add-Member -type NoteProperty -name MaxReceiveSize -Value $_.MaxReceiveSize -ErrorAction SilentlyContinue
    $myObject2 | Add-Member -type NoteProperty -name Database -Value $_.Database -ErrorAction SilentlyContinue
    $myObject2 | Add-Member -type NoteProperty -name RetainDeletedItemsFor -Value $_.RetainDeletedItemsFor -ErrorAction SilentlyContinue
	$myObject2 | Add-Member -type NoteProperty -name LegacyExchangeDN -Value $_.LegacyExchangeDN -ErrorAction SilentlyContinue
    $st = Get-MailboxStatistics $_.alias
	if ($st -eq $null){write-host "no tiene datos:" $_.alias}
	else{
	    $myObject2 | Add-Member -type NoteProperty -name DeletedItemCount -Value $st.DeletedItemCount -ErrorAction SilentlyContinue
	    $myObject2 | Add-Member -type NoteProperty -name ItemCount -Value $st.ItemCount -ErrorAction SilentlyContinue
	    $myObject2 | Add-Member -type NoteProperty -name LastLogoffTime -Value $st.LastLogoffTime -ErrorAction SilentlyContinue
	    $myObject2 | Add-Member -type NoteProperty -name LastLogonTime -Value $st.LastLogonTime -ErrorAction SilentlyContinue
	    $myObject2 | Add-Member -type NoteProperty -name LastLoggedOnUserAccount -Value $st.LastLoggedOnUserAccount -ErrorAction SilentlyContinue
		
	    $myObject2 | Add-Member -type NoteProperty -name StorageLimitStatus -Value $st.StorageLimitStatus -ErrorAction SilentlyContinue
	    $myObject2 | Add-Member -type NoteProperty -name TotalDeletedItemSizeMB -Value $st.TotalDeletedItemSize.value.toMB() -ErrorAction SilentlyContinue
	    $myObject2 | Add-Member -type NoteProperty -name TotalItemSizeMB -Value $st.TotalItemSize.value.toMB() -ErrorAction SilentlyContinue
	}
    $result += $myObject2 
$a = ($i/$number)*100 
Write-Progress -Activity "verificado $i de $number,"     -PercentComplete $a -CurrentOperation    "$a% complete"  -Status "Please wait."
$i++
	
}

$outfile = namefile_date -name "user-mailbox-size" -extension ".csv"
$outfilename = $Outputfile+"\"+$outfile
$result | Export-Csv -Path $outfilename -NoTypeInformation

