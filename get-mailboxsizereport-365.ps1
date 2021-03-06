####################################################
####################################################
#
#
#	name:		get-mailboxsizereport-365.ps1
#	Version:	0.1
#	date:		7/14/2016
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

function get_GBValue{
    param(
    [string]$valor
    )
#    write-host "valor ingresado: $valor"
    $ItemSizeString =  $valor
    $GBvalue = "{0:N2}" -f ($ItemSizeString.SubString(($ItemSizeString.IndexOf("(") + 1),($itemSizeString.IndexOf(" bytes") - ($ItemSizeString.IndexOf("(") + 1))).Replace(",","")/1024/1024/1024)
    return $GBvalue
    <#
para pasarlo a GB 
$ItemSizeString = $st.TotalItemSize.ToString() 
$GBvalue = "{0:N2}" -f ($ItemSizeString.SubString(($ItemSizeString.IndexOf("(") + 1),($itemSizeString.IndexOf(" bytes") - ($ItemSizeString.IndexOf("(") + 1))).Replace(",","")/1024/1024/1024)

#>
    [int](Get-Random -Minimum 100 -Maximum ($valor -as [int]))
}

##main 

If (!($args.Count -eq 0)){
	$outputfile = $args[0]
}
Else{
	#If the csv file has not been specified then prompt for it
	$Outputfile = Read-Host "Enter CSV file path (eg. C:\Reports)"

	#Remove any leading or trailing quotes
	$Outputfile = $Outputfile.TrimEnd('"')
	$Outputfile = $Outputfile.Trim('"')
}
#$stopWatch = [system.diagnostics.stopwatch]::startNew()
$result = @()
$a = Get-Mailbox -resultsize unlimited
#$a = Get-Mailbox -resultsize 100
$totalUsers = $a.count
Write-Host -ForegroundColor cyan "numero de usuarios: $totalUsers"
Write-Host  "segundos en busqueda: $($stopWatch.Elapsed.Seconds)"
#$stopWatch.Elapsed
#$secbuscar = $stopWatch.Elapsed

$i = 0;
$a | ForEach-Object {
    $myObject2 = New-Object System.Object  -ErrorAction SilentlyContinue
#    write-host -foregroundcolor cyan "." -NoNewline
    $myObject2 | Add-Member -type NoteProperty -name alias -Value $_.alias
    $myObject2 | Add-Member -type NoteProperty -name DisplayName -Value $_.DisplayName -ErrorAction SilentlyContinue
    $myObject2 | Add-Member -type NoteProperty -name IssueWarningQuota -Value $_.IssueWarningQuota -ErrorAction SilentlyContinue
    $myObject2 | Add-Member -type NoteProperty -name ProhibitSendQuota -Value $_.ProhibitSendQuota -ErrorAction SilentlyContinue
    $myObject2 | Add-Member -type NoteProperty -name ProhibitSendReceiveQuota -Value $_.ProhibitSendReceiveQuota -ErrorAction SilentlyContinue
    $myObject2 | Add-Member -type NoteProperty -name RetainDeletedItemsFor -Value $_.RetainDeletedItemsFor -ErrorAction SilentlyContinue
	$myObject2 | Add-Member -type NoteProperty -name LegacyExchangeDN -Value $_.LegacyExchangeDN -ErrorAction SilentlyContinue
	$myObject2 | Add-Member -type NoteProperty -name ServerName -Value $_.ServerName -ErrorAction SilentlyContinue
	$myObject2 | Add-Member -type NoteProperty -name Database -Value $_.Database -ErrorAction SilentlyContinue
	$myObject2 | Add-Member -type NoteProperty -name PrimarySmtpAddress -Value $_.PrimarySmtpAddress -ErrorAction SilentlyContinue
	$myObject2 | Add-Member -type NoteProperty -name WhenMailboxCreated -Value $_.WhenMailboxCreated -ErrorAction SilentlyContinue
	$myObject2 | Add-Member -type NoteProperty -name LitigationHoldEnabled -Value $_.LitigationHoldEnabled -ErrorAction SilentlyContinue
	$myObject2 | Add-Member -type NoteProperty -name LitigationHoldDate -Value $_.LitigationHoldDate -ErrorAction SilentlyContinue
	$myObject2 | Add-Member -type NoteProperty -name LitigationHoldDuration -Value $_.LitigationHoldDuration -ErrorAction SilentlyContinue
	$myObject2 | Add-Member -type NoteProperty -name AuditEnabled -Value $_.AuditEnabled -ErrorAction SilentlyContinue
	$myObject2 | Add-Member -type NoteProperty -name ArchiveStatus -Value $_.ArchiveStatus -ErrorAction SilentlyContinue
	$myObject2 | Add-Member -type NoteProperty -name ArchiveQuota -Value $_.ArchiveQuota -ErrorAction SilentlyContinue
    $st = Get-MailboxStatistics $_.alias
	if ($st -eq $null) {
		write-host "no tiene datos:" $_.alias
		$myObject2 | Add-Member -type NoteProperty -name DeletedItemCount -Value 0 -ErrorAction SilentlyContinue
		$myObject2 | Add-Member -type NoteProperty -name ItemCount -Value 0 -ErrorAction SilentlyContinue
		$myObject2 | Add-Member -type NoteProperty -name LastLogoffTime -Value "null" -ErrorAction SilentlyContinue
		$myObject2 | Add-Member -type NoteProperty -name LastLogonTime -Value "null" -ErrorAction SilentlyContinue
		$myObject2 | Add-Member -type NoteProperty -name StorageLimitStatus -Value "null" -ErrorAction SilentlyContinue
		
		$myObject2 | Add-Member -type NoteProperty -name TotalDeletedItemSize -Value 0 -ErrorAction SilentlyContinue
		$myObject2 | Add-Member -type NoteProperty -name TotalDeletedItemSizeGB -Value 0 -ErrorAction SilentlyContinue
		
		$myObject2 | Add-Member -type NoteProperty -name TotalItemSize -Value 0 -ErrorAction SilentlyContinue
		$myObject2 | Add-Member -type NoteProperty -name TotalItemSizeGB -Value 0 -ErrorAction SilentlyContinue
		
	}
	else{
	    $myObject2 | Add-Member -type NoteProperty -name DeletedItemCount -Value $st.DeletedItemCount -ErrorAction SilentlyContinue
	    $myObject2 | Add-Member -type NoteProperty -name ItemCount -Value $st.ItemCount -ErrorAction SilentlyContinue
	    $myObject2 | Add-Member -type NoteProperty -name LastLogoffTime -Value $st.LastLogoffTime -ErrorAction SilentlyContinue
	    $myObject2 | Add-Member -type NoteProperty -name LastLogonTime -Value $st.LastLogonTime -ErrorAction SilentlyContinue
	    $myObject2 | Add-Member -type NoteProperty -name StorageLimitStatus -Value $st.StorageLimitStatus -ErrorAction SilentlyContinue

	    $myObject2 | Add-Member -type NoteProperty -name TotalDeletedItemSize -Value $st.TotalDeletedItemSize -ErrorAction SilentlyContinue
	    $myObject2 | Add-Member -type NoteProperty -name TotalDeletedItemSizeGB -Value $(get_GBValue($st.TotalDeletedItemSize.tostring())) -ErrorAction SilentlyContinue

	    $myObject2 | Add-Member -type NoteProperty -name TotalItemSize -Value $st.TotalItemSize -ErrorAction SilentlyContinue
	    $myObject2 | Add-Member -type NoteProperty -name TotalItemSizeGB -Value $(get_GBValue($st.TotalItemSize.tostring()))  -ErrorAction SilentlyContinue
	}
	if ($_.ArchiveStatus -eq "Active") {
		$archt = Get-MailboxStatistics $_.alias -archive
		$myObject2 | Add-Member -type NoteProperty -name Archive_TotalItemSize -Value $archt.TotalItemSize -ErrorAction SilentlyContinue
		$myObject2 | Add-Member -type NoteProperty -name Archive_ItemCount -Value $archt.ItemCount -ErrorAction SilentlyContinue
		$myObject2 | Add-Member -type NoteProperty -name Archive_TotalDeletedItemSize -Value $archt.TotalDeletedItemSize -ErrorAction SilentlyContinue
		$myObject2 | Add-Member -type NoteProperty -name Archive_DeletedItemCount -Value $archt.DeletedItemCount -ErrorAction SilentlyContinue
	}
	else {
		$myObject2 | Add-Member -type NoteProperty -name Archive_TotalItemSize -Value 0 -ErrorAction SilentlyContinue
		$myObject2 | Add-Member -type NoteProperty -name Archive_ItemCount -Value 0 -ErrorAction SilentlyContinue
		$myObject2 | Add-Member -type NoteProperty -name Archive_TotalDeletedItemSize -Value 0 -ErrorAction SilentlyContinue
		$myObject2 | Add-Member -type NoteProperty -name Archive_DeletedItemCount -Value 0 -ErrorAction SilentlyContinue
		
	}	
	$result += $myObject2
	$i++
#	Write-Progress -Activity "colectando informacion..." -status "porcentaje: $($i/$totalUsers * 100) %"   -percentComplete ($i/$totalUsers * 100)
	Write-Progress -Activity "colectando informacion..." -status "verificado $i de $totalUsers,  porcentaje: $($i/$totalUsers * 100) %"   -percentComplete ($i/$totalUsers * 100)
	
}

$outfile = namefile_date -name "user-mailbox-size-365" -extension ".csv"
$outfilename = $Outputfile+"\"+$outfile
$result | Export-Csv -Path $outfilename -NoTypeInformation
Write-Host  "segundos en reporte: "
#$stopWatch.Elapsed
#Write-Host  "segundos en busqueda: $secbuscar "

