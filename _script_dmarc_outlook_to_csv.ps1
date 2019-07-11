
<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2017 v5.4.143
	 Created on:   	7/11/2019 6:24 AM
	 Created by:   	gregorio dot parra at microsoft dot com
	 Organization: 	
	 Filename:     	_script_dmarc_outlook_to_csv.ps1
	 Version:		1.0
	===========================================================================
	.DESCRIPTION
		script para creacion de reporte csv desde archivos en una carpeta 
		que contiene los reportes DMARC. tambien deja archivos xml para analisis

		requires module 7Zip4Powershell
		install by using command
		install-module 7Zip4Powershell -force
#>

# recursos
#https://ipdata.co/

#region ModuleCheck
If (!(Get-Module -ListAvailable | Where-Object { $_.Name -eq "7Zip4Powershell" })) {
	Write-Error "Required Module 7Zip4Powershell does not appear to be installed. Ensure pre-requisites are followed."
	exit;
}
else {
	Write-Host "loading module 7Zip4Powershell"
	Import-Module 7Zip4Powershell
}
#endregion

#region crea dir
$execdir = "DMARC-CSV_$(get-date -f yyyy-MM-dd_HH-mm)"
mkdir $env:temp\$execdir
$workingDir = "$env:temp\$execdir"
mkdir $workingDir\output
mkdir $workingDir\xml
#explorer $workingdir
#endregion

#region saca de outlook adjuntos
$o = New-Object -comobject outlook.application
$n = $o.GetNamespace("MAPI")

$f = $n.PickFolder()
write-host "#--- numero items encontrados: $($f.Items.Count)"
$f.Items | foreach {
	if ($_.unread -eq $true) {
		$SendName = $_.SenderName
		$_.attachments | foreach {
#			Write-Host $_.filename
			$a = $_.filename
			#    If ($a.Contains("xlsx")) {
			$_.saveasfile((Join-Path $workingDir\ "$SendName--$a"))
			#   }
		}
	}
}
#endregion

#region process adjuntos para entregar un csv
#$tmpfolder = 'c:\temp\hdi\6-28'
$filereport = "report.csv"
$filerecord = "record.csv"


# Expand-7Zip -ArchiveFileName -TargetPath $MessagePath
$Zfiles = dir  $workingDir -Filter "*.*z*"
write-host "numero archivos *z* a expandir: $($Zfiles.count)" -ForegroundColor Green
Foreach ($zfile in $Zfiles) {
#	Write-Host $zfile
	try {
		Expand-7Zip -ArchiveFileName $zfile.fullname -TargetPath "$workingDir\xml" -ErrorAction Stop
	}  catch [System.Exception] { <# write-host "archivo tar: {0}" -f  $_.Exception.Message#> }
}
#sobre los expandidos, cambia ext de tar a xml
foreach ($TarXML in $(Get-ChildItem -Path "$workingDir\xml" -Filter "*.xml.tar")) { Move-Item -Path $TarXML.Fullname -Destination "$($TarXML.Fullname).xml" }

##
$allxmls = $(Get-ChildItem -Path "$workingDir\xml" -Filter "*.xml")
$total = $allxmls.count
write-host "numero total xmls: $total"
$resultfilereport = @()
$resultrowreport = @()


foreach ($xmlFile in $allxmls) {
	write-verbose "::> working on $($xmlFile.Fullname)" -ForegroundColor Green
	try {
		$xmlContent = [xml](Get-Content $xmlFile.Fullname)
	}
	catch [System.Exception] { <#Write-Host "Other exception xml: {0}" -f  $_.Exception.Message #>}
	# Create a rowkey prefix for uniqueness, like a primary key for feedback and rows
	$rkprefix = "$($xmlFile.Fullname)"
	#    $rkprefix = "$($xmlContent.feedback.report_metadata.org_name)_$($xmlContent.feedback.report_metadata.report_id)"
	foreach ($feedback in $xmlContent.feedback) {
		
		
		$filereport = new-object System.Object -erroraction SilentlyContinue
		$filereport | add-member -type noteproperty -name fileID -value $rkprefix
		#report metadata
		$filereport | add-member -type noteproperty -name rm_org_name -value $feedback.report_metadata.org_name
		$filereport | add-member -type noteproperty -name rm_email -value $feedback.report_metadata.email
		$filereport | add-member -type noteproperty -name rm_extra_contact_info -value $feedback.report_metadata.extra_contact_info
		$filereport | add-member -type noteproperty -name rm_report_id -value $feedback.report_metadata.report_id
		$filereport | add-member -type noteproperty -name rm_unixdate_begin -value $feedback.report_metadata.date_range.begin
		$filereport | add-member -type noteproperty -name rm_unixdate_end -value $feedback.report_metadata.date_range.end
		#policy published
		$filereport | add-member -type noteproperty -name pp_domain -value $feedback.policy_published.domain
		$filereport | add-member -type noteproperty -name pp_adkim -value $feedback.policy_published.adkim
		$filereport | add-member -type noteproperty -name pp_aspf -value $feedback.policy_published.aspf
		$filereport | add-member -type noteproperty -name pp_p -value $feedback.policy_published.p
		$filereport | add-member -type noteproperty -name pp_sp -value $feedback.policy_published.sp
		$filereport | add-member -type noteproperty -name pp_pct -value $feedback.policy_published.pct
		$resultfilereport += $filereport
		#review rows
		foreach ($record in $feedback.record) {
#			write-host "." -NoNewline -ForegroundColor Yellow
			$rowreport = new-object System.Object -erroraction SilentlyContinue
			$rowreport | add-member -type noteproperty -name fileID -value $rkprefix
			#row
			$rowreport | add-member -type noteproperty -name row_ip -value $record.row.source_ip
			$rowreport | add-member -type noteproperty -name row_count -value $record.row.count
			$rowreport | add-member -type noteproperty -name row_disposition -value $record.row.policy_evaluated.disposition
			$rowreport | add-member -type noteproperty -name row_aligned_dkim -value $record.row.policy_evaluated.dkim
			$rowreport | add-member -type noteproperty -name row_aligned_spf -value $record.row.policy_evaluated.spf
			$rowreport | add-member -type noteproperty -name row_reason_type -value $record.row.policy_evaluated.reason.type
			$rowreport | add-member -type noteproperty -name row_reason_comment -value $record.row.policy_evaluated.reason.comment
			#identifiers
			$rowreport | add-member -type noteproperty -name header_from -value $record.identifiers.header_from
			#auth
			$rowreport | add-member -type noteproperty -name dkim_result -value $record.auth_results.dkim.result
			$rowreport | add-member -type noteproperty -name dkim_domain -value $record.auth_results.dkim.domain
			$rowreport | add-member -type noteproperty -name dkim_selector -value $record.auth_results.dkim.selector
			
			$rowreport | add-member -type noteproperty -name spf_result -value $record.auth_results.spf.result
			$rowreport | add-member -type noteproperty -name spf_domain -value $record.auth_results.spf.domain
			
			$resultrowreport += $rowreport
		}
	}
}
$resultfilereport | Export-Csv -Path "$workingDir\output\filereport.csv" -NoTypeInformation
$resultrowreport | Export-Csv -Path "$workingDir\output\rowreport.csv" -NoTypeInformation
explorer $workingDir\output
#endregion
