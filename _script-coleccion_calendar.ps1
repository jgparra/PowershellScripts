<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2017 v5.4.143
	 Created on:   	5/8/2019 4:25 PM
	 Created by:   	Gregorio Parra - (gregorio.parra@microsoft.com)
	 				

	 Organization: 	Microsoft

	 Filename:     	_script-coleccion_calendar.ps1

	===========================================================================

	.DESCRIPTION
		this Tool is not supported by Microsoft!!
		please, all comments ans suggestions here gregorio.parra@microsoft.com 

		this tool is created to check information about shared calendar for users  
		it will show shring urls and PublishEnabled status
	

		versions

		V 1.0  5/8/2019 - initial version

#>

$carpetas = Get-Mailbox -ResultSize unlimited| Get-MailboxFolderStatistics -FolderScope Calendar |?{$_.foldertype -ne "User Created"} |select identity 

$totalUsers = $carpetas.Count
$i=1
$cfs = @()
Foreach ($cpt in $carpetas) {
    $cptanalyze = ($cpt.Identity  -split '\\')[0] + ':\' + ($cpt.Identity  -split '\\')[1] 
    $user=($cpt.Identity  -split '\\')[0]
    write-host "$cptanalyze | $user" -ForegroundColor Green
    $cf = Get-MailboxCalendarFolder -Identity $cptanalyze
    if($cf.PublishEnabled -eq $true) {
        $cfs += New-Object -TypeName psobject -Property @{
            Usuario="$user"
            Folder = $cptanalyze
            PublishEnabled=$cf.PublishEnabled
            DetailLevel=$cf.DetailLevel
            PublishedCalendarUrl=$cf.PublishedCalendarUrl
            PublishedICalUrl=$cf.PublishedICalUrl

        }

    }    
    Write-Progress -Activity "colectando informacion..." -status "verificado $i de $totalUsers,  porcentaje: $($i/$totalUsers * 100) %"   -percentComplete ($i/$totalUsers * 100)
    $i++
}
#$cfs | Out-GridView

$timer = (Get-Date -Format yyy-mm-dd-hhmm)

$cfs | export-csv -notypeinformation -path "c:\temp\EXO-CalendarShares-$timer.csv"

#---
