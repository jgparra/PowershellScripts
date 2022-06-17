﻿
$obj = @()

$ADDomain = Get-ADDomain | Select DistinguishedName
$DSSearch = New-Object System.DirectoryServices.DirectorySearcher
$DSSearch.Filter = ‘(&(objectClass=serviceConnectionPoint)(|(keywords=67661d7F-8FC4-4fa7-BFAC-E1D7794C1F68)(keywords=77378F46-2C66-4aa9-A6A6-3E7A48B19596)))’
$DSSearch.SearchRoot = ‘LDAP://CN=Configuration,’+$ADDomain.DistinguishedName

write-host "domain: $( $ADDomain.DistinguishedName)" -ForegroundColor Green
$DSSearch.FindAll() | %{

$ADSI = [ADSI]$_.Path
$autodiscover = New-Object psobject -Property @{
Server = [string]$ADSI.cn
Site = $adsi.keywords[0]
DateCreated = $adsi.WhenCreated.ToShortDateString()
LastChanged = $adsi.whenChanged.ToShortDateString()
AutoDiscoverInternalURI = [string]$adsi.ServiceBindingInformation
DN = $ADSI.distinguishedName
}
$obj += $autodiscover

}

Write-Output $obj | Select Server,Site,DateCreated,LastChanged,AutoDiscoverInternalURI, DN | ft -AutoSize -Wrap
