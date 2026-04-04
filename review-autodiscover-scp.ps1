<#
.SYNOPSIS
Revisa objetos SCP de Exchange Autodiscover en Active Directory.

.DESCRIPTION
Busca objetos serviceConnectionPoint en el contenedor Configuration del bosque
filtrando por los GUID de keywords usados por Exchange Autodiscover.

Muestra para cada objeto:
- Server (CN del objeto)
- Site (extraido desde keywords tipo Site=...)
- SCPType (SCP URL o SCP Pointer)
- DateCreated y LastChanged
- AutoDiscoverInternalURI (ServiceBindingInformation)
- DN (Distinguished Name)

.NOTES
Requiere modulo ActiveDirectory (Get-ADDomain) y permisos de lectura en AD.

.OUTPUTS
Tabla en consola con columnas:
Server, Site, SCPType, DateCreated, LastChanged, AutoDiscoverInternalURI, DN

.EXAMPLE
.\review-autodiscover-scp.ps1

Ejecuta la busqueda de SCP Autodiscover y muestra los resultados en formato tabla.
#>

$obj = @()

# Obtiene el DN del dominio actual para apuntar la busqueda al contenedor Configuration.
$ADDomain = Get-ADDomain | Select-Object DistinguishedName
$DSSearch = New-Object System.DirectoryServices.DirectorySearcher

# Filtra objetos serviceConnectionPoint (SCP) de Autodiscover en Exchange.
# GUIDs usados en keywords:
# - 67661d7F-8FC4-4fa7-BFAC-E1D7794C1F68
# - 77378F46-2C66-4aa9-A6A6-3E7A48B19596
$DSSearch.Filter = '(&(objectClass=serviceConnectionPoint)(|(keywords=67661d7F-8FC4-4fa7-BFAC-E1D7794C1F68)(keywords=77378F46-2C66-4aa9-A6A6-3E7A48B19596)))'
$DSSearch.SearchRoot = 'LDAP://CN=Configuration,' + $ADDomain.DistinguishedName

Write-Host "domain: $($ADDomain.DistinguishedName)" -ForegroundColor Green
$DSSearch.FindAll() | ForEach-Object {
	$ADSI = [ADSI]$_.Path

	# Keywords puede incluir GUIDs y valores de sitio (ej. Site=Default-First-Site-Name).
	$allKeywords = @($ADSI.keywords)
	$siteKeyword = $allKeywords | Where-Object { $_ -like 'Site=*' } | Select-Object -First 1
	$siteValue = if ($siteKeyword) { ($siteKeyword -replace '^Site=', '') } else { $null }

	$scpType = if ($allKeywords -contains '67661d7F-8FC4-4fa7-BFAC-E1D7794C1F68') {
		'SCP URL'
	}
	elseif ($allKeywords -contains '77378F46-2C66-4aa9-A6A6-3E7A48B19596') {
		'SCP Pointer'
	}
	else {
		'Unknown'
	}

	$autodiscover = New-Object psobject -Property @{
		Server = [string]$ADSI.cn
		Site = $siteValue
		SCPType = $scpType
		DateCreated = $ADSI.WhenCreated.ToShortDateString()
		LastChanged = $ADSI.whenChanged.ToShortDateString()
		AutoDiscoverInternalURI = [string]$ADSI.ServiceBindingInformation
		DN = $ADSI.distinguishedName
	}
	$obj += $autodiscover
}

Write-Output $obj |
	Select-Object Server, Site, SCPType, DateCreated, LastChanged, AutoDiscoverInternalURI, DN |
	Format-Table -AutoSize -Wrap
