
<#
.SYNOPSIS
    Exchange Schema Version Checker
    Verificador de Versión de Schema de Exchange

.DESCRIPTION
    This script queries Active Directory to determine the Exchange Server version 
    installed in the environment by checking three version numbers:
    1. Schema Version (rangeupper attribute)
    2. Organization Version (objectVersion in Configuration)
    3. Domain Version (objectVersion in Domain)
    
    Este script consulta Active Directory para determinar la versión de Exchange Server
    instalada en el entorno verificando tres números de versión:
    1. Versión de Schema (atributo rangeupper)
    2. Versión de Organización (objectVersion en Configuration)
    3. Versión de Dominio (objectVersion en Domain)

.NOTES
    Author: Gregorio Parra
    Requires: ActiveDirectory PowerShell Module
    Requires: Read permissions to Active Directory Schema and Configuration partitions
    
.OUTPUTS
    PSCustomObject with SchemaType, VersionNumber, and VersionName properties
    
.EXAMPLE
    .\exchschemaversion.ps1
    
    SchemaType  VersionNumber           VersionName
    ----------  -------------           -----------
    Exchange    17003 – 16762 – 13237   2019 CU16-18 / SE
#>

#Requires -Modules ActiveDirectory

#region Exchange Version Lookup Table
<#
Hash table containing Exchange version mappings
Format: "SchemaVersion-OrgVersion-DomainVersion" = "Exchange Version Name"

Exchange uses three version numbers in Active Directory:
1. Schema Version (rangeupper): Updated when AD schema changes are required
2. Organization Version (objectVersion): Updated at Configuration partition level
3. Domain Version (objectVersion): Updated at Domain partition level

Tabla hash que contiene los mapeos de versiones de Exchange
Formato: "VersiónSchema-VersiónOrg-VersiónDominio" = "Nombre de Versión Exchange"

Exchange usa tres números de versión en Active Directory:
1. Versión de Schema (rangeupper): Se actualiza cuando se requieren cambios en el schema de AD
2. Versión de Organización (objectVersion): Se actualiza a nivel de partición Configuration
3. Versión de Dominio (objectVersion): Se actualiza a nivel de partición Domain
#>
$ExchSchemaVersions = @{
	# Exchange 2000
	"4397-N/A-4406" = "2000 RTM"
	"4406-N/A-4406" = "2000 SP3"
	
	# Exchange 2003
	"6870-6903-6936" = "2003 RTM/SP2"
	
	# Exchange 2007
	"10637-10666-10628" = "2007 RTM"
	"11116-11221-11221" = "2007 SP1"
	"14622-11222-11221" = "2007 SP2"
	"14625-11222-11221" = "2007 SP3"
	
	# Exchange 2010
	"14622-12640-12639" = "2010 RTM"
	"14726-13214-13040" = "2010 SP1"
	"14732-14247-13040" = "2010 SP2"
	"14734-14322-13040" = "2010 SP3"
	
	# Exchange 2013
	"15137-15449-13236" = "2013"
	"15254-15614-13236" = "2013 CU1"
	"15281-15688-13236" = "2013 CU2"
	"15283-15763-13236" = "2013 CU3"
	"15292-15844-13236" = "2013 SP1/CU4"
	"15300-15870-13236" = "2013 CU5"
	"15303-15965-13236" = "2013 CU6"
	"15312-15965-13236" = "2013 CU7-9"
	"15312-16130-13236" = "2013 CU10-14"
	"15312-16213-13236" = "2013 CU15-23"
	
	# Exchange 2016
	"15317-16041-13236" = "2016 Preview"
	"15317-16210-13236" = "2016 RTM"
	"15323-16211-13236" = "2016 CU1"
	"15325-16212-13236" = "2016 CU2"
	"15326-16212-13236" = "2016 CU3"
	"15327-16213-13236" = "2016 CU4"
	"15330-16213-13236" = "2016 CU5"
	"15330-16214-13236" = "2016 CU6"
	"15332-16214-13236" = "2016 CU7-9"
	"15332-16215-13237" = "2016 CU10-11"
	"15333-16215-13237" = "2016 CU12-18"
	"15334-16217-13237" = "2016 CU19-22"
	"15334-16219-13237" = "2016 CU23"
	
	# Exchange 2019 / Subscription Edition (SE)
	# Starting from CU12 (Oct 2021), supports both traditional and subscription licensing
	# A partir de CU12 (Oct 2021), soporta licenciamiento tradicional y por suscripción
	"17000-16751-13237" = "2019 RTM"
	"17001-16752-13237" = "2019 CU1"
	"17001-16754-13237" = "2019 CU2-7"
	"17002-16754-13237" = "2019 CU8-11"
	"17002-16756-13237" = "2019 CU12-13"
	"17002-16758-13237" = "2019 CU14"
	"17003-16760-13237" = "2019 CU15"
	"17003-16762-13237" = "2019 CU16-18 / SE"
	"17003-16764-13237" = "2019 CU19-20 / SE"
	"17003-16766-13237" = "2019 CU21-22 / SE"
	"17003-16768-13237" = "2019 CU23-24 / SE"
	
}
#endregion Exchange Version Lookup Table

#region Query Active Directory for Exchange Version Numbers
<#
This section queries Active Directory to extract the three version numbers
that uniquely identify an Exchange version:

Esta sección consulta Active Directory para extraer los tres números de versión
que identifican únicamente una versión de Exchange:
#>

# Get Active Directory Root DSE (Directory Service Entry)
# Contains naming contexts for Schema, Configuration, and Domain partitions
# Obtiene el Root DSE de Active Directory (Directory Service Entry)
# Contiene los contextos de nomenclatura para particiones Schema, Configuration y Domain
$RootDSE = Get-ADRootDSE

# 1. SCHEMA VERSION (Forest-level)
# Query the 'ms-Exch-Schema-Version-Pt' object in the Schema partition
# The 'rangeupper' attribute contains the Exchange Schema version number
# This changes when Exchange setup extends the AD schema
#
# VERSIÓN DE SCHEMA (Nivel de Forest)
# Consulta el objeto 'ms-Exch-Schema-Version-Pt' en la partición Schema
# El atributo 'rangeupper' contiene el número de versión del Schema de Exchange
# Esto cambia cuando la instalación de Exchange extiende el schema de AD
$ExForestRangeUpper = Get-ADObject -Filter "CN -eq 'ms-Exch-Schema-Version-Pt'" `
								   -Searchbase "$($RootDSE.SchemaNamingContext)" `
								   -SearchScope OneLevel `
								   -Property "rangeupper" |
Select-Object -expand rangeupper

# 2. ORGANIZATION VERSION (Configuration partition - Forest-level)
# Query the Exchange Organization container in Configuration partition
# The 'objectVersion' attribute indicates the Exchange organization level
# Updated when running Exchange setup with /PrepareAD
#
# VERSIÓN DE ORGANIZACIÓN (Partición Configuration - Nivel de Forest)
# Consulta el contenedor de Organización de Exchange en la partición Configuration
# El atributo 'objectVersion' indica el nivel de organización de Exchange
# Se actualiza al ejecutar la instalación de Exchange con /PrepareAD
$ExForestObjectVersion = Get-ADObject -Filter "objectClass -eq 'msExchOrganizationContainer'" `
									  -Searchbase "CN=Microsoft Exchange,CN=Services,$($RootDSE.configurationNamingContext)" `
									  -SearchScope OneLevel `
									  -Property "objectVersion" |
Select-Object -expand objectVersion

# 3. DOMAIN VERSION (Domain partition - Domain-level)
# Query the 'Microsoft Exchange System Objects' container in the domain
# The 'objectVersion' attribute indicates the Exchange domain preparation level
# Updated when running Exchange setup with /PrepareDomain
#
# VERSIÓN DE DOMINIO (Partición Domain - Nivel de Dominio)
# Consulta el contenedor 'Microsoft Exchange System Objects' en el dominio
# El atributo 'objectVersion' indica el nivel de preparación del dominio para Exchange
# Se actualiza al ejecutar la instalación de Exchange con /PrepareDomain
$ExDomainObjectVersion = Get-ADObject `
									  -Filter "CN -eq 'Microsoft Exchange System Objects'" `
									  -Searchbase "$($RootDSe.rootDomainNamingContext)" `
									  -SearchScope OneLevel `
									  -Property "objectVersion" |
Select-Object -expand objectVersion
#endregion Query Active Directory for Exchange Version Numbers

#region Build and Display Result
<#
Create a custom object with the version information
Combine the three version numbers and look up the Exchange version name

Crea un objeto personalizado con la información de versión
Combina los tres números de versión y busca el nombre de la versión de Exchange
#>
"" |
Select-Object @{ n = "SchemaType"; e = { "Exchange" } },
	   @{ n = "VersionNumber"; e = { "$ExForestRangeUpper – $ExForestObjectVersion – $ExDomainObjectVersion" } },
	   @{
	n = "VersionName"; e = {
		# Look up the version in the hash table using the combined version string
		# If not found, return "TBD" (To Be Determined)
		# Busca la versión en la tabla hash usando la cadena de versión combinada
		# Si no se encuentra, devuelve "TBD" (To Be Determined)
		if ($ExchSchemaVersions.ContainsKey("$ExForestRangeUpper-$ExForestObjectVersion-$ExDomainObjectVersion")) {
			$ExchSchemaVersions."$ExForestRangeUpper-$ExForestObjectVersion-$ExDomainObjectVersion"
		}
		else {
			"TBD"
		}
	}
}
#endregion Build and Display Result