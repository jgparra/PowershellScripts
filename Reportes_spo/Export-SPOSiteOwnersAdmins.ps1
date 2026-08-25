<#
.SYNOPSIS
    Exporta a CSV los administradores del sitio (Site Collection Admins) y los
    miembros del grupo Owners de cada site collection de SharePoint Online.

.DESCRIPTION
    Recorre todas (o un subconjunto) de las site collections del tenant y, por cada
    una, recolecta dos conjuntos de identidades:
      - Site Collection Administrators (los "admins" del sitio).
      - Miembros del grupo Owners asociado del sitio.

    Cuando un administrador u owner es un grupo de seguridad de Entra ID o un grupo
    de Microsoft 365, el script lo expande a sus miembros individuales usando
    Microsoft Graph (con caché para no repetir consultas de un mismo grupo).

    El resultado se exporta a un CSV con una fila por sitio. Como una celda puede
    contener varios usuarios, los UPN se unen con el separador pipe "|" para no
    confundirse con la coma o el punto y coma que delimitan las columnas del CSV.

    Es un script de SOLO LECTURA: no modifica permisos ni pertenencias de grupo.

    Flujo:
    1. Valida los módulos requeridos (PnP.PowerShell y Microsoft Graph).
    2. Conecta a Microsoft Graph (para expandir grupos) y al admin de SharePoint.
    3. Enumera site collections (todas o las que coincidan con -SiteFilter).
    4. Por cada sitio obtiene admins (Get-PnPSiteCollectionAdmin) y owners
       (grupo Owners asociado) y resuelve/expande cada identidad a UPN.
    5. Exporta un CSV con timestamp, con Admins y Owners separados por "|".

.PARAMETER TenantAdminUrl
    URL del centro de administración de SharePoint Online
    (ejemplo: https://contoso-admin.sharepoint.com). Requerido para enumerar los
    sitios del tenant.

.PARAMETER SiteFilter
    Patrón opcional (comodines) para limitar los sitios a procesar por su URL.
    Ejemplo: "*/sites/finanzas*". Si se omite, se procesan todos los sitios.

.PARAMETER IncludeOneDrive
    Incluye los sitios personales de OneDrive for Business en el análisis.
    Por defecto se excluyen (URLs que contienen "-my.sharepoint.com").

.PARAMETER ExpandAdminGroups
    Si un Site Collection Administrator es un grupo, lo expande a sus miembros
    individuales (igual que se hace siempre con los Owners). Activado por defecto.
    Usa -ExpandAdminGroups:$false para listar el grupo tal cual.

.PARAMETER OutputPath
    Carpeta donde guardar el CSV de resultados. Por defecto, la carpeta del script.

.PARAMETER ClientId
    (Opcional) Application (Client) ID de una app registrada en Entra ID para la
    conexión de PnP. Recomendado desde PnP.PowerShell 2.x, que ya no incluye una
    app multi-tenant por defecto.

.EXAMPLE
    .\Export-SPOSiteOwnersAdmins.ps1 -TenantAdminUrl "https://contoso-admin.sharepoint.com"

    Recorre todos los sitios y genera un CSV con admins y owners de cada uno.

.EXAMPLE
    .\Export-SPOSiteOwnersAdmins.ps1 -TenantAdminUrl "https://contoso-admin.sharepoint.com" `
        -SiteFilter "*/sites/finanzas*"

    Solo analiza los sitios cuya URL coincide con el patrón indicado (ideal para probar).

.NOTES
    Archivo:        Export-SPOSiteOwnersAdmins.ps1
    Autor:          jgparra
    Asistencia:     Claude Opus 4.8 (GitHub Copilot)
    Fecha:          Agosto 2026
    Versión:        1.0
    Requisitos:     - Módulo PnP.PowerShell
                    - Módulo Microsoft.Graph.Groups y Microsoft.Graph.Users
                      (o Microsoft.Graph)
                    - Rol de administrador de SharePoint para enumerar sitios del tenant
                    - Permisos de Graph: User.Read.All y GroupMember.Read.All (delegados)
                    - PowerShell 7.x recomendado (PnP.PowerShell 2.x)

.LINK
    https://pnp.github.io/powershell/cmdlets/Get-PnPSiteCollectionAdmin.html
.LINK
    https://pnp.github.io/powershell/cmdlets/Get-PnPGroupMember.html
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidatePattern('^https://[a-zA-Z0-9-]+-admin\.sharepoint\.com/?$')]
    [string]$TenantAdminUrl,

    [Parameter(Mandatory = $false)]
    [string]$SiteFilter,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeOneDrive,

    [Parameter(Mandatory = $false)]
    [switch]$ExpandAdminGroups = $true,

    [Parameter(Mandatory = $false)]
    [string]$OutputPath,

    [Parameter(Mandatory = $false)]
    [string]$ClientId
)

##################################################################
# FUNCIONES
##################################################################

<#
.SYNOPSIS
    Valida que los módulos requeridos estén instalados.

.OUTPUTS
    Ninguno. Lanza excepción si falta un módulo requerido.
#>
function Test-RequiredModules {
    [CmdletBinding()]
    param()

    if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
        throw "No se encontró el módulo PnP.PowerShell. Instálalo con: Install-Module PnP.PowerShell -Scope CurrentUser"
    }

    $graphAvailable = (Get-Module -ListAvailable -Name Microsoft.Graph.Groups) -or `
                      (Get-Module -ListAvailable -Name Microsoft.Graph)
    if (-not $graphAvailable) {
        throw "No se encontró el módulo de Microsoft Graph. Instálalo con: Install-Module Microsoft.Graph.Groups -Scope CurrentUser"
    }
}

<#
.SYNOPSIS
    Expande un grupo de Entra ID (por su object id) a los UPN de sus miembros.

.DESCRIPTION
    Usa Get-MgGroupTransitiveMember para obtener todos los usuarios (incluyendo los
    de subgrupos anidados). Solo devuelve principales de tipo usuario; ignora otros
    tipos de directoryObject. Los resultados se cachean por object id.

.PARAMETER GroupId
    Object ID (GUID) del grupo en Entra ID.

.PARAMETER Cache
    Tabla hash usada como caché (GroupId -> string[] de UPN).

.OUTPUTS
    [string[]] con los UPN de los miembros usuario del grupo.
#>
function Expand-GraphGroupMembers {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$GroupId,

        [Parameter(Mandatory = $true)]
        [hashtable]$Cache
    )

    if ($Cache.ContainsKey($GroupId)) {
        return $Cache[$GroupId]
    }

    $upns = [System.Collections.Generic.List[string]]::new()
    try {
        $members = Get-MgGroupTransitiveMember -GroupId $GroupId -All -ErrorAction Stop
        foreach ($m in $members) {
            # Solo nos interesan los miembros de tipo usuario
            $type = $m.AdditionalProperties['@odata.type']
            if ($type -eq '#microsoft.graph.user') {
                $upn = $m.AdditionalProperties['userPrincipalName']
                if ($upn) { $upns.Add([string]$upn) }
            }
        }
    }
    catch {
        Write-Warning "No se pudo expandir el grupo '$GroupId' en Graph: $($_.Exception.Message)"
    }

    $result = $upns.ToArray()
    $Cache[$GroupId] = $result
    return $result
}

<#
.SYNOPSIS
    Resuelve un principal de SharePoint (por su LoginName) a uno o varios UPN.

.DESCRIPTION
    Interpreta el claim del LoginName:
      - Usuario federado de Entra ID (i:0#.f|membership|upn) -> devuelve el UPN.
      - Grupo de seguridad (c:0t.c|tenant|<guid>) o grupo de Microsoft 365
        (c:0o.c|federateddirectoryclaimprovider|<guid>[_o]) -> si $Expand está
        activo, expande a los UPN de sus miembros vía Graph; si no, devuelve el
        propio Title/LoginName como etiqueta del grupo.
    Ignora cuentas de sistema, app principals y otros claims no relevantes.

.PARAMETER LoginName
    LoginName del principal (de Get-PnPSiteCollectionAdmin o Get-PnPGroupMember).

.PARAMETER Title
    Título del principal, usado como etiqueta cuando no se expande el grupo.

.PARAMETER Expand
    Si es $true, los grupos se expanden a sus miembros usuario vía Graph.

.PARAMETER GroupCache
    Caché de expansión de grupos (GroupId -> string[] de UPN).

.OUTPUTS
    [string[]] con los UPN (o etiquetas de grupo) resueltos.
#>
function Resolve-PrincipalToUpns {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$LoginName,

        [Parameter(Mandatory = $false)]
        [string]$Title,

        [Parameter(Mandatory = $true)]
        [bool]$Expand,

        [Parameter(Mandatory = $true)]
        [hashtable]$GroupCache
    )

    # Usuario federado de Entra ID
    if ($LoginName -match '^i:0#\.f\|membership\|(.+)$') {
        return , @($Matches[1])
    }

    # Grupo de seguridad de Entra ID: c:0t.c|tenant|<guid>
    if ($LoginName -match '^c:0t\.c\|tenant\|([0-9a-fA-F-]{36})$') {
        $guid = $Matches[1]
        if ($Expand) { return , (Expand-GraphGroupMembers -GroupId $guid -Cache $GroupCache) }
        return , @($(if ($Title) { $Title } else { $LoginName }))
    }

    # Grupo de Microsoft 365: c:0o.c|federateddirectoryclaimprovider|<guid>[_o]
    if ($LoginName -match '^c:0o\.c\|federateddirectoryclaimprovider\|([0-9a-fA-F-]{36})(?:_o)?$') {
        $guid = $Matches[1]
        if ($Expand) { return , (Expand-GraphGroupMembers -GroupId $guid -Cache $GroupCache) }
        return , @($(if ($Title) { $Title } else { $LoginName }))
    }

    # Otros claims (sistema, app principals, SharePoint groups sin UPN): se ignoran
    return , @()
}

<#
.SYNOPSIS
    Recolecta admins y owners de una site collection y los devuelve como UPN.

.DESCRIPTION
    Conecta al sitio indicado, obtiene los Site Collection Administrators y los
    miembros del grupo Owners asociado, resuelve cada principal a UPN (expandiendo
    grupos según corresponda) y devuelve un objeto con ambas listas deduplicadas.

.PARAMETER SiteUrl
    URL de la site collection a procesar.

.PARAMETER ExpandAdmins
    Si $true, los administradores que sean grupos también se expanden.

.PARAMETER GroupCache
    Caché de expansión de grupos compartida entre sitios.

.PARAMETER ClientId
    (Opcional) Client ID de la app de Entra ID para la conexión PnP.

.OUTPUTS
    PSCustomObject con Admins (string[]) y Owners (string[]), o $null si falla.
#>
function Get-SiteOwnersAndAdmins {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SiteUrl,

        [Parameter(Mandatory = $true)]
        [bool]$ExpandAdmins,

        [Parameter(Mandatory = $true)]
        [hashtable]$GroupCache,

        [Parameter(Mandatory = $false)]
        [string]$ClientId
    )

    $connectParams = @{
        Url         = $SiteUrl
        Interactive = $true
        ErrorAction = 'Stop'
    }
    if ($ClientId) { $connectParams['ClientId'] = $ClientId }

    try {
        Connect-PnPOnline @connectParams
    }
    catch {
        Write-Warning "No se pudo conectar a '$SiteUrl': $($_.Exception.Message)"
        return $null
    }

    $admins = [System.Collections.Generic.List[string]]::new()
    $owners = [System.Collections.Generic.List[string]]::new()

    # --- Site Collection Administrators ---
    try {
        $siteAdmins = Get-PnPSiteCollectionAdmin -ErrorAction Stop
        foreach ($a in $siteAdmins) {
            $resolved = Resolve-PrincipalToUpns -LoginName $a.LoginName -Title $a.Title `
                            -Expand $ExpandAdmins -GroupCache $GroupCache
            foreach ($u in $resolved) { $admins.Add($u) }
        }
    }
    catch {
        Write-Warning "No se pudieron obtener admins de '$SiteUrl': $($_.Exception.Message)"
    }

    # --- Grupo Owners asociado ---
    try {
        $ownerGroup = Get-PnPGroup -AssociatedOwnerGroup -ErrorAction Stop
        if ($ownerGroup) {
            $ownerMembers = Get-PnPGroupMember -Group $ownerGroup -ErrorAction Stop
            foreach ($m in $ownerMembers) {
                $resolved = Resolve-PrincipalToUpns -LoginName $m.LoginName -Title $m.Title `
                                -Expand $true -GroupCache $GroupCache
                foreach ($u in $resolved) { $owners.Add($u) }
            }
        }
    }
    catch {
        Write-Warning "No se pudo obtener el grupo Owners de '$SiteUrl': $($_.Exception.Message)"
    }

    # Deduplicar sin distinguir mayúsculas/minúsculas
    $adminsUnique = $admins | Sort-Object -Unique
    $ownersUnique = $owners | Sort-Object -Unique

    return [pscustomobject]@{
        Admins = @($adminsUnique)
        Owners = @($ownersUnique)
    }
}

##################################################################
# BLOQUE PRINCIPAL DE EJECUCIÓN
##################################################################

# 1. Validar módulos
Test-RequiredModules

# Resolver carpeta de salida
if (-not $OutputPath) {
    $OutputPath = Split-Path -Parent $MyInvocation.MyCommand.Path
}
if (-not (Test-Path -Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
}

# 2. Conectar a Microsoft Graph (solo lectura para expandir grupos)
Write-Host "Conectando a Microsoft Graph..." -ForegroundColor Cyan
Connect-MgGraph -Scopes "User.Read.All", "GroupMember.Read.All" -NoWelcome

# 3. Conectar al admin de SharePoint y enumerar sitios
Write-Host "Conectando al centro de administración de SharePoint: $TenantAdminUrl" -ForegroundColor Cyan
$adminConnectParams = @{
    Url         = $TenantAdminUrl
    Interactive = $true
    ErrorAction = 'Stop'
}
if ($ClientId) { $adminConnectParams['ClientId'] = $ClientId }
Connect-PnPOnline @adminConnectParams

Write-Host "Obteniendo lista de sitios del tenant..." -ForegroundColor Cyan
$allSites = Get-PnPTenantSite

# Filtrar OneDrive salvo que se pida incluirlo
if (-not $IncludeOneDrive) {
    $allSites = $allSites | Where-Object { $_.Url -notmatch '-my\.sharepoint\.com' }
}

# Aplicar filtro de URL opcional
if ($SiteFilter) {
    $allSites = $allSites | Where-Object { $_.Url -like $SiteFilter }
}

$totalSites = @($allSites).Count
Write-Host "Sitios a procesar: $totalSites" -ForegroundColor Yellow

if ($totalSites -eq 0) {
    Write-Host "No hay sitios que coincidan con los criterios. Finalizando." -ForegroundColor Yellow
    Disconnect-PnPOnline
    Disconnect-MgGraph | Out-Null
    return
}

# 4. Procesar cada sitio
$report      = [System.Collections.ArrayList]::new()
$groupCache  = @{}   # Caché de expansión de grupos (GroupId -> UPN[])
$counter     = 0
$interrupted = $false

# Configurar handler para Ctrl+C sin perder los datos ya recolectados
$originalAction = [Console]::TreatControlCAsInput
[Console]::TreatControlCAsInput = $true

foreach ($site in $allSites) {
    # Verificar interrupción manual (Ctrl+C)
    if ([Console]::KeyAvailable) {
        $key = [Console]::ReadKey($true)
        if ($key.Key -eq 'C' -and $key.Modifiers -eq 'Control') {
            Write-Host "`n`nCaptura interrumpida por el usuario (Ctrl+C)." -ForegroundColor Yellow
            Write-Host "   Continuando con los datos recolectados hasta el momento..." -ForegroundColor Yellow
            $interrupted = $true
            break
        }
    }

    $counter++
    $percentComplete = [math]::Round(($counter / $totalSites) * 100, 2)
    Write-Progress -Activity "Analizando site collections" `
                   -Status "Sitio $counter de $totalSites - $percentComplete% completado" `
                   -CurrentOperation "Procesando: $($site.Url)" `
                   -PercentComplete $percentComplete

    $data = Get-SiteOwnersAndAdmins -SiteUrl $site.Url -ExpandAdmins $ExpandAdminGroups `
                -GroupCache $groupCache -ClientId $ClientId
    if (-not $data) { continue }

    $row = [pscustomobject]@{
        SiteUrl     = $site.Url
        Title       = $site.Title
        Template    = $site.Template
        Admins      = ($data.Admins -join '|')
        AdminsCount = $data.Admins.Count
        Owners      = ($data.Owners -join '|')
        OwnersCount = $data.Owners.Count
    }
    [void]$report.Add($row)
}

# Restaurar comportamiento original de Ctrl+C
[Console]::TreatControlCAsInput = $originalAction
Write-Progress -Activity "Analizando site collections" -Completed

Write-Host "`nSitios procesados: $counter de $totalSites" -ForegroundColor Green

# 5. Exportar CSV de resultados
$timestamp = Get-Date -Format "yyyy-MM-dd--HH-mm"
$csvName   = "SPO_SiteOwnersAdmins_$timestamp.csv"
$csvPath   = Join-Path $OutputPath $csvName

if ($report.Count -gt 0) {
    $report | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
    Write-Host "`nReporte guardado en: $csvPath" -ForegroundColor Cyan
}
else {
    Write-Host "`nNo se generaron filas. No se crea CSV." -ForegroundColor Yellow
}

# Resumen
Write-Host "`n=== RESUMEN ===" -ForegroundColor Yellow
Write-Host "  Sitios procesados:   $counter de $totalSites" -ForegroundColor White
Write-Host "  Filas en el reporte: $($report.Count)" -ForegroundColor White
if ($interrupted) {
    Write-Host "  NOTA: ejecución interrumpida manualmente (datos parciales)." -ForegroundColor Yellow
}

# Desconectar servicios
Write-Host "`nDesconectando de SharePoint Online y Microsoft Graph..." -ForegroundColor Cyan
Disconnect-PnPOnline
Disconnect-MgGraph | Out-Null
