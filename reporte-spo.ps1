
<#
.SYNOPSIS
    Script de Generación de Reportes de SharePoint Online para Microsoft 365

.DESCRIPTION
    Este script automatiza la generación de reportes comprehensivos sobre la configuración 
    y estado del tenant de SharePoint Online y todos sus sitios asociados. Incluye información 
    detallada sobre configuraciones de uso compartido, almacenamiento, acceso condicional, 
    etiquetas de sensibilidad, integración con Teams y otras características de gobernanza.
    
    El script genera dos archivos de salida:
    1. Reporte del tenant (formato TXT) con configuraciones globales
    2. Reporte de sitios (formato CSV) con información detallada de cada sitio

.FUNCTIONALITY
    - Conexión automática al tenant de SharePoint Online
    - Recopilación de configuraciones del tenant (sharing, storage, conditional access, sensitivity labels)
    - Enumeración y análisis detallado de todos los sitios SharePoint
    - Conversión de cuotas de almacenamiento a formato legible (GB)
    - Cálculo de días desde la última modificación de contenido
    - Traducción de códigos de plantilla a nombres legibles
    - Soporte para interrupción manual (Ctrl+C) preservando datos recolectados
    - Exportación automatizada con timestamp para auditoría

.NOTES
    Archivo:        reporte-spo.ps1
    Autor:          jgparra
    Asistencia:     Claude Sonnet 4 (GitHub Copilot)
    Fecha:          Marzo 2026
    Versión:        1.0
    Requisitos:     - Módulo Microsoft.Online.SharePoint.PowerShell
                    - Permisos de administrador de SharePoint Online
                    - PowerShell 5.1 o superior
    
.EXAMPLE
    .\reporte-spo.ps1
    
    Ejecuta el script con el dominio de tenant configurado en la variable $tenantDomain.
    Genera archivos de reporte con timestamp en el mismo directorio del script.

.LINK
    https://docs.microsoft.com/en-us/powershell/sharepoint/sharepoint-online/connect-sharepoint-online
#>

##################################################################
# CONFIGURACIÓN DE VARIABLES GLOBALES
##################################################################

<#
.VARIABLE tenantDomain
    Dominio del tenant de Microsoft 365 en formato *.onmicrosoft.com
    IMPORTANTE: Reemplazar "contoso" con el nombre de su tenant antes de ejecutar
#>
#$tenantDomain = "contoso.onmicrosoft.com"
$tenantDomain = "contoso.onmicrosoft.com"

##################################################################
# FUNCIONES
##################################################################

<#
.SYNOPSIS
    Establece conexión al servicio de administración de SharePoint Online.

.DESCRIPTION
    Conecta al centro de administración de SharePoint Online utilizando el dominio 
    del tenant proporcionado. Valida la existencia del módulo requerido y construye 
    automáticamente la URL de administración del tenant.

.PARAMETER TenantOnMicrosoftDomain
    Dominio del tenant en formato *.onmicrosoft.com (ejemplo: contoso.onmicrosoft.com)

.EXAMPLE
    Connect-SPOFromTenantDomain -TenantOnMicrosoftDomain "contoso.onmicrosoft.com"
    
    Conecta al centro de administración https://contoso-admin.sharepoint.com

.OUTPUTS
    Ninguno. Establece la sesión de conexión SPO para comandos subsecuentes.

.NOTES
    Requiere el módulo Microsoft.Online.SharePoint.PowerShell instalado.
    Lanza una excepción si el módulo no está disponible.
#>
function Connect-SPOFromTenantDomain {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidatePattern('^[a-zA-Z0-9-]+\.onmicrosoft\.com$')]
        [string]$TenantOnMicrosoftDomain
    )

    # Requiere módulo: Microsoft.Online.SharePoint.PowerShell
    if (-not (Get-Module -ListAvailable -Name Microsoft.Online.SharePoint.PowerShell)) {
        throw "No se encontró el módulo Microsoft.Online.SharePoint.PowerShell. Instálalo primero con: Install-Module Microsoft.Online.SharePoint.PowerShell"
    }

    $tenantName = $TenantOnMicrosoftDomain -replace '\.onmicrosoft\.com$', ''
    $adminUrl   = "https://$tenantName-admin.sharepoint.com"

    Write-Host "Conectando a: $adminUrl" -ForegroundColor Cyan
    Connect-SPOService -Url $adminUrl
}

<#
.SYNOPSIS
    Genera un reporte estructurado de la configuración del tenant de SharePoint Online.

.DESCRIPTION
    Procesa la información del tenant obtenida mediante Get-SPOTenant y crea un objeto 
    personalizado con las propiedades más relevantes organizadas por categorías:
    - Configuraciones de uso compartido (Sharing)
    - Cuotas y almacenamiento (Storage/Quota)
    - Políticas de acceso condicional (Conditional Access)
    - Etiquetas de sensibilidad (Sensitivity Labels)
    
    Convierte valores de almacenamiento de MB a GB para mejor legibilidad.

.PARAMETER tenant
    Objeto de tenant obtenido mediante Get-SPOTenant

.EXAMPLE
    $tenant = Get-SPOTenant
    $report = report-SAMTenant -tenant $tenant
    
    Genera un reporte estructurado del tenant con valores en GB.

.OUTPUTS
    PSCustomObject con propiedades categorizadas del tenant

.NOTES
    Las unidades de almacenamiento se convierten automáticamente a GB con 2 decimales de precisión.
#>
function report-SAMTenant {
    param (
        [Parameter(Mandatory = $true)]
        $tenant
    )
    $report = [pscustomobject]@{
        # Sharing
        SharingCapability                     = $tenant.SharingCapability
        OneDriveSharingCapability             = $tenant.OneDriveSharingCapability
        SharingAllowedDomainList              = if ($tenant.SharingAllowedDomainList) { ($tenant.SharingAllowedDomainList -join '; ') } else { $null }
        SharingBlockedDomainList              = if ($tenant.SharingBlockedDomainList) { ($tenant.SharingBlockedDomainList -join '; ') } else { $null }
        PreventExternalUsersFromResharing     = $tenant.PreventExternalUsersFromResharing
        DefaultSharingLinkType                = $tenant.DefaultSharingLinkType
        DefaultLinkPermission                 = $tenant.DefaultLinkPermission

        # Storage / Quota (convertido a GB)
        StorageQuotaGB                        = [math]::Round($tenant.StorageQuota / 1024, 2)
        StorageQuotaAllocatedGB               = [math]::Round($tenant.StorageQuotaAllocated / 1024, 2)
        BonusStorageQuotaGB                   = [math]::Round($tenant.BonusStorageQuotaMB / 1024, 2)
        ArchivedFileStorageUsageGB            = [math]::Round($tenant.ArchivedFileStorageUsageMB / 1024, 2)
        OneDriveStorageQuotaGB                = [math]::Round($tenant.OneDriveStorageQuota / 1024, 2)
        M365AdditionalStorageSPOEnabled       = $tenant.M365AdditionalStorageSPOEnabled

        # Conditional Access
        ConditionalAccessPolicy               = $tenant.ConditionalAccessPolicy
        ApplyAppEnforcedRestrictionsToGuests  = $tenant.ApplyAppEnforcedRestrictionsToAdHocRecipients
        ReduceTempTokenLifetimeEnabled        = $tenant.ReduceTempTokenLifetimeEnabled

        # Sensitivity Labels
        EnableAIPIntegration                  = $tenant.EnableAIPIntegration
        MarkNewFilesSensitiveByDefault        = $tenant.MarkNewFilesSensitiveByDefault
        EnableSensitivityLabelForPDF          = $tenant.EnableSensitivityLabelForPDF
        EnableSensitivityLabelForOneNote      = $tenant.EnableSensitivityLabelForOneNote
        EnableSensitivityLabelForVideoFiles   = $tenant.EnableSensitivityLabelForVideoFiles
    }

    # Retornar reporte
    return $report
}

<#
.SYNOPSIS
    Convierte códigos de plantilla de sitio SharePoint en nombres legibles.

.DESCRIPTION
    Traduce los códigos técnicos de plantillas de sitio (ejemplo: STS#0, GROUP#0) 
    a nombres descriptivos que facilitan la identificación del tipo de sitio en 
    reportes y auditorías.

.PARAMETER Template
    Código de plantilla de sitio SharePoint (ejemplo: "STS#0", "GROUP#0", "SITEPAGEPUBLISHING#0")

.EXAMPLE
    Get-HumanTemplateName -Template "GROUP#0"
    
    Retorna: "Team site (Microsoft 365 group connected)"

.EXAMPLE
    Get-HumanTemplateName -Template "SITEPAGEPUBLISHING#0"
    
    Retorna: "Communication site"

.OUTPUTS
    String con el nombre descriptivo de la plantilla

.NOTES
    Plantillas no reconocidas retornan "Unknown / Custom" seguido del código original.
#>
function Get-HumanTemplateName {
    param ($Template)

    switch ($Template) {
        "STS#0"   { "Team site (classic)" }
        "STS#3"   { "Team site (no Office 365 group)" }
        "GROUP#0" { "Team site (Microsoft 365 group connected)" }
        "SITEPAGEPUBLISHING#0" { "Communication site" }
        "BDR#0"   { "Document Center" }
        "PROJECTSITE#0" { "Project site" }
        "DEV#0"   { "Developer site" }
        "SRCHCEN#0" { "Enterprise Search Center" }
        "SRCHCENTERLITE#0" { "Basic Search Center" }
        "TENANTADMIN#0" { "Tenant admin site" }
        default   { "Unknown / Custom ($Template)" }
    }
}

<#
.SYNOPSIS
    Genera un reporte detallado de todos los sitios SharePoint Online del tenant.

.DESCRIPTION
    Enumera todos los sitios del tenant y recopila información comprehensiva de cada uno, 
    incluyendo:
    - Información general (URL, título, plantilla, propietario, fechas)
    - Almacenamiento y cuotas (en GB)
    - Control de versiones
    - Configuraciones de uso compartido
    - Políticas de acceso condicional
    - Etiquetas de sensibilidad y barreras de información
    - Configuraciones de Copilot, búsqueda y Loop
    - Integración con Teams y grupos de Microsoft 365
    
    Características adicionales:
    - Barra de progreso visual
    - Soporte para interrupción manual (Ctrl+C) sin pérdida de datos
    - Cálculo automático de días sin modificación
    - Optimización de memoria usando ArrayList

.EXAMPLE
    $sitesReport = Report-SAMSites
    $sitesReport | Export-Csv -Path "sites.csv" -NoTypeInformation
    
    Genera reporte de todos los sitios y lo exporta a CSV.

.OUTPUTS
    System.Collections.ArrayList de objetos PSCustomObject, cada uno representando un sitio 
    con todas sus propiedades detalladas.

.NOTES
    - El proceso puede tardar varios minutos en tenants con muchos sitios
    - Presionar Ctrl+C durante la ejecución detiene la recopilación pero preserva los datos obtenidos
    - Las cuotas de almacenamiento se convierten a GB con hasta 4 decimales
#>
function Report-SAMSites {
    [CmdletBinding()]
    param()

    Write-Host "`nObteniendo lista de sitios..." -ForegroundColor Cyan
    
    $today = Get-Date
    
    # Obtener lista de sitios primero para conocer el total
    $allSites = Get-SPOSite -Limit All
    $totalSites = $allSites.Count
    
    Write-Host "Total de sitios encontrados: $totalSites" -ForegroundColor Yellow
    Write-Host "Procesando sitios individuales para obtener información detallada...`n" -ForegroundColor Cyan
    
    # Usar ArrayList para eficiencia de memoria con grandes cantidades de sitios
    $report = [System.Collections.ArrayList]::new()
    $counter = 0
    $interrupted = $false
    
    # Configurar handler para Ctrl+C
    $originalAction = [Console]::TreatControlCAsInput
    [Console]::TreatControlCAsInput = $true
    
    # Usar GET individual por sitio para obtener valores reales (no del tenant)
    foreach ($siteInfo in $allSites) {
        # Verificar si se presionó Ctrl+C
        if ([Console]::KeyAvailable) {
            $key = [Console]::ReadKey($true)
            if ($key.Key -eq 'C' -and $key.Modifiers -eq 'Control') {
                Write-Host "`n`n⚠️  Captura interrumpida por el usuario (Ctrl+C)." -ForegroundColor Yellow
                Write-Host "   Continuando con los datos recolectados hasta el momento..." -ForegroundColor Yellow
                $interrupted = $true
                break
            }
        }
        
        $counter++
        
        # Mostrar progreso
        $percentComplete = [math]::Round(($counter / $totalSites) * 100, 2)
        Write-Progress -Activity "Procesando sitios SharePoint Online" `
                       -Status "Sitio $counter de $totalSites - $percentComplete% completado" `
                       -CurrentOperation "Procesando: $($siteInfo.Url)" `
                       -PercentComplete $percentComplete
        
        # Obtener información completa del sitio individual
        $site = Get-SPOSite -Identity $siteInfo.Url
        
        # Cálculo de días sin modificación
        $daysLastModify = if ($site.LastContentModifiedDate) {
            (New-TimeSpan -Start $site.LastContentModifiedDate -End $today).Days
        } else {
            $null
        }

        $siteObject = [pscustomobject]@{
            Url                                     = $site.Url
            Title                                   = $site.Title
            Template                                = $site.Template
            TemplateHuman                           = Get-HumanTemplateName $site.Template
            CreatedTime                             = $site.CreatedTime
            LastContentModifiedDate                 = $site.LastContentModifiedDate
            DaysLastModify                          = $daysLastModify
            Status                                  = $site.Status
            Owner                                   = $site.Owner

            # Storage (GB)
            StorageQuotaGB                          = [math]::Round($site.StorageQuota / 1024, 2)
            StorageUsageCurrentGB                   = [math]::Round($site.StorageUsageCurrent / 1024, 4)
            ArchivedFileDiskUsedGB                  = [math]::Round($site.ArchivedFileDiskUsed / 1024, 4)
            StorageQuotaType                        = $site.StorageQuotaType

            # Versioning
            VersionCount                            = $site.VersionCount
            InheritVersionPolicyFromTenant          = $site.InheritVersionPolicyFromTenant
            ExpireVersionsAfterDays                 = $site.ExpireVersionsAfterDays
            EnableAutoExpirationVersionTrim         = $site.EnableAutoExpirationVersionTrim

            # Sharing
            SiteDefinedSharingCapability             = $site.SiteDefinedSharingCapability
            OverrideSharingCapability                = $site.OverrideSharingCapability
            SharingCapability                        = $site.SharingCapability
            SharingDomainRestrictionMode             = $site.SharingDomainRestrictionMode
            SharingAllowedDomainList                 = if ($site.SharingAllowedDomainList) { ($site.SharingAllowedDomainList -join '; ') } else { $null }
            SharingBlockedDomainList                 = if ($site.SharingBlockedDomainList) { ($site.SharingBlockedDomainList -join '; ') } else { $null }
            DisableSharingForNonOwnersStatus         = $site.DisableSharingForNonOwnersStatus

            DefaultSharingLinkType                  = $site.DefaultSharingLinkType
            DefaultLinkPermission                   = $site.DefaultLinkPermission
            DefaultShareLinkScope                   = $site.DefaultShareLinkScope
            DefaultShareLinkRole                    = $site.DefaultShareLinkRole
            DefaultLinkToExistingAccess             = $site.DefaultLinkToExistingAccess

            # Conditional Access
            ConditionalAccessPolicy                 = $site.ConditionalAccessPolicy
            AuthenticationContextName               = $site.AuthenticationContextName
            AuthenticationContextLimitedAccess      = $site.AuthenticationContextLimitedAccess
            AllowDownloadingNonWebViewableFiles     = $site.AllowDownloadingNonWebViewableFiles
            LimitedAccessFileType                   = $site.LimitedAccessFileType

            # Sensitivity / Information Protection
            SensitivityLabel                        = $site.SensitivityLabel
            IsAuthoritative                         = $site.IsAuthoritative
            InformationSegment                      = if ($site.InformationSegment) { ($site.InformationSegment -join '; ') } else { $null }
            InformationBarriersMode                 = $site.InformationBarriersMode

            # Copilot / Search / Loop
            RestrictedContentDiscoveryforCopilotAndAgents = $site.RestrictedContentDiscoveryforCopilotAndAgents
            RestrictContentOrgWideSearch            = $site.RestrictContentOrgWideSearch
            RestrictedAccessControl                 = $site.RestrictedAccessControl
            RestrictedAccessControlGroups           = if ($site.RestrictedAccessControlGroups) { ($site.RestrictedAccessControlGroups -join '; ') } else { $null }
            LoopDefaultSharingLinkScope             = $site.LoopDefaultSharingLinkScope
            MediaTranscription                      = $site.MediaTranscription

            # Teams / Groups
            IsTeamsConnected                        = $site.IsTeamsConnected
            IsTeamsChannelConnected                 = $site.IsTeamsChannelConnected
            RelatedGroupId                          = $site.RelatedGroupId
            GroupId                                 = $site.GroupId
        }
        
        # Agregar al ArrayList (usar [void] para suprimir salida del índice)
        [void]$report.Add($siteObject)
    }
    
    # Restaurar comportamiento original de Ctrl+C
    [Console]::TreatControlCAsInput = $originalAction
    
    # Limpiar barra de progreso
    Write-Progress -Activity "Procesando sitios SharePoint Online" -Completed
    
    Write-Host "`nProcesamiento completado. Total de sitios procesados: $counter" -ForegroundColor Green

    # Retornar reporte
    return $report
}


##################################################################
# BLOQUE PRINCIPAL DE EJECUCIÓN
##################################################################

<#
    Este bloque ejecuta el flujo completo del script:
    1. Conecta al tenant de SharePoint Online
    2. Obtiene información del tenant y genera reporte
    3. Enumera y procesa todos los sitios SharePoint
    4. Exporta resultados en archivos con timestamp
    5. Desconecta del servicio
#>

# Conectar al tenant de SharePoint Online
Connect-SPOFromTenantDomain -TenantOnMicrosoftDomain $tenantDomain

# Obtener información del tenant
$tenantrpt = Get-SPOTenant

# Crear reporte personalizado con valores en GB para mejor legibilidad
Write-Host "Generando reporte del tenant..." -ForegroundColor Green
$reportTenant = report-SAMTenant -tenant $tenantrpt

# Preparar nombres de archivo con timestamp
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$tenantName = $tenantDomain -replace '\.onmicrosoft\.com$', ''
$timestamp = Get-Date -Format "yyyy-MM-dd--HH-mm"
$fileNameTenant = "$tenantName`_tenant_$timestamp.txt"
$fileNameSites = "$tenantName`_sites_$timestamp.csv"
$filePathTenant = Join-Path $scriptPath $fileNameTenant
$filePathSites = Join-Path $scriptPath $fileNameSites

# Exportar reporte del tenant como archivo de texto formateado
Write-Host "`nGuardando reporte del tenant en: $filePathTenant" -ForegroundColor Cyan
$reportTenant | Format-List | Out-String | Out-File -FilePath $filePathTenant -Encoding UTF8
Write-Host "Reporte del tenant guardado exitosamente." -ForegroundColor Green

# Obtener información detallada de todos los sitios
Write-Host "`nGenerando reporte de sitios..." -ForegroundColor Green
$reportSites = Report-SAMSites

# Exportar reporte de sitios como CSV para análisis en Excel
Write-Host "`nGuardando reporte de sitios en: $filePathSites" -ForegroundColor Cyan
$reportSites | Export-Csv -Path $filePathSites -NoTypeInformation -Encoding UTF8
Write-Host "Reporte de sitios guardado exitosamente." -ForegroundColor Green

# Mostrar resumen de archivos generados
Write-Host "`n=== RESUMEN ==="-ForegroundColor Yellow
Write-Host "Archivos generados:" -ForegroundColor Yellow
Write-Host "  - Tenant: $filePathTenant" -ForegroundColor White
Write-Host "  - Sitios: $filePathSites ($($reportSites.Count) sitios)" -ForegroundColor White

# Cerrar sesión de SharePoint Online
Write-Host "`nDesconectando del tenant SharePoint Online..." -ForegroundColor Cyan
Disconnect-SPOService

##################################################################
# FIN DEL SCRIPT
##################################################################