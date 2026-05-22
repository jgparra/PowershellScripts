# SharePoint Online Reports / Reportes de SharePoint Online

Comprehensive SharePoint Online tenant and site inventory reporting for Microsoft 365.

Reportes comprehensivos de tenant y sitios de SharePoint Online para Microsoft 365.

---

## 📄 Script: reporte-spo.ps1

### Description / Descripción

**English:**

PowerShell script to generate comprehensive **SharePoint Online** reports in Microsoft 365. Automates the collection of tenant configurations and detailed inventory of all sites, exporting results to audit-ready files.

Connects to SharePoint Online admin center, collects global tenant configuration, and processes each site individually to extract detailed properties.

**Español:**

Script de PowerShell para generar reportes comprehensivos de **SharePoint Online** en Microsoft 365. Automatiza la recopilación de configuraciones del tenant y el inventario detallado de todos los sitios, exportando los resultados en archivos listos para auditoría o análisis.

El script conecta al centro de administración de SharePoint Online, recopila la configuración global del tenant y procesa cada sitio individualmente para extraer propiedades detalladas.

---

## 📤 Output Files / Archivos de Salida

| File / Archivo | Format / Formato | Content / Contenido |
|---|---|---|
| `<tenant>_tenant_<timestamp>.txt` | TXT | Global tenant configuration / Configuración global del tenant |
| `<tenant>_sites_<timestamp>.csv` | CSV | Detailed site inventory / Inventario detallado de todos los sitios |

---

## 📋 Prerequisites / Requisitos

**English:**
- **Windows PowerShell** 5.1 or higher **(currently does not work well on PowerShell 7.0+)**
- **Module:** `Microsoft.Online.SharePoint.PowerShell`
- **Permissions:** SharePoint Online Administrator role

**Español:**
- **Windows PowerShell** 5.1 o superior **(por el momento no se ejecuta bien en PowerShell 7.0+)**
- **Módulo:** `Microsoft.Online.SharePoint.PowerShell`
- **Permisos:** Administrador de SharePoint Online (rol *SharePoint Administrator*)

### Module Installation / Instalación del Módulo

```powershell
Install-Module Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser
```

---

## ⚙️ Configuration / Configuración

**English:**
Before running, edit the `$tenantDomain` variable in the configuration section with your actual tenant domain:

**Español:**
Antes de ejecutar, editar la variable `$tenantDomain` en la sección de configuración con el dominio real del tenant:

```powershell
$tenantDomain = "contoso.onmicrosoft.com"
```

---

## 🚀 Usage / Uso

```powershell
.\reporte-spo.ps1
```

**English:** The script will prompt for authentication when connecting to the tenant. Output files are generated in the same directory as the script.

**Español:** El script solicitará autenticación al conectarse al tenant. Los archivos de salida se generan en el mismo directorio del script.

---

## 📊 Collected Data / Datos Recopilados

> **Note:** The detailed technical documentation below is provided in Spanish due to its extensive nature. It includes comprehensive field descriptions for both tenant configuration and site inventory reports.  
> **Nota:** La documentación técnica detallada a continuación está en español e incluye descripciones comprehensivas de campos para reportes de configuración del tenant e inventario de sitios.

---

### Reporte del tenant (`_tenant_*.txt`)

Contiene la configuración global del tenant, independiente de cada sitio individual.

#### Sharing (uso compartido)

| Campo | Descripción |
|---|---|
| `SharingCapability` | Nivel de uso compartido externo permitido para SharePoint Online a nivel de tenant. Valores: `Disabled`, `ExistingExternalUserSharingOnly`, `ExternalUserSharingOnly`, `ExternalUserAndGuestSharing`. |
| `OneDriveSharingCapability` | Igual que `SharingCapability` pero aplicado específicamente a OneDrive. Puede ser más restrictivo que el nivel SPO. |
| `SharingAllowedDomainList` | Lista de dominios externos habilitados para compartir (cuando se usa modo de lista de permitidos). Separados por punto y coma. |
| `SharingBlockedDomainList` | Lista de dominios externos bloqueados para compartir (cuando se usa modo de lista de bloqueados). Separados por punto y coma. |
| `PreventExternalUsersFromResharing` | Si es `True`, los usuarios externos no pueden recompartir contenido que les fue compartido. |
| `DefaultSharingLinkType` | Tipo de enlace predeterminado al compartir. Valores: `None`, `Direct`, `Internal`, `AnonymousAccess`. |
| `DefaultLinkPermission` | Permiso predeterminado del enlace al compartir. Valores: `None`, `View`, `Edit`. |

#### Almacenamiento / Cuotas

| Campo | Descripción |
|---|---|
| `StorageQuotaGB` | Cuota de almacenamiento total asignada al tenant de SharePoint (en GB). |
| `StorageQuotaAllocatedGB` | Almacenamiento total ya asignado a los sitios del tenant (en GB). |
| `BonusStorageQuotaGB` | Almacenamiento adicional adquirido como complemento (en GB). |
| `ArchivedFileStorageUsageGB` | Espacio ocupado por archivos en estado archivado (Microsoft 365 Archive, en GB). |
| `OneDriveStorageQuotaGB` | Cuota de almacenamiento predeterminada asignada a cada OneDrive de usuario (en GB). |
| `M365AdditionalStorageSPOEnabled` | Indica si el complemento de almacenamiento adicional de Microsoft 365 está habilitado para SPO. |

#### Acceso condicional

| Campo | Descripción |
|---|---|
| `ConditionalAccessPolicy` | Política de acceso condicional aplicada al tenant. Valores: `AllowFullAccess`, `AllowLimitedAccess`, `BlockAccess`. |
| `ApplyAppEnforcedRestrictionsToGuests` | Si es `True`, las restricciones de acceso condicional impuestas por la app se aplican también a destinatarios ad hoc (invitados). |
| `ReduceTempTokenLifetimeEnabled` | Si es `True`, se reduce la duración de los tokens temporales para acceso a archivos, aumentando la seguridad en dispositivos no administrados. |

#### Etiquetas de sensibilidad

| Campo | Descripción |
|---|---|
| `EnableAIPIntegration` | Habilita la integración con Microsoft Purview Information Protection (antes AIP) para clasificar y proteger contenido en SPO. |
| `MarkNewFilesSensitiveByDefault` | Si está habilitado, los archivos nuevos se marcan como sensibles por defecto hasta que sean analizados por las políticas de DLP. |
| `EnableSensitivityLabelForPDF` | Permite aplicar etiquetas de sensibilidad a archivos PDF en SharePoint y OneDrive. |
| `EnableSensitivityLabelForOneNote` | Permite aplicar etiquetas de sensibilidad a blocs de notas de OneNote. |
| `EnableSensitivityLabelForVideoFiles` | Permite aplicar etiquetas de sensibilidad a archivos de vídeo almacenados en SPO. |

---

### Reporte de sitios (`_sites_*.csv`)

Contiene una fila por cada sitio del tenant. Los valores son los configurados específicamente en cada sitio, consultados individualmente mediante `Get-SPOSite -Identity`.

#### General

| Campo | Descripción |
|---|---|
| `Url` | URL completa del sitio (ej. `https://contoso.sharepoint.com/sites/marketing`). |
| `Title` | Nombre visible del sitio. |
| `Template` | Código de plantilla interno del sitio (ej. `GROUP#0`, `SITEPAGEPUBLISHING#0`). |
| `TemplateHuman` | Nombre legible de la plantilla, traducido por el script (ej. `Team site (Microsoft 365 group connected)`). |
| `CreatedTime` | Fecha y hora de creación del sitio. |
| `LastContentModifiedDate` | Fecha y hora de la última modificación de contenido registrada en el sitio. |
| `DaysLastModify` | Número de días transcurridos desde `LastContentModifiedDate` hasta la fecha de ejecución del script. Útil para identificar sitios inactivos. |
| `Status` | Estado actual del sitio. Valores típicos: `Active`, `Recycled`. |
| `Owner` | Cuenta del propietario principal del sitio (UPN). |

#### Almacenamiento

| Campo | Descripción |
|---|---|
| `StorageQuotaGB` | Cuota de almacenamiento máxima asignada al sitio (en GB). `0` indica que hereda la cuota del tenant. |
| `StorageUsageCurrentGB` | Espacio de almacenamiento actualmente utilizado por el sitio (en GB, hasta 4 decimales). |
| `ArchivedFileDiskUsedGB` | Espacio ocupado por archivos archivados en este sitio mediante Microsoft 365 Archive (en GB). |
| `StorageQuotaType` | Tipo de cuota: `PooledSitesDefault` (hereda del pool del tenant) o `Sandboxed` (cuota individual asignada). |

#### Control de versiones

| Campo | Descripción |
|---|---|
| `VersionCount` | Número máximo de versiones de archivos que se conservan en las bibliotecas del sitio. |
| `InheritVersionPolicyFromTenant` | Si es `True`, el sitio aplica la política de versiones del tenant en lugar de una configuración propia. |
| `ExpireVersionsAfterDays` | Número de días tras los cuales las versiones antiguas de archivos expiran y se eliminan automáticamente. `0` significa sin expiración. |
| `EnableAutoExpirationVersionTrim` | Cuando es `True`, el sistema gestiona automáticamente la eliminación de versiones según criterios de Microsoft (sin límite fijo de versiones). |

#### Uso compartido (Sharing)

| Campo | Descripción |
|---|---|
| `SiteDefinedSharingCapability` | Configuración de sharing definida directamente en el sitio, antes de aplicar las restricciones del tenant. |
| `OverrideSharingCapability` | Indica si el sitio tiene permitido sobreescribir la política de sharing del tenant. |
| `SharingCapability` | Nivel de sharing efectivo del sitio (resultado de combinar la configuración del sitio con el límite del tenant). |
| `SharingDomainRestrictionMode` | Modo de restricción de dominios: `None`, `AllowList` (solo dominios permitidos) o `BlockList` (todos excepto bloqueados). |
| `SharingAllowedDomainList` | Dominios externos habilitados para compartir en este sitio (solo aplica cuando `SharingDomainRestrictionMode = AllowList`). |
| `SharingBlockedDomainList` | Dominios externos bloqueados en este sitio (solo aplica cuando `SharingDomainRestrictionMode = BlockList`). |
| `DisableSharingForNonOwnersStatus` | Si está activado, solo los propietarios del sitio pueden compartir contenido con usuarios externos. |
| `DefaultSharingLinkType` | Tipo de enlace predeterminado al compartir desde este sitio. Sobreescribe el valor del tenant si está configurado. |
| `DefaultLinkPermission` | Permiso predeterminado del enlace generado en este sitio (`View` o `Edit`). |
| `DefaultShareLinkScope` | Ámbito predeterminado del enlace al compartir: `Anyone`, `Organization`, `SpecificPeople`, `Uninitialized`. |
| `DefaultShareLinkRole` | Rol predeterminado del enlace: `None`, `View`, `Edit`, `ManageList`, `Owner`, `RestrictedView`. |
| `DefaultLinkToExistingAccess` | Si es `True`, el enlace predeterminado otorga acceso solo a personas que ya tienen permisos en el sitio. |

#### Acceso condicional

| Campo | Descripción |
|---|---|
| `ConditionalAccessPolicy` | Política de acceso condicional aplicada a este sitio específicamente. Puede ser más restrictiva que la del tenant. Valores: `AllowFullAccess`, `AllowLimitedAccess`, `BlockAccess`. |
| `AuthenticationContextName` | Nombre del contexto de autenticación de Azure AD Conditional Access configurado para el sitio (requiere acceso escalonado). |
| `AuthenticationContextLimitedAccess` | Si es `True`, se aplica la restricción de acceso condicional mediante el contexto de autenticación definido. |
| `AllowDownloadingNonWebViewableFiles` | Cuando el sitio tiene acceso limitado (solo web), controla si se permite la descarga de archivos que no pueden visualizarse en el navegador. |
| `LimitedAccessFileType` | Define qué archivos son accesibles en modo de acceso limitado: `OfficeOnlineFilesOnly` (solo Office en web) o `WebPreviewableFiles` (todos los que tienen vista previa). |

#### Sensibilidad / Barreras de información

| Campo | Descripción |
|---|---|
| `SensitivityLabel` | GUID de la etiqueta de sensibilidad de Microsoft Purview aplicada al sitio. Vacío si no tiene etiqueta. |
| `IsAuthoritative` | Indica si el sitio está marcado como fuente autoritativa de contenido para búsquedas (Authoritative Page en SharePoint Search). |
| `InformationSegment` | Lista de segmentos de barreras de información (Information Barriers) asignados al sitio. Separados por punto y coma. |
| `InformationBarriersMode` | Modo de barreras de información del sitio. Valores: `Open`, `Explicit`, `Implicit`, `OwnerModerated`, `Disabled`. |

#### Copilot / Búsqueda / Loop

| Campo | Descripción |
|---|---|
| `RestrictedContentDiscoveryforCopilotAndAgents` | Si es `True`, el contenido del sitio no será descubierto ni utilizado por Copilot para Microsoft 365 ni por agentes de IA. |
| `RestrictContentOrgWideSearch` | Si es `True`, el contenido del sitio se excluye de los resultados de búsqueda para toda la organización (útil para sitios con información sensible). |
| `RestrictedAccessControl` | Indica si el sitio tiene habilitado el control de acceso restringido, que limita el acceso únicamente a grupos de seguridad específicos incluso si existen otros permisos. |
| `RestrictedAccessControlGroups` | Lista de grupos de seguridad autorizados cuando `RestrictedAccessControl` está habilitado. Separados por punto y coma. |
| `LoopDefaultSharingLinkScope` | Ámbito predeterminado de los enlaces de compartición para componentes de Microsoft Loop alojados en este sitio. |
| `MediaTranscription` | Estado de la transcripción automática de medios (vídeo/audio) en el sitio: `Enabled` o `Disabled`. |

#### Teams / Grupos de Microsoft 365

| Campo | Descripción |
|---|---|
| `IsTeamsConnected` | `True` si el sitio está conectado a un equipo de Microsoft Teams (Teams estándar basado en canal General). |
| `IsTeamsChannelConnected` | `True` si el sitio corresponde a un canal privado o compartido de Microsoft Teams. |
| `RelatedGroupId` | GUID del grupo de Microsoft 365 asociado al sitio (si lo tiene). Permite cruzar datos con Entra ID / Exchange. |
| `GroupId` | GUID del grupo interno de SharePoint. Puede diferir de `RelatedGroupId` en sitios de canal. |

---

## ✨ Features / Características

- **Progress bar** / **Barra de progreso** con porcentaje completado durante el procesamiento de sitios
- **Safe interruption** / **Interrupción segura** con `Ctrl+C`: detiene la recopilación sin perder los datos ya procesados
- **GB quotas** / **Cuotas en GB** con conversión automática desde MB para mejor legibilidad
- **Readable template names** / **Nombres de plantilla legibles**: traduce códigos como `GROUP#0` a `Team site (Microsoft 365 group connected)`
- **Timestamp** en los nombres de archivo para facilitar el versionado histórico de reportes

### Recognized Site Templates / Plantillas de Sitio Reconocidas

| Código | Descripción |
|---|---|
| `STS#0` | Team site (classic) |
| `STS#3` | Team site (no Office 365 group) |
| `GROUP#0` | Team site (Microsoft 365 group connected) |
| `SITEPAGEPUBLISHING#0` | Communication site |
| `BDR#0` | Document Center |
| `PROJECTSITE#0` | Project site |
| `SRCHCEN#0` | Enterprise Search Center |
| `TENANTADMIN#0` | Tenant admin site |

---

## 📝 Example Output / Ejemplo de Salida

```
=== RESUMEN ===
Archivos generados:
  - Tenant: C:\...\contoso_tenant_2026-04-27--10-30.txt
  - Sitios:  C:\...\contoso_sites_2026-04-27--10-30.csv (342 sitios)
```

---

## ⚠️ Notes / Notas

**English:**
- Processing may take several minutes on tenants with many sites, as it queries each site individually to obtain actual values rather than inherited values from the tenant
- The script automatically disconnects the SPO session when finished (`Disconnect-SPOService`)

**Español:**
- El procesamiento puede tardar varios minutos en tenants con muchos sitios, ya que consulta cada sitio individualmente para obtener sus valores reales en lugar de los heredados del tenant
- El script desconecta automáticamente la sesión SPO al finalizar (`Disconnect-SPOService`)

---

## 📚 References / Referencias

- [Connect to SharePoint Online PowerShell](https://docs.microsoft.com/en-us/powershell/sharepoint/sharepoint-online/connect-sharepoint-online)
- [Get-SPOTenant](https://learn.microsoft.com/en-us/powershell/module/sharepoint-online/get-spotenant)
- [Get-SPOSite](https://learn.microsoft.com/en-us/powershell/module/sharepoint-online/get-sposite)

---

## 💬 Feedback

Any feedback is very valuable!  
¡Cualquier comentario es muy valioso!

**Author / Autor:** Gregorio Parra  
**Version / Versión:** 1.0  
**Date / Fecha:** March 2026
