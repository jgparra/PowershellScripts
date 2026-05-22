# Exchange Scripts (Online & On-Premises)

Scripts for inventory, diagnostics, and reporting in Exchange Online and Exchange On-Premises environments.

Scripts de inventario, diagnóstico y reportería para entornos Exchange Online y Exchange On-Premises.

---

## 📋 Available Scripts / Scripts Disponibles

### 1. get-mailboxsizereport-365.ps1

**English:**
- **Purpose:** Generates detailed mailbox report for Exchange Online
- **Collects:**
  - Alias, name, quotas, audit status
  - Archive status and Litigation Hold data
  - Item counts and mailbox sizes (converted to GB)
  - Archive mailbox statistics when applicable
  - Deleted item statistics
- **Output:** CSV file with naming pattern: `user-mailbox-size-365_yyyyMMdd-HHmms.csv`

**Español:**
- **Objetivo:** Genera reporte detallado de buzones en Exchange Online
- **Recopila:**
  - Alias, nombre, cuotas, estado de auditoría
  - Estado de archivo y datos de Litigation Hold
  - Conteo de elementos y tamaños de buzones (convertidos a GB)
  - Estadísticas de buzones de archivo cuando aplica
  - Estadísticas de elementos eliminados
- **Salida:** Archivo CSV con patrón: `user-mailbox-size-365_yyyyMMdd-HHmms.csv`

---

### 2. get-mailboxsizereport-onpremise.ps1

**English:**
- **Purpose:** Generates size and configuration report for Exchange On-Premises mailboxes
- **Collects:**
  - Alias, primary SMTP, mailbox and database quotas
  - Server, OU, database, retention settings
  - LegacyExchangeDN
  - Usage statistics: item count, total size, deleted size
  - Logon/logoff times, quota limit status
- **Output:** CSV file with naming pattern: `user-mailbox-size_yyyyMMdd-HHmms.csv`

**Español:**
- **Objetivo:** Genera reporte de tamaño y configuración de buzones Exchange On-Premises
- **Recopila:**
  - Alias, SMTP primario, cuotas de buzón y base de datos
  - Servidor, OU, base de datos, configuración de retención
  - LegacyExchangeDN
  - Estadísticas de uso: conteo de elementos, tamaño total, tamaño eliminado
  - Tiempos de inicio/cierre de sesión, estado de límite de cuota
- **Salida:** Archivo CSV con patrón: `user-mailbox-size_yyyyMMdd-HHmms.csv`

---

### 3. review-autodiscover-scp.ps1

**English:**
- **Purpose:** Reviews Exchange Autodiscover SCP objects in Active Directory
- **How it works:**
  - Searches for serviceConnectionPoint objects in Configuration container
  - Filters by Autodiscover GUID keywords (SCP URL / SCP Pointer)
- **Output:** Console table displaying:
  - Server, Site, SCPType
  - DateCreated, LastChanged
  - AutoDiscoverInternalURI, DN

**Español:**
- **Objetivo:** Revisa objetos SCP de Autodiscover de Exchange en Active Directory
- **Funcionamiento:**
  - Busca objetos serviceConnectionPoint en el contenedor Configuration
  - Filtra por palabras clave GUID de Autodiscover (SCP URL / SCP Pointer)
- **Salida:** Tabla en consola mostrando:
  - Server, Site, SCPType
  - DateCreated, LastChanged
  - AutoDiscoverInternalURI, DN

---

### 4. _script-coleccion_calendar.ps1

**English:**
- **Purpose:** Detects shared/published calendars in Exchange Online
- **Collects:**
  - User account, calendar folder
  - PublishEnabled status, detail level
  - Publishing URLs (PublishedCalendarUrl and PublishedICalUrl)
- **Output:** CSV file at `C:\temp\EXO-CalendarShares-yyy-mm-dd-hhmm.csv`

**Español:**
- **Objetivo:** Detecta calendarios compartidos/publicados en Exchange Online
- **Recopila:**
  - Cuenta de usuario, carpeta de calendario
  - Estado PublishEnabled, nivel de detalle
  - URLs de publicación (PublishedCalendarUrl y PublishedICalUrl)
- **Salida:** Archivo CSV en `C:\temp\EXO-CalendarShares-yyy-mm-dd-hhmm.csv`

---

### 5. exchschemaversion.ps1

**English:**
- **Purpose:** Detects Exchange Server version by querying Active Directory schema, organization, and domain version numbers
- **Coverage:** Supports all Exchange versions from 2000 to 2019/Subscription Edition (SE) including all service packs and cumulative updates

**Español:**
- **Objetivo:** Detecta la versión de Exchange Server consultando los números de versión de schema, organización y dominio en Active Directory
- **Cobertura:** Soporta todas las versiones de Exchange desde 2000 hasta 2019/Subscription Edition (SE) incluyendo todos los service packs y cumulative updates

---

## ⚠️ Important Notes / Notas Importantes

**English:**
- Some scripts were created for specific scenarios or particular Exchange versions
- Validate permissions, connectivity, and available cmdlets in your environment before running in production

**Español:**
- Algunos scripts fueron creados para escenarios específicos o versiones puntuales de Exchange
- Valida permisos, conectividad y cmdlets disponibles en tu entorno antes de ejecutar en producción

---

## 💬 Feedback

Any feedback is very valuable!  
¡Cualquier comentario es muy valioso!

**Gregorio Parra**
