# Exchange Scripts (Online y On-Prem)

Esta carpeta contiene scripts orientados a tareas de inventario, diagnostico y reporteo para entornos Exchange Online y Exchange On-Premises.

## Scripts incluidos

### 1) get-mailboxsizereport-365.ps1
- Objetivo:
  Genera un reporte detallado de buzones en Exchange Online.
- Que recopila:
  Alias, nombre, cuotas, estado de auditoria, estado de archivo (archive), datos de Litigation Hold, conteos de elementos, tamanos de buzones y borrados (incluyendo conversion a GB), y estadisticas de archive mailbox cuando aplica.
- Salida:
  Exporta un CSV con nombre tipo user-mailbox-size-365_yyyyMMdd-HHmms.csv en la ruta indicada al ejecutar el script.

### 2) get-mailboxsizereport-onpremise.ps1
- Objetivo:
  Genera un reporte de tamano y configuracion de buzones en Exchange On-Premises.
- Que recopila:
  Alias, SMTP primario, quotas de buzon y de base de datos, servidor, OU, base de datos, retencion, LegacyExchangeDN y estadisticas de uso (item count, total size, deleted size, logon/logoff, estado de limite).
- Salida:
  Exporta un CSV con nombre tipo user-mailbox-size_yyyyMMdd-HHmms.csv en la ruta indicada por el usuario.

### 3) review-autodiscover-scp.ps1
- Objetivo:
  Revisa en Active Directory los objetos SCP de Autodiscover de Exchange.
- Como lo hace:
  Busca objetos serviceConnectionPoint en el contenedor Configuration y filtra por keywords GUID de Autodiscover (SCP URL / SCP Pointer).
- Salida:
  Muestra en consola una tabla con Server, Site, SCPType, DateCreated, LastChanged, AutoDiscoverInternalURI y DN.

### 4) _script-coleccion_calendar.ps1
- Objetivo:
  Detecta calendarios compartidos/publicados en Exchange Online.
- Que recopila:
  Usuario, carpeta de calendario, estado PublishEnabled, nivel de detalle y URLs de publicacion (PublishedCalendarUrl y PublishedICalUrl).
- Salida:
  Exporta CSV en C:\temp\EXO-CalendarShares-yyy-mm-dd-hhmm.csv con los calendarios publicados encontrados.

## Nota

Algunos scripts fueron creados para escenarios especificos o versiones puntuales de Exchange. Antes de ejecutar en produccion, valida permisos, conectividad y cmdlets disponibles en tu entorno.
