# PowershellScripts

Useful PowerShell scripts for Exchange, Microsoft 365, Active Directory, Entra ID, SharePoint, and network troubleshooting.  
Scripts utiles de PowerShell para Exchange, Microsoft 365, Active Directory, Entra ID, SharePoint y diagnostico de red.

---

## Available Scripts / Scripts Disponibles

### 1. Exchange (On-Premises & Online)
Folder: [Exchange_onprem_online](Exchange_onprem_online) | Docs: [Exchange_onprem_online/README.md](Exchange_onprem_online/README.md)
- EN: Mailbox management, migration, and troubleshooting for Exchange Online/On-Prem.
- ES: Gestion de buzones, migracion y troubleshooting para Exchange Online/On-Prem.

### 2. DMARC Reports Export (Outlook to CSV)
Folder: [Dmarc_export](Dmarc_export) | Docs: [Dmarc_export/README.md](Dmarc_export/README.md)
- EN: Reads unread Outlook emails, extracts XML/ZIP DMARC attachments, and generates filereport.csv and rowreport.csv.
- ES: Lee correos no leidos en Outlook, extrae adjuntos XML/ZIP DMARC y genera filereport.csv y rowreport.csv.

### 3. SharePoint Online Reports
Folder: [Reportes_spo](Reportes_spo) | Docs: [Reportes_spo/README.md](Reportes_spo/README.md)
- EN: Tenant and site inventory reporting with TXT (tenant config) and CSV (site inventory) outputs.
- ES: Reporteria de tenant y sitios con salida TXT (configuracion tenant) y CSV (inventario de sitios).

### 4. Windows AD / Entra ID
Folder: [Windows_AD_EntraID](Windows_AD_EntraID) | Docs: [Windows_AD_EntraID/README.md](Windows_AD_EntraID/README.md)
- EN: Includes TLS security audit and Azure AD Connect ADSync DB diagnostics; no external modules required.
- ES: Incluye auditoria TLS y diagnostico de BD ADSync de Azure AD Connect; sin modulos externos.

### 5. Network
Folder: [Network](Network) | Docs: [Network/README.md](Network/README.md)
- EN: Microsoft 365 connectivity checks, proxy attribution (WinINet/WinHTTP/GPO/WPAD/PAC), and offline TLS PCAP/PCAPNG analysis with TXT/JSON/CSV exports.
- ES: Validacion de conectividad Microsoft 365, atribucion de proxy (WinINet/WinHTTP/GPO/WPAD/PAC) y analisis offline TLS de PCAP/PCAPNG con exportes TXT/JSON/CSV.

### 6. Entra Group Members and Managers Export
Script: [Get-GroupsMembersManagers.ps1](Get-GroupsMembersManagers.ps1)
- EN: Exports Entra users with group context and manager mapping to EntraGroupMembersExport.csv (all groups, single group, or CSV input with UPN).
- ES: Exporta usuarios Entra con contexto de grupo y manager a EntraGroupMembersExport.csv (todos los grupos, grupo unico o CSV con UPN).
- Requirements / Requisitos: Microsoft Graph PowerShell, scopes Group.Read.All and User.Read.All.

---

## Feedback

Any feedback is welcome. / Cualquier comentario es bienvenido.

Gregorio Parra
