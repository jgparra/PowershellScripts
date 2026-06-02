# PowershellScripts

Repository of useful PowerShell scripts related to Exchange, Active Directory, Microsoft 365, and system troubleshooting.

Repositorio de scripts útiles de PowerShell relacionados con Exchange, Active Directory, Microsoft 365 y resolución de problemas del sistema.

---

## 📁 Available Scripts / Scripts Disponibles

### 1. Exchange (On-Premises & Online)

**📂 Folder:** [Exchange_onprem_online](Exchange_onprem_online)  
**📄 Documentation:** [README.md](Exchange_onprem_online/README.md)

**English:**
- Scripts for managing Exchange Online and Exchange On-Premises environments
- Tools for mailbox management, migration, and troubleshooting

**Español:**
- Scripts para administrar entornos de Exchange Online y Exchange On-Premises
- Herramientas para gestión de buzones, migración y resolución de problemas

---

### 2. DMARC Reports Export (Outlook → CSV)

**📂 Folder:** [Dmarc_export](Dmarc_export)  
**📄 Documentation:** [README.md](Dmarc_export/README.md)

**English:**
- Reads unread emails from a selected Outlook folder
- Saves and extracts DMARC report attachments (XML/ZIP)
- Parses XML reports and generates two CSV files: `filereport.csv` and `rowreport.csv`
- Automated DMARC compliance reporting

**Español:**
- Lee correos no leídos desde una carpeta de Outlook seleccionada
- Guarda y extrae adjuntos de reportes DMARC (XML/ZIP)
- Procesa reportes XML y genera dos archivos CSV: `filereport.csv` y `rowreport.csv`
- Reportería automatizada de cumplimiento DMARC

---

### 3. SharePoint Online Reports

**📂 Folder:** [Reportes_spo](Reportes_spo)  
**📄 Documentation:** [README.md](Reportes_spo/README.md)

**English:**
- Comprehensive SharePoint Online tenant and site inventory reports
- Connects to SharePoint Online admin center
- Collects tenant global configuration (sharing, storage, security)
- Processes all sites and exports detailed inventory
- Generates timestamped TXT (tenant config) and CSV (sites inventory) reports

**Español:**
- Reportes comprehensivos de tenant y sitios de SharePoint Online
- Conecta al centro de administración de SharePoint Online
- Recopila configuración global del tenant (compartir, almacenamiento, seguridad)
- Procesa todos los sitios y exporta inventario detallado
- Genera reportes TXT (config del tenant) y CSV (inventario de sitios) con timestamp

---

### 4. Windows AD / Entra ID Scripts

**📂 Folder:** [Windows_AD_EntraID](Windows_AD_EntraID)  
**📄 Documentation:** [README.md](Windows_AD_EntraID/README.md)

**English:**
- Scripts for on-premises Active Directory, Microsoft Entra ID (Azure AD), and Azure AD Connect
- **TLS Security Audit:** Analyzes TLS protocol versions and cipher suite configurations
- **Azure AD Connect DB Operations:** Interactive diagnostic tool for ADSync database troubleshooting
- Identifies security vulnerabilities (weak protocols, insecure cipher suites)
- Database space management and fragmentation analysis
- Provides structured audit results with severity levels and recommendations
- No external modules required

**Español:**
- Scripts para Active Directory on-premises, Microsoft Entra ID (Azure AD) y Azure AD Connect
- **Auditoría de Seguridad TLS:** Analiza versiones de protocolo TLS y configuraciones de cipher suites
- **Operaciones de BD Azure AD Connect:** Herramienta de diagnóstico interactiva para resolución de problemas de base de datos ADSync
- Identifica vulnerabilidades de seguridad (protocolos débiles, cipher suites inseguros)
- Gestión de espacio de base de datos y análisis de fragmentación
- Proporciona resultados de auditoría estructurados con niveles de severidad y recomendaciones
- No requiere módulos externos

---

### 5. Network Scripts

**📂 Folder:** [Network](Network)  
**📄 Documentation:** [README.md](Network/README.md)

**English:**
- Network diagnostics scripts for Microsoft 365 connectivity and proxy source attribution
- **TLS Connectivity & Certificate Inspection:** validates port 443 reachability, extracts certificate details, and detects TLS interception
- **Proxy Attribution Tool:** identifies proxy configuration origin using WinINet, WinHTTP, GPO, WPAD, and PAC validation
- Supports direct and explicit-proxy scenarios
- Generates actionable console output and timestamped artifacts (TXT/JSON)

**Español:**
- Scripts de diagnóstico de red para conectividad de Microsoft 365 y atribución de origen de proxy
- **Inspección TLS y Certificados:** valida alcance por puerto 443, extrae detalles de certificado y detecta interceptación TLS
- **Herramienta de Atribución de Proxy:** identifica origen de configuración de proxy usando WinINet, WinHTTP, GPO, WPAD y validación PAC
- Soporta escenarios de conexión directa y proxy explícito
- Genera salida accionable en consola y artefactos con timestamp (TXT/JSON)

---

## 💬 Feedback

Any feedback is very valuable!  
¡Cualquier comentario es muy valioso!

**Best Regards / Saludos,**  
Gregorio Parra
