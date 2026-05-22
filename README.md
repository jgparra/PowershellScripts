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
- Scripts for on-premises Active Directory and Microsoft Entra ID (Azure AD)
- **TLS Security Audit:** Analyzes TLS protocol versions and cipher suite configurations
- Identifies security vulnerabilities (weak protocols, insecure cipher suites)
- Provides structured audit results with severity levels and recommendations
- No external modules required

**Español:**
- Scripts para Active Directory on-premises y Microsoft Entra ID (Azure AD)
- **Auditoría de Seguridad TLS:** Analiza versiones de protocolo TLS y configuraciones de cipher suites
- Identifica vulnerabilidades de seguridad (protocolos débiles, cipher suites inseguros)
- Proporciona resultados de auditoría estructurados con niveles de severidad y recomendaciones
- No requiere módulos externos

---

## 💬 Feedback

Any feedback is very valuable!  
¡Cualquier comentario es muy valioso!

**Best Regards / Saludos,**  
Gregorio Parra
