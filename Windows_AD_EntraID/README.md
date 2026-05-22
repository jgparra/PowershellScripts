# Windows AD / Entra ID Scripts

Scripts for on-premises Active Directory, Microsoft Entra ID (Azure AD), and Azure AD Connect administration and auditing.

Scripts para administración y auditoría de Active Directory on-premises, Microsoft Entra ID (Azure AD) y Azure AD Connect.

---

## 📋 Available Scripts / Scripts Disponibles

### 1. TLS Security Audit

**📄 Script:** `TLS_Security_Audit.ps1`

#### Description / Descripción

**English:**

Security audit script that analyzes TLS and cipher suite configurations on Windows systems to identify potential security vulnerabilities.

**Verifies:**
- Enabled/disabled TLS protocol versions
- Cipher suite configurations
- Windows registry settings (SCHANNEL)
- Identification of insecure protocols and cipher suites

**Español:**

Script de auditoría de seguridad que analiza la configuración de TLS y cipher suites en sistemas Windows para identificar vulnerabilidades de seguridad potenciales.

**Verifica:**
- Versiones de protocolo TLS habilitadas/deshabilitadas
- Configuración de cipher suites (conjuntos de cifrado)
- Configuración del registro de Windows (SCHANNEL)
- Identificación de protocolos y cipher suites inseguros

#### Prerequisites / Requisitos

**English:**
- **Windows PowerShell** 5.1 or higher
- **Permissions:** Run as Administrator (recommended for full registry access)
- **OS:** Windows Server 2012+ / Windows 8+

**Español:**
- **Windows PowerShell** 5.1 o superior
- **Permisos:** Ejecutar como Administrador (recomendado para acceso completo al registro)
- **SO:** Windows Server 2012+ / Windows 8+

#### Usage / Uso

```powershell
# Run as Administrator / Ejecutar como Administrador
.\TLS_Security_Audit.ps1
```

---

### 2. Azure AD Connect Database Operations

**📄 Script:** `azureADconnectDBOps.ps1`

#### Description / Descripción

**English:**

Interactive diagnostic tool for Azure AD Connect (formerly DirSync) SQL database. Provides menu-driven interface to query and analyze the ADSync database for troubleshooting and space management.

**Features:**
- Database space usage analysis
- Object count reporting (metaverse and connector space)
- SQL error log review
- Database fragmentation analysis
- Table space usage breakdown

**Note:** This tool is **not officially supported by Microsoft**. Created for troubleshooting scenarios.

**Español:**

Herramienta de diagnóstico interactiva para la base de datos SQL de Azure AD Connect (anteriormente DirSync). Proporciona interfaz de menú para consultar y analizar la base de datos ADSync para resolución de problemas y gestión de espacio.

**Funcionalidades:**
- Análisis de uso de espacio de base de datos
- Reporte de conteo de objetos (metaverse y connector space)
- Revisión de log de errores SQL
- Análisis de fragmentación de base de datos
- Desglose de uso de espacio por tabla

**Nota:** Esta herramienta **no tiene soporte oficial de Microsoft**. Creada para escenarios de resolución de problemas.

#### Menu Options / Opciones del Menú

| Option | English | Español |
|---|---|---|
| **1** | Get space used | Obtener espacio usado |
| **2** | Get object count | Obtener conteo de objetos |
| **3** | Get error log | Obtener log de errores |
| **4** | Get DB fragmentation | Obtener fragmentación de DB |
| **5** | Get table space used | Obtener espacio usado por tabla |
| **98** | Restart the server | Reiniciar el servidor |
| **99** | Exit | Salir |

#### Prerequisites / Requisitos

**English:**
- Must be run on the **Azure AD Connect server**
- **Windows PowerShell** 5.1 or higher
- **Permissions:** Administrator access to query LocalDB instance
- SQL LocalDB instance must be running

**Español:**
- Debe ejecutarse en el **servidor de Azure AD Connect**
- **Windows PowerShell** 5.1 o superior
- **Permisos:** Acceso de Administrador para consultar instancia LocalDB
- La instancia SQL LocalDB debe estar ejecutándose

#### Usage / Uso

```powershell
.\azureADconnectDBOps.ps1
```

**English:** The script presents an interactive menu. Select options 1-5 for diagnostics, 99 to exit.

**Español:** El script presenta un menú interactivo. Seleccione opciones 1-5 para diagnósticos, 99 para salir.

#### Output Information / Información de Salida

**English:**

Each menu option queries the ADSync SQL database and returns:

1. **Space Used:** Database version, total size, used/unused space
2. **Object Count:** Count of objects by type in metaverse and connector space
3. **Error Log:** Recent SQL Server error log entries
4. **Fragmentation:** Index fragmentation percentage for tables
5. **Table Space:** Detailed breakdown of space usage per table (KB)

**Español:**

Cada opción del menú consulta la base de datos SQL ADSync y devuelve:

1. **Espacio Usado:** Versión de base de datos, tamaño total, espacio usado/sin usar
2. **Conteo de Objetos:** Conteo de objetos por tipo en metaverse y connector space
3. **Log de Errores:** Entradas recientes del log de errores de SQL Server
4. **Fragmentación:** Porcentaje de fragmentación de índices por tabla
5. **Espacio de Tabla:** Desglose detallado de uso de espacio por tabla (KB)

---

## ✨ Features / Características

### TLS Security Audit

**English:**
- ✅ Verifies TLS 1.0, 1.1, 1.2, and 1.3 configuration
- ✅ Identifies potentially insecure cipher suites
- ✅ Generates structured report with severity levels
- ✅ Provides security recommendations
- ✅ No external modules required

**Español:**
- ✅ Verifica configuración de TLS 1.0, 1.1, 1.2 y 1.3
- ✅ Identifica cipher suites potencialmente inseguros
- ✅ Genera reporte estructurado con niveles de severidad
- ✅ Proporciona recomendaciones de seguridad
- ✅ No requiere módulos externos

### Azure AD Connect DB Operations

**English:**
- ✅ Interactive menu-driven interface
- ✅ Real-time SQL database diagnostics
- ✅ Space usage and fragmentation analysis
- ✅ Object counting for capacity planning
- ✅ Error log review capabilities
- ✅ No external SQL tools required (uses sqlcmd)

**Español:**
- ✅ Interfaz interactiva basada en menú
- ✅ Diagnósticos de base de datos SQL en tiempo real
- ✅ Análisis de uso de espacio y fragmentación
- ✅ Conteo de objetos para planificación de capacidad
- ✅ Capacidades de revisión de logs de error
- ✅ No requiere herramientas SQL externas (usa sqlcmd)

---

## 📋 Prerequisites / Requisitos

**English:**
- **Windows PowerShell** 5.1 or higher
- **Permissions:** Run as Administrator (recommended for full registry access)
- **OS:** Windows Server 2012+ / Windows 8+

**Español:**
- **Windows PowerShell** 5.1 o superior
- **Permisos:** Ejecutar como Administrador (recomendado para acceso completo al registro)
- **SO:** Windows Server 2012+ / Windows 8+

---

## 🚀 Usage / Uso

```powershell
# Run as Administrator / Ejecutar como Administrador
.\TLS_Security_Audit.ps1
```

---

## 📤 Output / Salida

**English:**

The script generates a PSCustomObject with the following fields:

**Español:**

El script genera un objeto PSCustomObject con los siguientes campos:

## 📤 Output / Salida

**English:**

The script generates a PSCustomObject with the following fields:

**Español:**

El script genera un objeto PSCustomObject con los siguientes campos:

| Field / Campo | Description / Descripción |
|---|---|
| **Protocol/Component** | Protocol or component name analyzed / Nombre del protocolo o componente analizado |
| **Status** | State: `Healthy`, `Warning`, `Critical`, `Missing` / Estado: `Healthy`, `Warning`, `Critical`, `Missing` |
| **Severity** | Severity level of the finding / Nivel de severidad del hallazgo |
| **Message** | Detailed result message / Mensaje detallado del resultado |
| **Recommendations** | Applicable security recommendations / Recomendaciones de seguridad aplicables |

---

## 🎯 Status Levels / Niveles de Estado

| Status / Estado | Icon | Meaning / Significado |
|---|---|---|
| **Healthy** | ✅ | Secure configuration / Configuración segura |
| **Warning** | ⚠️ | Requires attention, possible improvement / Requiere atención, posible mejora |
| **Critical** | ❌ | Security vulnerability identified / Vulnerabilidad de seguridad identificada |
| **Missing** | ℹ️ | Configuration not found / Configuración no encontrada |

---

## 📝 Example Output / Ejemplo de Salida

```powershell
Protocol/Component : TLS 1.0
Status             : Critical
Severity           : High
Message            : TLS 1.0 está habilitado - protocolo inseguro
Recommendations    : Deshabilitar TLS 1.0 en producción

Protocol/Component : TLS 1.2
Status             : Healthy
Severity           : Info
Message            : TLS 1.2 correctamente configurado
Recommendations    : Mantener habilitado
```

---

## 🔒 Security Recommendations / Recomendaciones de Seguridad

**English:**

Based on current security best practices:

**1. Disable obsolete protocols:**
   - TLS 1.0 (RFC 8996 - Deprecated)
   - TLS 1.1 (RFC 8996 - Deprecated)
   - SSL 2.0, SSL 3.0

**2. Enable modern protocols:**
   - TLS 1.2 (minimum recommended)
   - TLS 1.3 (latest and most secure)

**3. Cipher Suites:**
   - Disable weak cipher suites (DES, RC4, MD5)
   - Prioritize cipher suites with Forward Secrecy (ECDHE, DHE)
   - Use AES-GCM ciphers when possible

**Español:**

Basado en las mejores prácticas de seguridad actuales:

**1. Deshabilitar protocolos obsoletos:**
   - TLS 1.0 (RFC 8996 - Deprecated)
   - TLS 1.1 (RFC 8996 - Deprecated)
   - SSL 2.0, SSL 3.0

**2. Habilitar protocolos modernos:**
   - TLS 1.2 (mínimo recomendado)
   - TLS 1.3 (más reciente y seguro)

**3. Cipher Suites:**
   - Deshabilitar cipher suites con cifrados débiles (DES, RC4, MD5)
   - Priorizar cipher suites con Forward Secrecy (ECDHE, DHE)
   - Usar cifrados AES-GCM cuando sea posible

---

## 📚 References / Referencias

- [Microsoft TLS Best Practices](https://learn.microsoft.com/en-us/windows-server/security/tls/tls-registry-settings)
- [NIST TLS Guidelines](https://csrc.nist.gov/publications/detail/sp/800-52/rev-2/final)
- [RFC 8996 - Deprecating TLS 1.0 and 1.1](https://www.rfc-editor.org/rfc/rfc8996)

---

## 💬 Feedback

Any feedback is very valuable!  
¡Cualquier comentario es muy valioso!

**Gregorio Parra**
