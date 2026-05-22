# Windows AD / Entra ID Scripts

Scripts for on-premises Active Directory and Microsoft Entra ID (Azure AD) administration and auditing.

Scripts para administración y auditoría de Active Directory on-premises y Microsoft Entra ID (Azure AD).

---

## 📋 Available Scripts / Scripts Disponibles

### TLS Security Audit

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

---

## ✨ Features / Características

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
