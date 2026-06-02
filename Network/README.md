# Network Scripts / Scripts de Red

Scripts for network connectivity validation, TLS certificate inspection, and proxy source attribution.

Scripts para validación de conectividad de red, inspección de certificados TLS y atribución de origen de proxy.

---

## 📋 Available Scripts / Scripts Disponibles

### 1. test-certificateinconnection.ps1

**English:**
- **Purpose:** Validates Microsoft 365 endpoint connectivity and inspects TLS certificates to detect SSL/TLS interception.
- **Checks:**
  - TCP connectivity on port 443 for key Microsoft 365 endpoints
  - Certificate details (Subject, Issuer, validity period, thumbprint)
  - Proxy/interception indicators in HTTP response headers
- **Modes:**
  - `NORMAL`: alerts only when issuer matches known proxy/security appliances
  - `STRICT`: alerts when issuer is not DigiCert/Microsoft
- **Proxy Support:**
  - Direct connection or explicit proxy via `-ProxyServer` and `-ProxyPort`

**Español:**
- **Objetivo:** Valida conectividad a endpoints de Microsoft 365 e inspecciona certificados TLS para detectar interceptación SSL/TLS.
- **Verifica:**
  - Conectividad TCP por puerto 443 hacia endpoints clave de Microsoft 365
  - Detalles de certificado (Subject, Issuer, vigencia, thumbprint)
  - Indicadores de proxy/intercepción en headers HTTP
- **Modos:**
  - `NORMAL`: alerta solo cuando el emisor coincide con proxies/appliances conocidos
  - `STRICT`: alerta cuando el emisor no es DigiCert/Microsoft
- **Soporte de Proxy:**
  - Conexión directa o proxy explícito con `-ProxyServer` y `-ProxyPort`

#### Parameters / Parámetros

| Parameter | Type | Default | Description |
|---|---|---|---|
| `StrictMode` | `bool` | `false` | Enables strict issuer validation / Activa validación estricta de emisor |
| `ProxyServer` | `string` | `""` | Proxy host or IP / Host o IP del proxy |
| `ProxyPort` | `int` | `8080` | Proxy port / Puerto del proxy |

#### Usage / Uso

```powershell
# Direct mode
.\test-certificateinconnection.ps1

# Strict mode
.\test-certificateinconnection.ps1 -StrictMode $true

# Via proxy
.\test-certificateinconnection.ps1 -ProxyServer "172.202.41.246" -ProxyPort 3128
```

---

### 2. Test-ProxyDetectionTool.ps1

**English:**
- **Purpose:** Detects proxy configuration sources and attributes where proxy behavior is coming from.
- **Collects evidence from:**
  - WinINet (user and machine-default)
  - WinHTTP
  - Policy-based registry settings (GPO)
  - WPAD DNS
  - PAC URL reachability and PAC logic validation (`FindProxyForURL`)
- **Output:**
  - Console summary with confidence and primary detection path
  - Timestamped TXT log
  - Timestamped JSON report

**Español:**
- **Objetivo:** Detecta fuentes de configuración de proxy y atribuye el origen del comportamiento de proxy.
- **Recopila evidencia de:**
  - WinINet (usuario y máquina por defecto)
  - WinHTTP
  - Configuración por políticas (GPO)
  - DNS de WPAD
  - Alcance de URL PAC y validación de lógica PAC (`FindProxyForURL`)
- **Salida:**
  - Resumen en consola con nivel de confianza y ruta primaria de detección
  - Log TXT con timestamp
  - Reporte JSON con timestamp

#### Parameters / Parámetros

| Parameter | Type | Default | Description |
|---|---|---|---|
| `LogPath` | `string` | `.\Test-ProxyDetectionTool.log` | Base log file path (timestamp is appended) / Ruta base del log (se agrega timestamp) |
| `JsonPath` | `string` | `.\Test-ProxyDetectionTool.json` | Base JSON report path (timestamp is appended) / Ruta base del JSON (se agrega timestamp) |
| `TimeoutSec` | `int` | `6` | Timeout for PAC/web requests / Timeout para solicitudes PAC/web |
| `RawEvidence` | `switch` | `false` | Adds raw WinHTTP text output / Agrega salida cruda de WinHTTP |

#### Usage / Uso

```powershell
# Default run
.\Test-ProxyDetectionTool.ps1

# Custom timeout and raw evidence
.\Test-ProxyDetectionTool.ps1 -TimeoutSec 10 -RawEvidence
```

---

## ⚠️ Notes / Notas

**English:**
- Scripts are designed for diagnostics and troubleshooting scenarios.
- Validate permissions and environment requirements before production use.
- Results can vary depending on execution context (user/admin/SYSTEM) and network controls.

**Español:**
- Los scripts están orientados a diagnóstico y resolución de problemas.
- Validar permisos y requisitos del entorno antes de uso en producción.
- Los resultados pueden variar según el contexto de ejecución (usuario/admin/SYSTEM) y controles de red.

---

## 💬 Feedback

Any feedback is very valuable!  
¡Cualquier comentario es muy valioso!

**Gregorio Parra**
