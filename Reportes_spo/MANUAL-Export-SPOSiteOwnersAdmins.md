# Manual de uso — Export-SPOSiteOwnersAdmins.ps1

> Reporte de **Owners** (propietarios) y **Admins** (administradores) de los sitios de
> SharePoint Online del tenant. Este manual está pensado para alguien que **está
> empezando** en la administración de SharePoint Online. Sigue los pasos en orden.

---

## 1. ¿Qué hace el script?

Recorre todas (o algunas) de las *site collections* de tu organización y, por cada
sitio, genera un reporte CSV con:

- **Site Collection Administrators** (los administradores del sitio).
- **Miembros del grupo Owners** (los propietarios del sitio).

Si un administrador u owner es un **grupo** (de seguridad de Entra ID o de Microsoft
365), el script lo **expande** para mostrar los usuarios individuales que hay dentro.

Es un script de **solo lectura**: NO cambia permisos ni borra nada. Solo lee y exporta.

En el CSV, cuando una celda tiene varios usuarios, estos se separan con el carácter
**pipe `|`** para no confundirse con la coma o el punto y coma que separan las columnas.

---

## 2. Requerimientos

### 2.1 Software

| Requisito | Detalle |
|-----------|---------|
| **PowerShell 7.2 o superior** | El módulo PnP.PowerShell 2.x NO funciona en Windows PowerShell 5.1. Debes usar `pwsh`. |
| **Módulo PnP.PowerShell** | Conexión a SharePoint Online. |
| **Módulo Microsoft.Graph.Groups** | Para expandir los grupos a sus miembros. |
| **Módulo Microsoft.Graph.Users** | Para validar/leer usuarios. |

### 2.2 Permisos y roles de tu cuenta

| Necesitas | Para qué |
|-----------|----------|
| Rol **Administrador de SharePoint** (o Global Admin) | Enumerar todos los sitios del tenant. |
| Consentir permisos de **Microsoft Graph**: `User.Read.All` y `GroupMember.Read.All` | Leer usuarios y expandir grupos. |
| Poder **registrar aplicaciones** en Entra ID (Global Admin o Application Administrator) | Crear la App Registration que exige PnP (solo la primera vez). |

> **Nota:** los permisos de Graph solo necesitan que **tu cuenta de administrador**
> los consienta (para ti). NO es necesario habilitarlos "para toda la organización".

---

## 3. Preparación (se hace una sola vez)

### Paso 3.1 — Abrir PowerShell 7

Abre una terminal **pwsh** (no la de Windows PowerShell azul). Verifica la versión:

```powershell
$PSVersionTable.PSVersion    # Debe ser 7.2 o superior
```

### Paso 3.2 — Instalar los módulos

```powershell
Install-Module PnP.PowerShell        -Scope CurrentUser
Install-Module Microsoft.Graph.Groups -Scope CurrentUser
Install-Module Microsoft.Graph.Users  -Scope CurrentUser
```

Confirmar que quedaron instalados:

```powershell
Get-Module -ListAvailable PnP.PowerShell, Microsoft.Graph.Groups, Microsoft.Graph.Users |
    Select-Object Name, Version
```

### Paso 3.3 — Registrar la App de PnP (obligatorio en PnP 2.x)

Las versiones nuevas de PnP.PowerShell **ya no traen una app por defecto**, así que
debes crear una App Registration en Entra ID. PnP lo hace por ti con este comando
(reemplaza el `-Tenant` por tu dominio `.onmicrosoft.com`):

```powershell
Register-PnPEntraIDAppForInteractiveLogin `
    -ApplicationName "PnP-Rocks" `
    -Tenant "TUDOMINIO.onmicrosoft.com"
```

> ⚠️ **Cuidado:** este comando **NO** lleva el parámetro `-Interactive` (da error).
> Ya es interactivo por su nombre.

Al terminar:
1. Se abre el navegador para iniciar sesión y **consentir** los permisos de la app.
2. El comando imprime un **AppId (ClientId)** con formato GUID
   (ej.: `a38a0aab-2f62-4686-9604-21cb017d33d4`). **Cópialo y guárdalo**, lo usarás
   en cada ejecución.

Si perdiste el ClientId, búscalo en el portal:
**Entra ID → App registrations → "PnP-Rocks" → Application (client) ID**.

> ⏱️ Una app recién creada puede tardar **2 a 5 minutos** en activarse. Si la primera
> ejecución falla diciendo que la app no existe, espera unos minutos y reintenta.

---

## 4. Ejecución

### Paso 4.1 — Ir a la carpeta del script

```powershell
cd "C:\ODB-MSN\OneDrive\scripts\en-github\PowerShellScripts\PowershellScripts\SharePoint_online"
```

### Paso 4.2 — Probar primero con UN solo sitio (recomendado)

Antes de correrlo en todo el tenant, valida la salida con un sitio usando
`-SiteFilter`. Reemplaza el ClientId y el nombre del sitio:

```powershell
.\Export-SPOSiteOwnersAdmins.ps1 `
    -TenantAdminUrl "https://TUTENANT-admin.sharepoint.com" `
    -ClientId "TU-CLIENT-ID" `
    -SiteFilter "*/sites/NOMBRE-DE-UN-SITIO*"
```

**Qué va a pasar:**
- Se abrirán ventanas de inicio de sesión: una para **Microsoft Graph** (consiente
  `User.Read.All` y `GroupMember.Read.All`) y otra para **SharePoint/PnP**.
- Inicia sesión con tu cuenta de **administrador**.
- Verás una barra de progreso y, al final, la ruta del CSV generado.

### Paso 4.3 — Ejecutar en todo el tenant

Cuando la prueba se vea bien, quita el filtro:

```powershell
.\Export-SPOSiteOwnersAdmins.ps1 `
    -TenantAdminUrl "https://TUTENANT-admin.sharepoint.com" `
    -ClientId "TU-CLIENT-ID"
```

### Paso 4.4 — Parámetros disponibles

| Parámetro | Obligatorio | Descripción |
|-----------|:-----------:|-------------|
| `-TenantAdminUrl` | Sí | URL del centro de administración: `https://TUTENANT-admin.sharepoint.com`. |
| `-ClientId` | Sí (PnP 2.x) | AppId de la App Registration creada en el paso 3.3. |
| `-SiteFilter` | No | Comodín para limitar sitios por URL. Ej.: `"*/sites/finanzas*"`. Ideal para probar. |
| `-IncludeOneDrive` | No | Incluye los OneDrive personales (por defecto se excluyen). |
| `-ExpandAdminGroups` | No | Expande a miembros los admins que sean grupos. Activo por defecto. Usa `-ExpandAdminGroups:$false` para dejarlos como grupo. |
| `-OutputPath` | No | Carpeta donde guardar el CSV (por defecto, la carpeta del script). |

> 💡 Puedes cancelar con **Ctrl+C**: el script conserva y exporta lo recolectado
> hasta ese momento.

---

## 5. Formato del reporte (CSV)

El archivo se llama `SPO_SiteOwnersAdmins_AAAA-MM-DD--HH-mm.csv` y se guarda en la
carpeta del script (o en `-OutputPath`). Tiene **una fila por sitio**.

| Columna | Contenido |
|---------|-----------|
| `SiteUrl` | URL de la site collection. |
| `Title` | Nombre del sitio. |
| `Template` | Plantilla del sitio (ej.: `GROUP#0`, `SITEPAGEPUBLISHING#0`). |
| `Admins` | UPN de los administradores, separados por `|`. |
| `AdminsCount` | Cuántos administradores tiene el sitio. |
| `Owners` | UPN de los propietarios (grupo Owners), separados por `|`. |
| `OwnersCount` | Cuántos propietarios tiene el sitio. |

**Ejemplo de una celda con varios usuarios:**

```
ana@contoso.com|luis@contoso.com|maria@contoso.com
```

### Revisar el CSV rápidamente en consola

```powershell
Get-ChildItem .\SPO_SiteOwnersAdmins_*.csv | Sort-Object LastWriteTime |
    Select-Object -Last 1 | Import-Csv |
    Format-Table SiteUrl, AdminsCount, Admins, OwnersCount, Owners -AutoSize
```

> Si abres el CSV en Excel y ves los usuarios "pegados", recuerda que el separador
> **dentro** de una celda es `|`; el separador de **columnas** es la coma. Excel los
> distingue correctamente al importar.

---

## 6. Solución de problemas (según el error)

### ❌ "Please specify a valid client id for an Entra ID App Registration"
### ❌ "Connect-PnPOnline: Specified method is not supported"
**Causa:** PnP 2.x necesita un ClientId y no se lo pasaste.
**Solución:** registra la app (paso 3.3) y ejecuta el script con `-ClientId "TU-CLIENT-ID"`.

---

### ❌ "You are not signed in. Please use Connect-PnPOnline to connect"
**Causa:** la conexión a SharePoint falló antes (normalmente por el ClientId faltante),
así que `Get-PnPTenantSite` no tenía sesión.
**Solución:** corrige el ClientId. Cuando la conexión funcione, este error desaparece.

---

### ❌ "A parameter cannot be found that matches parameter name 'Interactive'"
**Causa:** le pasaste `-Interactive` a `Register-PnPEntraIDAppForInteractiveLogin`.
**Solución:** quita ese parámetro. El cmdlet ya es interactivo por sí mismo.

---

### ⚠️ WARNING: "No se pudo expandir el grupo '...' en Graph"
### El CSV muestra `OwnersCount` o `AdminsCount` en 0 donde esperabas usuarios
**Causa:** no consentiste los permisos de Graph (`User.Read.All`,
`GroupMember.Read.All`), así que la expansión de grupos falló silenciosamente.
**Solución:** vuelve a consentir y re-ejecuta:

```powershell
Disconnect-MgGraph
Connect-MgGraph -Scopes "User.Read.All","GroupMember.Read.All" -Prompt Consent
(Get-MgContext).Scopes    # Confirma que ambos scopes aparecen
```

Luego corre el script de nuevo. **El consentimiento se puede dar en cualquier
ejecución posterior**, no se pierde la oportunidad.

---

### ⚠️ WARNING: "Sign in by Web Account Manager (WAM)... la ventana puede estar oculta"
**Causa:** en terminales integradas (como la de VS Code) la ventana de login puede
quedar detrás de otras ventanas.
**Solución:** busca la ventana de inicio de sesión en la barra de tareas / con Alt+Tab.
Si es muy molesto, ejecuta el script desde una terminal `pwsh` independiente.

---

### ❌ La app recién registrada "no existe" en la primera ejecución
**Causa:** la App Registration tarda unos minutos en propagarse en Entra ID.
**Solución:** espera **2 a 5 minutos** y vuelve a ejecutar.

---

### ❌ "No se encontró el módulo PnP.PowerShell / Microsoft Graph"
**Causa:** falta instalar un módulo.
**Solución:** repite el paso 3.2 (`Install-Module ...`).

---

### ❌ El script no corre / errores raros de tipos
**Causa:** lo estás ejecutando en **Windows PowerShell 5.1** en vez de PowerShell 7.
**Solución:** ábrelo en `pwsh` y verifica con `$PSVersionTable.PSVersion` (≥ 7.2).

---

## 7. Checklist rápido

- [ ] Estoy en **PowerShell 7** (`pwsh`).
- [ ] Instalé **PnP.PowerShell**, **Microsoft.Graph.Groups** y **Microsoft.Graph.Users**.
- [ ] Registré la app con `Register-PnPEntraIDAppForInteractiveLogin` y tengo el **ClientId**.
- [ ] Consentí los scopes de Graph (`User.Read.All`, `GroupMember.Read.All`).
- [ ] Tengo rol de **Administrador de SharePoint**.
- [ ] Probé primero con `-SiteFilter` en un solo sitio.
- [ ] Revisé el CSV y las celdas múltiples usan `|`.
- [ ] Ejecuté en todo el tenant sin `-SiteFilter`.

---

*Documento de apoyo para `Export-SPOSiteOwnersAdmins.ps1`.*
