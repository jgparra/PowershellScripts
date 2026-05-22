# DMARC Reports Export (Outlook → CSV)

Automates extraction and transformation of DMARC reports received via email in Outlook.

Automatiza la extracción y transformación de reportes DMARC recibidos por correo en Outlook.

---

## 📄 Script: _script_dmarc_outlook_to_csv.ps1

### Description / Descripción

**English:**

This script automates the extraction and transformation of DMARC reports received via email in Outlook into structured CSV files.

**Workflow:**
1. Verifies and imports the `7Zip4Powershell` module
2. Creates a temporary working directory in `%TEMP%` with `output` and `xml` subfolders
3. Prompts to select an Outlook folder (PickFolder)
4. Processes **unread emails** from that folder and saves attachments to temp directory
5. Decompresses compressed attachments (`*.*z*`) to extract DMARC XML files
6. Converts XML information into two datasets:
   - File/report metadata
   - Row-level records from DMARC reports
7. Exports two CSV files and opens output folder in Explorer

**Español:**

Este script automatiza la extracción y transformación de reportes DMARC recibidos por correo en Outlook en archivos CSV estructurados.

**Flujo de trabajo:**
1. Verifica e importa el módulo `7Zip4Powershell`
2. Crea un directorio de trabajo temporal en `%TEMP%` con subcarpetas `output` y `xml`
3. Solicita seleccionar una carpeta de Outlook (PickFolder)
4. Procesa correos **no leídos** de esa carpeta y guarda adjuntos en directorio temporal
5. Descomprime archivos adjuntos comprimidos (`*.*z*`) para extraer archivos XML DMARC
6. Convierte la información XML en dos conjuntos de datos:
   - Metadatos por archivo/reporte
   - Registros por fila del reporte DMARC
7. Exporta dos archivos CSV y abre la carpeta de salida en el Explorador

---

## 📋 Prerequisites / Prerrequisitos

**English:**
- Windows with Outlook installed and MAPI profile configured
- PowerShell with permissions to use Outlook COM
- `7Zip4Powershell` module installed

**Español:**
- Windows con Outlook instalado y perfil MAPI configurado
- PowerShell con permisos para usar COM de Outlook
- Módulo `7Zip4Powershell` instalado

### Module Installation / Instalación del Módulo

```powershell
Install-Module 7Zip4Powershell -Force
```

---

## ⚙️ Parameters / Parámetros

**English:**
- The script **does not define formal parameters** (`param(...)`)
- Implicit inputs that affect execution:
  - Outlook folder selected manually via `PickFolder()` dialog
  - Only processes emails where `Unread -eq $true`
  - Attachments found in those emails
  - Compressed files detected with `*.*z*` filter

**Español:**
- El script **no define parámetros formales** (`param(...)`)
- Entradas implícitas que afectan su ejecución:
  - Carpeta de Outlook seleccionada manualmente en el diálogo `PickFolder()`
  - Solo procesa correos donde `Unread -eq $true`
  - Adjuntos encontrados en esos correos
  - Archivos comprimidos detectados con el filtro `*.*z*`

---

## 🔧 Behavior / Comportamiento

**English:**
- Creates temporary directory with pattern: `DMARC-CSV_yyyy-MM-dd_HH-mm`
- Saves attachments with format: `Sender--AttachmentName`
- Attempts to expand compressed files in `xml` folder
- If `*.xml.tar` files appear, renames them to `*.xml.tar.xml`
- Consolidates results in memory and exports CSV at the end
- **Closes Outlook** when executing: `Get-Process -Name OUTLOOK | Stop-Process`

**Español:**
- Crea directorio temporal con patrón: `DMARC-CSV_yyyy-MM-dd_HH-mm`
- Guarda adjuntos con formato: `Remitente--NombreAdjunto`
- Intenta expandir archivos comprimidos en carpeta `xml`
- Si aparecen archivos `*.xml.tar`, los renombra a `*.xml.tar.xml`
- Consolida resultados en memoria y exporta CSV al final
- **Cierra Outlook** al ejecutar: `Get-Process -Name OUTLOOK | Stop-Process`

> ⚠️ **Warning / Advertencia:** This behavior forcefully terminates the Outlook process / Este comportamiento finaliza el proceso de Outlook de forma forzada

---

## 📤 Output / Salida

**English:**

Automatically opens folder: `%TEMP%\DMARC-CSV_<date_time>\output`

**Generated files:**

**1. `filereport.csv`**
- One row per `feedback` block (report metadata)
- Typical columns:
  - `fileID`
  - `rm_org_name`, `rm_email`, `rm_report_id`
  - `rm_unixdate_begin`, `rm_unixdate_end`
  - `pp_domain`, `pp_adkim`, `pp_aspf`, `pp_p`, `pp_sp`, `pp_pct`

**2. `rowreport.csv`**
- One row per DMARC `record`
- Typical columns:
  - `fileID`
  - `row_ip`, `row_count`
  - `row_disposition`, `row_aligned_dkim`, `row_aligned_spf`
  - `row_reason_type`, `row_reason_comment`
  - `header_from`
  - `dkim_result`, `dkim_domain`, `dkim_selector`
  - `spf_result`, `spf_domain`

**Español:**

Abre automáticamente la carpeta: `%TEMP%\DMARC-CSV_<fecha_hora>\output`

**Archivos generados:**

**1. `filereport.csv`**
- Una fila por bloque `feedback` (metadatos del reporte)
- Columnas típicas:
  - `fileID`
  - `rm_org_name`, `rm_email`, `rm_report_id`
  - `rm_unixdate_begin`, `rm_unixdate_end`
  - `pp_domain`, `pp_adkim`, `pp_aspf`, `pp_p`, `pp_sp`, `pp_pct`

**2. `rowreport.csv`**
- Una fila por cada `record` DMARC
- Columnas típicas:
  - `fileID`
  - `row_ip`, `row_count`
  - `row_disposition`, `row_aligned_dkim`, `row_aligned_spf`
  - `row_reason_type`, `row_reason_comment`
  - `header_from`
  - `dkim_result`, `dkim_domain`, `dkim_selector`
  - `spf_result`, `spf_domain`

---

## 🚀 Usage / Uso

```powershell
.\_script_dmarc_outlook_to_csv.ps1
```

---

## ⚠️ Common Errors / Errores Comunes

**English:**
- If `7Zip4Powershell` is not installed, the script terminates with error
- If there are no unread emails or no valid attachments, CSVs may be empty
- Damaged or invalid XML files are skipped in `try/catch` blocks

**Español:**
- Si no está instalado `7Zip4Powershell`, el script termina con error
- Si no hay correos no leídos o no hay adjuntos válidos, los CSV pueden quedar vacíos
- XML dañados o no válidos se omiten en bloques `try/catch`

---

## 💬 Feedback

Any feedback is very valuable!  
¡Cualquier comentario es muy valioso!

**Gregorio Parra**
