# Documentación: _script_dmarc_outlook_to_csv.ps1

## Resumen
Este script automatiza la extracción y transformación de reportes DMARC recibidos por correo en Outlook.

Flujo general:
1. Verifica e importa el módulo `7Zip4Powershell`.
2. Crea un directorio de trabajo temporal en `%TEMP%` con subcarpetas `output` y `xml`.
3. Pide seleccionar una carpeta de Outlook (PickFolder).
4. Recorre los correos **no leídos** de esa carpeta y guarda sus adjuntos en el directorio temporal.
5. Descomprime archivos adjuntos comprimidos (`*.*z*`) para extraer XML DMARC.
6. Convierte la información XML en dos conjuntos de datos:
   - Metadatos por archivo/reporte.
   - Registros por fila (`record`) del reporte DMARC.
7. Exporta dos CSV y abre la carpeta de salida en el Explorador.

## Prerrequisitos
- Windows con Outlook instalado y perfil MAPI configurado.
- PowerShell con permisos para usar COM de Outlook.
- Módulo `7Zip4Powershell` instalado:

```powershell
Install-Module 7Zip4Powershell -Force
```

## Parámetros
El script **no define parámetros formales** (`param(...)`).

Entradas implícitas que afectan su ejecución:
- Carpeta de Outlook seleccionada manualmente en el diálogo `PickFolder()`.
- Solo procesa correos donde `Unread -eq $true`.
- Adjuntos encontrados en esos correos.
- Archivos comprimidos detectados con el filtro `*.*z*`.

## Comportamiento relevante
- Crea un directorio temporal con este patrón:
  - `DMARC-CSV_yyyy-MM-dd_HH-mm`
- Guarda adjuntos con el formato:
  - `Remitente--NombreAdjunto`
- Intenta expandir comprimidos en `xml`.
- Si aparecen archivos `*.xml.tar`, los renombra a `*.xml.tar.xml`.
- Consolida resultados en memoria y exporta CSV al final.
- Cierra Outlook al ejecutar:

```powershell
Get-Process -Name OUTLOOK | Stop-Process
```

> Nota: este comportamiento finaliza el proceso de Outlook de forma forzada.

## Salida esperada
Al finalizar, abre automáticamente la carpeta:

- `%TEMP%\DMARC-CSV_<fecha_hora>\output`

Archivos generados:

1. `filereport.csv`
   - Una fila por bloque `feedback` (metadatos del reporte).
   - Columnas típicas:
     - `fileID`
     - `rm_org_name`, `rm_email`, `rm_report_id`
     - `rm_unixdate_begin`, `rm_unixdate_end`
     - `pp_domain`, `pp_adkim`, `pp_aspf`, `pp_p`, `pp_sp`, `pp_pct`

2. `rowreport.csv`
   - Una fila por cada `record` DMARC.
   - Columnas típicas:
     - `fileID`
     - `row_ip`, `row_count`
     - `row_disposition`, `row_aligned_dkim`, `row_aligned_spf`
     - `row_reason_type`, `row_reason_comment`
     - `header_from`
     - `dkim_result`, `dkim_domain`, `dkim_selector`
     - `spf_result`, `spf_domain`

## Errores comunes
- Si no está instalado `7Zip4Powershell`, el script termina con error.
- Si no hay correos no leídos o no hay adjuntos válidos, los CSV pueden quedar vacíos.
- XML dañados o no válidos se omiten en bloques `try/catch`.

## Ejecución
```powershell
.\_script_dmarc_outlook_to_csv.ps1
```
