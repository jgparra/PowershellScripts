<#
    .NOTES
    ===========================================================================
     Created with:   PowerShell 5+, VS Code
     Created on:     2026-06-02
     Created by:     Gregorio Parra - (gregorio.parra@microsoft.com)

     Organization:   Microsoft

     Filename:       Analyze-TlsCapturesforCertificates.ps1
    ===========================================================================

    .DESCRIPTION
        Esta herramienta no es soportada oficialmente por Microsoft.
        Comentarios y sugerencias: gregorio.parra@microsoft.com

        Analiza capturas PCAP/PCAPNG de forma offline para identificar
        handshakes TLS/SSL, puertos usados, destinos solicitados y datos
        de certificados observados durante la negociacion.

        Objetivo principal:
        - Investigacion rapida de trafico HTTPS/TLS en escenarios con y sin proxy explicito
        - Identificacion de puertos relevantes (priorizando 443 y puertos tipicos de proxy)
        - Extraccion de Subject, Issuer, vigencia y Thumbprint desde certificados capturados
        - Correlacion por flujo TCP (stream) con evidencia de SNI y/o CONNECT

        Alcance actual:
        - Escenario optimizado para proxy explicito (CONNECT + TLS)
        - Soporta entrada por archivo unico o carpeta completa
        - Reporte en consola y export opcional (TXT/JSON/CSV)

        Limitaciones conocidas:
        - No realiza conexiones activas a Internet (modo offline)
        - En sesiones truncadas o incompletas puede no existir certificado visible
        - Deteccion de proxy transparente se considera fase posterior (heuristica)

        Flujo general:
        1) Carga de captura(s) y validacion de tshark
        2) Extraccion de eventos TLS handshake y CONNECT HTTP
        3) Correlacion por tcp.stream
        4) Priorizacion de puertos:
           - Primero 443
           - Luego puertos tipicos de proxy (8080, 3128, 3129, 8000, 8888, 8118)
           - Finalmente otros puertos detectados
        5) Generacion de reporte consolidado

        Versiones:
        V 1.0  2026-06-02 - Version inicial funcional para analisis de proxy explicito.
        V 1.1  2026-06-02 - Mejora de extraccion de certificados por stream y normalizacion de CONNECT host.

    .PARAMETER InputFile
        Ruta de un archivo .pcap o .pcapng para analisis individual.

    .PARAMETER InputPath
        Ruta de carpeta que contiene uno o mas archivos .pcap/.pcapng.

    .PARAMETER ProxyPorts
        Lista de puertos considerados tipicos de proxy para priorizacion en reporte.
        Default: 8080, 3128, 3129, 8000, 8888, 8118

    .PARAMETER NonInteractive
        Ejecuta sin prompts interactivos (por ejemplo, para pipelines/automatizacion).

    .PARAMETER ExportFormats
        Formatos de exportacion opcionales: txt, json, csv.
        Si no se especifica, el script pregunta de forma interactiva.

    .PARAMETER OutputDirectory
        Carpeta de salida para exportes. Si no se define, usa ./reports.

    .EXAMPLE
        .\Analyze-TlsCapturesforCertificates.ps1 -InputFile .\tls_3128_prx_TLS_inspect_capture.pcap

    .EXAMPLE
        .\Analyze-TlsCapturesforCertificates.ps1 -InputPath .\captures

    .EXAMPLE
        .\Analyze-TlsCapturesforCertificates.ps1 -InputPath . -NonInteractive

    .EXAMPLE
        .\Analyze-TlsCapturesforCertificates.ps1 -InputPath . -ExportFormats json,csv -OutputDirectory .\reports

#>
[CmdletBinding(DefaultParameterSetName = 'ByPath')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'ByFile')]
    [ValidateNotNullOrEmpty()]
    [string]$InputFile,

    [Parameter(Mandatory = $true, ParameterSetName = 'ByPath')]
    [ValidateNotNullOrEmpty()]
    [string]$InputPath,

    [Parameter()]
    [int[]]$ProxyPorts = @(8080, 3128, 3129, 8000, 8888, 8118),

    [Parameter()]
    [switch]$NonInteractive,

    [Parameter()]
    [ValidateSet('txt', 'json', 'csv')]
    [string[]]$ExportFormats = @(),

    [Parameter()]
    [string]$OutputDirectory = ''
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Resolve tshark from PATH first, then from common Windows install paths.
# This keeps the script portable across machines where Wireshark CLI is installed
# but not necessarily added to environment variables.
function Find-Tshark {
    $command = Get-Command tshark -ErrorAction SilentlyContinue
    if ($command) {
        return $command.Source
    }

    $candidates = @(
        'C:\Program Files\Wireshark\tshark.exe',
        'C:\Program Files (x86)\Wireshark\tshark.exe'
    )

    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath $candidate) {
            return $candidate
        }
    }

    return $null
}

# Build the capture file workload.
# ByFile mode analyzes only one capture; ByPath mode scans a folder.
function Get-CaptureFiles {
    param(
        [string]$Mode,
        [string]$FilePath,
        [string]$FolderPath
    )

    if ($Mode -eq 'ByFile') {
        if (-not (Test-Path -LiteralPath $FilePath)) {
            throw "InputFile not found: $FilePath"
        }

        $item = Get-Item -LiteralPath $FilePath
        if ($item.Extension -notin @('.pcap', '.pcapng')) {
            throw "InputFile must be .pcap or .pcapng: $FilePath"
        }

        return @($item)
    }

    if (-not (Test-Path -LiteralPath $FolderPath)) {
        throw "InputPath not found: $FolderPath"
    }

    $files = Get-ChildItem -LiteralPath $FolderPath -File | Where-Object {
        $_.Extension -in @('.pcap', '.pcapng')
    }

    if (-not $files) {
        throw "No .pcap or .pcapng files were found in: $FolderPath"
    }

    return @($files | Sort-Object Name)
}

# Generic tshark field extractor.
# Always returns an array of rows (possibly empty) so downstream parsing
# can iterate safely without extra null checks.
function Invoke-TsharkFields {
    param(
        [string]$TsharkPath,
        [string]$CaptureFile,
        [string]$DisplayFilter,
        [string[]]$Fields
    )

    $tsharkCmdParts = @('-r', $CaptureFile, '-Y', $DisplayFilter, '-T', 'fields')
    foreach ($field in $Fields) {
        $tsharkCmdParts += @('-e', $field)
    }

    # '/t' tells tshark to use tab as separator in field output.
    $tsharkCmdParts += @('-E', 'separator=/t', '-E', 'quote=n')

    $output = & $TsharkPath @tsharkCmdParts
    if ($LASTEXITCODE -ne 0) {
        throw "tshark failed for file: $CaptureFile"
    }

    if (-not $output) {
        return @()
    }

    return @($output)
}

# Utility parser for comma-separated tshark fields (e.g., handshake types).
function Split-ListField {
    param([string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return @()
    }

    return @($Value -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' })
}

# Convert hex string payloads into byte arrays for certificate decoding.
function Convert-HexToBytes {
    param([string]$Hex)

    if ([string]::IsNullOrWhiteSpace($Hex)) {
        return $null
    }

    $clean = ($Hex -replace '[^0-9A-Fa-f]', '')
    if (($clean.Length % 2) -ne 0) {
        return $null
    }

    $bytes = New-Object byte[] ($clean.Length / 2)
    for ($i = 0; $i -lt $bytes.Length; $i++) {
        $bytes[$i] = [Convert]::ToByte($clean.Substring($i * 2, 2), 16)
    }

    return $bytes
}

# Best-effort certificate decoder from tshark hex fields.
# Returns first decodable certificate found in the field content.
function Get-CertificateInfoFromHexField {
    param([string]$CertificateField)

    $hexEntries = Split-ListField -Value $CertificateField
    if (-not $hexEntries) {
        return $null
    }

    foreach ($hexEntry in $hexEntries) {
        try {
            $bytes = Convert-HexToBytes -Hex $hexEntry
            if (-not $bytes) {
                continue
            }

            $cert = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($bytes)
            return [PSCustomObject]@{
                Subject    = $cert.Subject
                Issuer     = $cert.Issuer
                NotBefore  = $cert.NotBefore
                NotAfter   = $cert.NotAfter
                Thumbprint = $cert.Thumbprint
            }
        }
        catch {
            continue
        }
    }

    return $null
}

# Host resolution priority in reports:
# 1) SNI from ClientHello, 2) CONNECT host, 3) server IP fallback.
function Get-TargetHost {
    param($Session)

    if (-not [string]::IsNullOrWhiteSpace($Session.SNI)) {
        return $Session.SNI
    }

    if (-not [string]::IsNullOrWhiteSpace($Session.ConnectHost)) {
        return $Session.ConnectHost
    }

    if (-not [string]::IsNullOrWhiteSpace($Session.ServerIP)) {
        return $Session.ServerIP
    }

    return 'unknown-target'
}

# Normalize CONNECT host values to plain lowercase hostnames.
# Handles raw host:port and URI-style values.
function Normalize-ConnectHost {
    param(
        [string]$HostValue,
        [string]$UriValue
    )

    $candidate = if (-not [string]::IsNullOrWhiteSpace($HostValue)) { $HostValue.Trim() } else { $UriValue.Trim() }
    if ([string]::IsNullOrWhiteSpace($candidate)) {
        return ''
    }

    if ($candidate -match '^https?://([^/:]+)') {
        return $matches[1].ToLowerInvariant()
    }

    if ($candidate -match '^([^/:]+):\d+$') {
        return $matches[1].ToLowerInvariant()
    }

    return $candidate.ToLowerInvariant()
}

# Second-pass certificate extraction per TCP stream.
# Some captures expose certs in stream-level rows where first pass may miss
# or not map them during mixed handshake packet parsing.
function Populate-CertificateByStream {
    param(
        [string]$TsharkPath,
        [System.IO.FileInfo]$FileInfo,
        [object]$Session
    )

    if ($Session.CertificateThumbprint) {
        return
    }

    $filter = "tcp.stream==$($Session.Stream) && tls.handshake.certificate"
    $certLines = Invoke-TsharkFields -TsharkPath $TsharkPath -CaptureFile $FileInfo.FullName -DisplayFilter $filter -Fields @('tls.handshake.certificate')
    foreach ($certLine in $certLines) {
        if ([string]::IsNullOrWhiteSpace($certLine)) {
            continue
        }

        try {
            $firstHex = ($certLine -split ',')[0]
            $clean = ($firstHex -replace '[^0-9A-Fa-f]', '')
            if ([string]::IsNullOrWhiteSpace($clean) -or ($clean.Length % 2) -ne 0) {
                continue
            }

            $bytes = New-Object byte[] ($clean.Length / 2)
            for ($i = 0; $i -lt $bytes.Length; $i++) {
                $bytes[$i] = [Convert]::ToByte($clean.Substring($i * 2, 2), 16)
            }

            $x509 = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($bytes)
            $Session.CertificateSubject = $x509.Subject
            $Session.CertificateIssuer = $x509.Issuer
            $Session.CertificateNotBefore = $x509.NotBefore
            $Session.CertificateNotAfter = $x509.NotAfter
            $Session.CertificateThumbprint = $x509.Thumbprint
            [void]$Session.Evidence.Add('CERT')
            break
        }
        catch {
            continue
        }
    }
}

# Session model used for correlation by tcp.stream.
# Keeps raw evidence and normalized output fields in a single object.
function New-SessionRecord {
    param(
        [string]$FileName,
        [string]$FilePath,
        [string]$Stream
    )

    return [PSCustomObject]@{
        FileName              = $FileName
        FilePath              = $FilePath
        Stream                = $Stream
        ClientIP              = ''
        ClientPort            = 0
        ServerIP              = ''
        ServerPort            = 0
        SNI                   = ''
        ConnectHost           = ''
        ConnectUri            = ''
        HandshakeTypes        = New-Object System.Collections.Generic.HashSet[string]
        Evidence              = New-Object System.Collections.Generic.HashSet[string]
        CertificateSubject    = ''
        CertificateIssuer     = ''
        CertificateNotBefore  = $null
        CertificateNotAfter   = $null
        CertificateThumbprint = ''
        Status                = 'Partial'
    }
}

# Classify each session result for reporting.
function Update-SessionStatus {
    param($Session)

    if ($Session.CertificateThumbprint) {
        $Session.Status = 'CertificateExtracted'
        return
    }

    if ($Session.HandshakeTypes.Contains('1')) {
        $Session.Status = 'HandshakeObservedNoCert'
        return
    }

    if ($Session.Evidence.Count -gt 0) {
        $Session.Status = 'EvidenceWithoutHandshake'
        return
    }

    $Session.Status = 'NoUsefulEvidence'
}

# Core analyzer per capture file.
# Correlates TLS handshake data and CONNECT metadata by tcp.stream.
function Analyze-CaptureFile {
    param(
        [string]$TsharkPath,
        [System.IO.FileInfo]$FileInfo
    )

    $tlsFields = @(
        'tcp.stream',
        'ip.src',
        'tcp.srcport',
        'ip.dst',
        'tcp.dstport',
        'tls.handshake.type',
        'tls.handshake.extensions_server_name',
        'tls.handshake.certificate'
    )

    $connectFields = @(
        'tcp.stream',
        'ip.src',
        'tcp.srcport',
        'ip.dst',
        'tcp.dstport',
        'http.host',
        'http.request.full_uri'
    )

    # TLS handshake view for SNI, handshake type, and certificate presence.
    $tlsLines = Invoke-TsharkFields -TsharkPath $TsharkPath -CaptureFile $FileInfo.FullName -DisplayFilter 'tls.handshake' -Fields $tlsFields
    # CONNECT view for explicit-proxy metadata and target host inference.
    $connectLines = Invoke-TsharkFields -TsharkPath $TsharkPath -CaptureFile $FileInfo.FullName -DisplayFilter 'http.request.method == "CONNECT"' -Fields $connectFields

    $sessions = @{}

    foreach ($line in $tlsLines) {
        $parts = @($line -split "`t", 8)
        while ($parts.Count -lt 8) {
            $parts += ''
        }

        $stream = $parts[0]
        if ([string]::IsNullOrWhiteSpace($stream)) {
            continue
        }

        if (-not $sessions.ContainsKey($stream)) {
            $sessions[$stream] = New-SessionRecord -FileName $FileInfo.Name -FilePath $FileInfo.FullName -Stream $stream
        }

        $session = $sessions[$stream]

        $ipSrc = $parts[1]
        $srcPortRaw = $parts[2]
        $ipDst = $parts[3]
        $dstPortRaw = $parts[4]
        $handshakeTypesRaw = $parts[5]
        $sniRaw = $parts[6]
        $certificateRaw = $parts[7]

        $srcPort = 0
        $dstPort = 0
        [void][int]::TryParse($srcPortRaw, [ref]$srcPort)
        [void][int]::TryParse($dstPortRaw, [ref]$dstPort)

        $types = Split-ListField -Value $handshakeTypesRaw
        foreach ($t in $types) {
            [void]$session.HandshakeTypes.Add($t)
        }

        # Handshake type 1 = ClientHello, usually client -> server direction.
        if ($types -contains '1') {
            if (-not $session.ClientIP) {
                $session.ClientIP = $ipSrc
                $session.ClientPort = $srcPort
            }

            if (-not $session.ServerIP) {
                $session.ServerIP = $ipDst
                $session.ServerPort = $dstPort
            }

            if (-not [string]::IsNullOrWhiteSpace($sniRaw)) {
                $session.SNI = $sniRaw
                [void]$session.Evidence.Add('SNI')
            }
        }

        # Handshake type 11 = Certificate, often server -> client direction.
        if (($types -contains '11') -and (-not $session.ServerIP)) {
            $session.ServerIP = $ipSrc
            $session.ServerPort = $srcPort
        }

        # First-pass certificate decode from inline handshake fields.
        if ($certificateRaw -and (-not $session.CertificateThumbprint)) {
            $cert = Get-CertificateInfoFromHexField -CertificateField $certificateRaw
            if ($cert) {
                $session.CertificateSubject = $cert.Subject
                $session.CertificateIssuer = $cert.Issuer
                $session.CertificateNotBefore = $cert.NotBefore
                $session.CertificateNotAfter = $cert.NotAfter
                $session.CertificateThumbprint = $cert.Thumbprint
                [void]$session.Evidence.Add('Certificate')
            }
        }

        if (-not $session.ServerPort) {
            if ($types -contains '1') {
                $session.ServerPort = $dstPort
            }
            elseif ($types -contains '11') {
                $session.ServerPort = $srcPort
            }
        }
    }

    # Enrich sessions with explicit proxy context (CONNECT host/uri).
    foreach ($line in $connectLines) {
        $parts = @($line -split "`t", 7)
        while ($parts.Count -lt 7) {
            $parts += ''
        }

        $stream = $parts[0]
        if ([string]::IsNullOrWhiteSpace($stream)) {
            continue
        }

        if (-not $sessions.ContainsKey($stream)) {
            $sessions[$stream] = New-SessionRecord -FileName $FileInfo.Name -FilePath $FileInfo.FullName -Stream $stream
        }

        $session = $sessions[$stream]

        if (-not $session.ClientIP) {
            $session.ClientIP = $parts[1]
            $clientPort = 0
            [void][int]::TryParse($parts[2], [ref]$clientPort)
            $session.ClientPort = $clientPort
        }

        if (-not $session.ServerIP) {
            $session.ServerIP = $parts[3]
            $serverPort = 0
            [void][int]::TryParse($parts[4], [ref]$serverPort)
            $session.ServerPort = $serverPort
        }

        $connectHostValue = $parts[5]
        $uri = $parts[6]
        $normalizedConnectHost = Normalize-ConnectHost -HostValue $connectHostValue -UriValue $uri

        if (-not [string]::IsNullOrWhiteSpace($normalizedConnectHost)) {
            $session.ConnectHost = $normalizedConnectHost
            [void]$session.Evidence.Add('CONNECT')
        }

        if (-not [string]::IsNullOrWhiteSpace($uri)) {
            $session.ConnectUri = $uri
            [void]$session.Evidence.Add('CONNECT')
        }
    }

    # Final pass: complete certificate data and derive session status.
    $streamRecords = @($sessions.Values)
    foreach ($record in $streamRecords) {
        Populate-CertificateByStream -TsharkPath $TsharkPath -FileInfo $FileInfo -Session $record
        Update-SessionStatus -Session $record
    }

    return $streamRecords
}

# Sort priority used by report sections:
# 443 first, then known proxy ports, then any other discovered ports.
function Get-OrderedPorts {
    param(
        [object[]]$Sessions,
        [int[]]$ProxyPortPriority
    )

    $ports = @($Sessions | Where-Object { $_.ServerPort -gt 0 } | Select-Object -ExpandProperty ServerPort -Unique)

    $ordered = New-Object System.Collections.Generic.List[int]

    if ($ports -contains 443) {
        [void]$ordered.Add(443)
    }

    foreach ($proxyPort in $ProxyPortPriority) {
        if (($ports -contains $proxyPort) -and (-not $ordered.Contains($proxyPort))) {
            [void]$ordered.Add($proxyPort)
        }
    }

    $remaining = @($ports | Where-Object { -not $ordered.Contains($_) } | Sort-Object)
    foreach ($port in $remaining) {
        [void]$ordered.Add([int]$port)
    }

    return @($ordered)
}

# Map each port to its ordering index for stable report rendering.
function Get-PortRank {
    param(
        [int]$Port,
        [int[]]$PortOrder
    )

    for ($i = 0; $i -lt $PortOrder.Count; $i++) {
        if ($PortOrder[$i] -eq $Port) {
            return $i
        }
    }

    return 9999
}

# Human-readable console report with three blocks:
# connectivity, certificate details, and proxy traffic summary.
function Write-ConsoleReport {
    param(
        [object[]]$AllSessions,
        [int[]]$PortOrder,
        [string[]]$CaptureNames
    )

        $portList = @($PortOrder | Where-Object { $null -ne $_ })

    Write-Host ''
    Write-Host '=== TLS Handshake Analysis (Offline PCAP) ==='
    Write-Host '[Mode: OFFLINE + TSHARK]'
    Write-Host ("Captures analyzed: {0}" -f ($CaptureNames -join ', '))

    Write-Host ''
    Write-Host '=== Port Discovery Order ==='
        if ($portList.Count -eq 0) {
        Write-Host 'No TLS handshake ports discovered.'
    }
    else {
            Write-Host ("443 first, proxy-typical next, others last: {0}" -f (($portList | ForEach-Object { [string]$_ }) -join ', '))
    }

    Write-Host ''
    Write-Host '=== Connectivity Observed (From Captures) ==='
    if (-not $AllSessions) {
        Write-Host 'No sessions were detected.'
    }
    else {
        $sortedConnectivity = $AllSessions | Sort-Object @{ Expression = { Get-PortRank -Port $_.ServerPort -PortOrder $portList } }, @{ Expression = { $_.ServerPort } }, FileName, Stream
        foreach ($session in $sortedConnectivity) {
            $targetHostValue = Get-TargetHost -Session $session
            $port = if ($session.ServerPort -gt 0) { $session.ServerPort } else { 'unknown-port' }
            $state = if ($session.HandshakeTypes.Contains('1')) { 'HANDSHAKE_OK' } else { 'NO_CLIENT_HELLO' }
            Write-Host ("{0}:{1} : {2} (file={3}, stream={4})" -f $targetHostValue, $port, $state, $session.FileName, $session.Stream)
        }
    }

    Write-Host ''
    Write-Host '=== Certificate & TLS Inspection ==='
    if (-not $AllSessions) {
        Write-Host 'No certificate data available.'
    }
    else {
        $sortedCerts = $AllSessions | Sort-Object @{ Expression = { Get-PortRank -Port $_.ServerPort -PortOrder $portList } }, @{ Expression = { $_.ServerPort } }, FileName, Stream
        foreach ($session in $sortedCerts) {
            $targetHostValue = Get-TargetHost -Session $session
            $port = if ($session.ServerPort -gt 0) { $session.ServerPort } else { 'unknown-port' }

            Write-Host ''
            Write-Host ("Host: {0}:{1}" -f $targetHostValue, $port)
            Write-Host ("  Capture     : {0}" -f $session.FileName)
            Write-Host ("  Stream      : {0}" -f $session.Stream)
            Write-Host ("  Evidence    : {0}" -f (($session.Evidence | Sort-Object) -join ', '))
            Write-Host ("  Status      : {0}" -f $session.Status)

            if ($session.CertificateThumbprint) {
                Write-Host ("  Subject     : {0}" -f $session.CertificateSubject)
                Write-Host ("  Issuer      : {0}" -f $session.CertificateIssuer)
                Write-Host ("  Not Before  : {0}" -f $session.CertificateNotBefore)
                Write-Host ("  Not After   : {0}" -f $session.CertificateNotAfter)
                Write-Host ("  Thumbprint  : {0}" -f $session.CertificateThumbprint)
            }
            else {
                Write-Host '  Certificate : Not visible in capture or not decodable.'
            }
        }
    }

    Write-Host ''
    Write-Host '=== Proxy Traffic Summary ==='
    $proxyCandidates = $AllSessions | Where-Object {
        ($_.ServerPort -ne 443) -and ($_.ServerPort -gt 0)
    }

    if (-not $proxyCandidates) {
        Write-Host 'No non-443 TLS sessions found.'
        return
    }

    $sortedProxy = $proxyCandidates | Sort-Object @{ Expression = { Get-PortRank -Port $_.ServerPort -PortOrder $portList } }, @{ Expression = { $_.ServerPort } }, FileName, Stream
    foreach ($session in $sortedProxy) {
        $targetHostValue = Get-TargetHost -Session $session
        Write-Host ("Port {0} -> target={1} evidence={2} file={3} stream={4}" -f $session.ServerPort, $targetHostValue, (($session.Evidence | Sort-Object) -join ','), $session.FileName, $session.Stream)
    }
}

# Build export-friendly payload (summary + flat session rows).
function Build-ExportPayload {
    param(
        [object[]]$AllSessions,
        [int[]]$PortOrder,
        [System.IO.FileInfo[]]$CaptureFiles
    )

    $sessionRows = foreach ($session in $AllSessions) {
        $targetHost = Get-TargetHost -Session $session
        [PSCustomObject]@{
            capture_file      = $session.FileName
            stream            = $session.Stream
            client_ip         = $session.ClientIP
            client_port       = $session.ClientPort
            server_ip         = $session.ServerIP
            server_port       = $session.ServerPort
            target_host       = $targetHost
            sni               = $session.SNI
            connect_host      = $session.ConnectHost
            connect_uri       = $session.ConnectUri
            handshake_types   = (($session.HandshakeTypes | Sort-Object) -join ',')
            evidence          = (($session.Evidence | Sort-Object) -join ',')
            cert_subject      = $session.CertificateSubject
            cert_issuer       = $session.CertificateIssuer
            cert_not_before   = $session.CertificateNotBefore
            cert_not_after    = $session.CertificateNotAfter
            cert_thumbprint   = $session.CertificateThumbprint
            status            = $session.Status
        }
    }

    $summary = [PSCustomObject]@{
        generated_at_utc = (Get-Date).ToUniversalTime().ToString('o')
        mode             = 'OFFLINE + TSHARK'
        capture_count    = $CaptureFiles.Count
        capture_files    = @($CaptureFiles | ForEach-Object { $_.FullName })
        discovered_ports = @($PortOrder)
        session_count    = $AllSessions.Count
    }

    return [PSCustomObject]@{
        summary  = $summary
        sessions = @($sessionRows)
    }
}

# Persist report files in selected output formats.
function Export-Reports {
    param(
        [object]$Payload,
        [string[]]$Formats,
        [string]$Directory
    )

    if (-not (Test-Path -LiteralPath $Directory)) {
        [void](New-Item -ItemType Directory -Path $Directory)
    }

    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'

    if ($Formats -contains 'json') {
        $jsonPath = Join-Path $Directory ("tls_report_{0}.json" -f $timestamp)
        $Payload | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $jsonPath -Encoding UTF8
        Write-Host ("Saved JSON: {0}" -f $jsonPath)
    }

    if ($Formats -contains 'csv') {
        $csvPath = Join-Path $Directory ("tls_sessions_{0}.csv" -f $timestamp)
        $Payload.sessions | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding UTF8
        Write-Host ("Saved CSV: {0}" -f $csvPath)
    }

    if ($Formats -contains 'txt') {
        $txtPath = Join-Path $Directory ("tls_summary_{0}.txt" -f $timestamp)
        $lines = New-Object System.Collections.Generic.List[string]
        $lines.Add('=== TLS Handshake Analysis (Offline PCAP) ===')
        $lines.Add('[Mode: OFFLINE + TSHARK]')
        $lines.Add(("Generated UTC: {0}" -f $Payload.summary.generated_at_utc))
        $lines.Add(("Capture count: {0}" -f $Payload.summary.capture_count))
        $lines.Add(("Discovered ports: {0}" -f (($Payload.summary.discovered_ports | ForEach-Object { $_.ToString() }) -join ', ')))
        $lines.Add('')

        foreach ($row in $Payload.sessions) {
            $lines.Add(("Host: {0}:{1}" -f $row.target_host, $row.server_port))
            $lines.Add(("  Capture     : {0}" -f $row.capture_file))
            $lines.Add(("  Stream      : {0}" -f $row.stream))
            $lines.Add(("  Subject     : {0}" -f $row.cert_subject))
            $lines.Add(("  Issuer      : {0}" -f $row.cert_issuer))
            $lines.Add(("  Not Before  : {0}" -f $row.cert_not_before))
            $lines.Add(("  Not After   : {0}" -f $row.cert_not_after))
            $lines.Add(("  Thumbprint  : {0}" -f $row.cert_thumbprint))
            $lines.Add(("  Evidence    : {0}" -f $row.evidence))
            $lines.Add(("  Status      : {0}" -f $row.status))
            $lines.Add('')
        }

        Set-Content -LiteralPath $txtPath -Value $lines -Encoding UTF8
        Write-Host ("Saved TXT: {0}" -f $txtPath)
    }
}

$tsharkPath = Find-Tshark
if (-not $tsharkPath) {
    throw 'tshark was not found. Install Wireshark/tshark or add tshark to PATH.'
}

Write-Host ("Using tshark: {0}" -f $tsharkPath)

$captureFiles = Get-CaptureFiles -Mode $PSCmdlet.ParameterSetName -FilePath $InputFile -FolderPath $InputPath
$allSessions = New-Object System.Collections.Generic.List[object]

foreach ($capture in $captureFiles) {
    # File-level processing keeps tracing clear when reviewing mixed datasets.
    Write-Host ("Analyzing capture: {0}" -f $capture.FullName)
    $sessionsForFile = Analyze-CaptureFile -TsharkPath $tsharkPath -FileInfo $capture
    foreach ($session in $sessionsForFile) {
        [void]$allSessions.Add($session)
    }
}

$allSessionsArray = @($allSessions.ToArray())
$portOrder = Get-OrderedPorts -Sessions $allSessionsArray -ProxyPortPriority $ProxyPorts

Write-ConsoleReport -AllSessions $allSessionsArray -PortOrder $portOrder -CaptureNames @($captureFiles | ForEach-Object { $_.Name })

$payload = Build-ExportPayload -AllSessions $allSessionsArray -PortOrder $portOrder -CaptureFiles $captureFiles

$shouldExport = $false
$chosenFormats = @()

if ($ExportFormats.Count -gt 0) {
    # Non-interactive export path when formats are provided as parameters.
    $shouldExport = $true
    $chosenFormats = @($ExportFormats)
}
elseif (-not $NonInteractive) {
    $answer = Read-Host 'Do you want to export report files (txt/json/csv)? [y/N]'
    if ($answer -match '^(y|yes)$') {
        $shouldExport = $true
        $formatInput = Read-Host 'Choose formats separated by comma (txt,json,csv). Enter for txt,json'
        if ([string]::IsNullOrWhiteSpace($formatInput)) {
            $chosenFormats = @('txt', 'json')
        }
        else {
            $chosenFormats = @(
                $formatInput -split ',' |
                ForEach-Object { $_.Trim().ToLowerInvariant() } |
                Where-Object { $_ -in @('txt', 'json', 'csv') } |
                Select-Object -Unique
            )

            if (-not $chosenFormats) {
                Write-Host 'No valid formats selected. Skipping export.'
                $shouldExport = $false
            }
        }
    }
}

if ($shouldExport) {
    if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
        $OutputDirectory = Join-Path (Get-Location) 'reports'
    }

    Export-Reports -Payload $payload -Formats $chosenFormats -Directory $OutputDirectory
}
else {
    Write-Host 'Export skipped.'
}
