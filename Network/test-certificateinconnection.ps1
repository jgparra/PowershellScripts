<#
	.NOTES
	===========================================================================
	 Created with: 	PowerShell 5+, VS Code
	 Created on:   	2026-06-01
	 Created by:   	Gregorio Parra - (gregorio.parra@microsoft.com)

	 Organization: 	Microsoft

	 Filename:     	test-certificateinconnection.ps1
	===========================================================================

	.DESCRIPTION
		Esta herramienta no es soportada oficialmente por Microsoft.
		Comentarios y sugerencias: gregorio.parra@microsoft.com

		Valida la conectividad TCP y certificados SSL/TLS de endpoints criticos
		de Microsoft 365 y cloud.microsoft. Permite detectar intercepcion TLS
		por parte de proxies o appliances de seguridad corporativos.

		Funcionalidad:
		- Prueba conectividad TCP puerto 443 en 6 endpoints de M365/cloud
		- Extrae certificados: Subject, Issuer, Vigencia, Thumbprint
		- Modo NORMAL: alerta si el emisor es un proxy conocido (Netskope, Zscaler, etc.)
		- Modo STRICT: alerta si el emisor NO es DigiCert o Microsoft
		- Inspecciona headers HTTP para identificar indicadores de proxy

		Endpoints evaluados:
		- substrate.office.com       (Exchange Online / Outlook)
		- login.microsoftonline.com  (Azure AD / Entra ID)
		- m365.cloud.microsoft       (Microsoft 365 Copilot - endpoint empresarial)
		- outlook.cloud.microsoft    (Outlook nuevo endpoint)
		- teams.cloud.microsoft      (Teams nuevo endpoint)
		- graph.microsoft.com        (Microsoft Graph API)

		Versiones:
		V 1.0  2026-06-01 - Version inicial. Dual-mode TLS interception detection.

	.PARAMETER StrictMode
		Activa validacion estricta de emisores de certificados.
		$false (default): solo alerta por proxies conocidos
		$true           : solo acepta DigiCert o Microsoft como emisores validos

	.PARAMETER ProxyServer
		Nombre o IP del servidor proxy (opcional).
		Si se especifica, todas las conexiones se realizan a través del proxy.
		$null o "" (default): conexión directa sin proxy

	.PARAMETER ProxyPort
		Puerto del servidor proxy (default: 8080).
		Solo se usa si ProxyServer está especificado.

	.EXAMPLE
		# Modo normal - solo alerta si detecta proxy conocido
		.\test-certificateinconnection.ps1

	.EXAMPLE
		# Modo strict - alerta si el emisor no es DigiCert o Microsoft
		& ".\test-certificateinconnection.ps1" -StrictMode $true

	.EXAMPLE
		# Conexión a través de proxy corporativo
		& ".\test-certificateinconnection.ps1" -ProxyServer "proxy.empresa.com" -ProxyPort 8080

	.EXAMPLE
		# Conexión a través de proxy en modo strict
		& ".\test-certificateinconnection.ps1" -StrictMode $true -ProxyServer "10.0.0.1" -ProxyPort 3128

#>

param(
    [bool]$StrictMode = $false,
    [string]$ProxyServer = "",
    [int]$ProxyPort = 8080
)

# Forzar uso de TLS 1.2 para todas las conexiones HTTP del proceso
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Emisores de certificados considerados legítimos para M365 (Strict Mode)
# Cualquier emisor fuera de esta lista generará alerta en modo STRICT
$authorizedIssuers = @('DigiCert', 'Microsoft')

# Proxies/appliances conocidos que interceptan TLS (Modo NORMAL)
# Si el campo Issuer del certificado contiene alguno de estos strings, se genera alerta
$suspiciousIssuers = @('Netskope', 'Zscaler', 'Fortinet', 'Palo Alto', 'Cisco', 'Squid', 'Blue Coat', 'Symantec', 'McAfee', 'Trend Micro', 'F5 Networks')

# Lista maestra de endpoints - usada en las 3 fases (TCP, certificados, headers HTTP)
# Para agregar o quitar endpoints, modificar SOLO esta lista.
$endpoints = @(
    'substrate.office.com',       # Exchange Online / Outlook
    'login.microsoftonline.com',  # Azure AD / Entra ID
    'm365.cloud.microsoft',       # Microsoft 365 Copilot (endpoint empresarial)
    'm365copilot.com',            # Microsoft 365 Copilot (dominio alternativo)
    'outlook.cloud.microsoft',    # Outlook nuevo endpoint cloud.microsoft
    'teams.cloud.microsoft',      # Teams nuevo endpoint cloud.microsoft
    'graph.microsoft.com'         # Microsoft Graph API
)

Write-Host ""
Write-Host "=== TCP Connectivity Test (Port 443) ===" -ForegroundColor Cyan
$modeInfo = "Mode: $(if ($StrictMode) { 'STRICT' } else { 'NORMAL' })"
if ($ProxyServer) {
    $modeInfo += " | Proxy: $ProxyServer`:$ProxyPort"
}
Write-Host "[$modeInfo]" -ForegroundColor Yellow

# Fase 1: Test de conectividad TCP
# Verifica que el puerto 443 sea alcanzable antes de intentar extraer certificados.
# Un fallo aquí indica bloqueo de red (firewall, DNS, routing) - no necesariamente proxy.
foreach ($ep in $endpoints) {
    if ($ProxyServer) {
        # Si se especifica proxy, hacer CONNECT tunnel a través del proxy
        try {
            $tcpClient = New-Object System.Net.Sockets.TcpClient
            $tcpClient.Connect($ProxyServer, $ProxyPort)
            
            # Enviar comando CONNECT para tunelizar a través del proxy
            $stream = $tcpClient.GetStream()
            $connectCmd = "CONNECT $($ep):443 HTTP/1.1`r`nHost: $($ep):443`r`nConnection: close`r`n`r`n"
            $buffer = [System.Text.Encoding]::ASCII.GetBytes($connectCmd)
            $stream.Write($buffer, 0, $buffer.Length)
            $stream.Flush()
            
            # Leer respuesta (esperar "HTTP/1." y código de estado)
            $reader = New-Object System.IO.StreamReader($stream)
            $response = $reader.ReadLine()
            
            $status = if ($response -match "HTTP/1\.\d\s+200") { "OK (via proxy)" } else { "FAILED" }
            $color = if ($response -match "HTTP/1\.\d\s+200") { "Green" } else { "Red" }
            
            Write-Host "$ep : $status" -ForegroundColor $color
            $stream.Close()
            $tcpClient.Close()
        }
        catch {
            Write-Host "$ep : FAILED (proxy error)" -ForegroundColor Red
        }
    }
    else {
        # Conexión directa sin proxy
        $result = Test-NetConnection -ComputerName $ep -Port 443 -ErrorAction SilentlyContinue
        $status = if ($result.TcpTestSucceeded) { "OK" } else { "FAILED" }
        Write-Host "$ep : $status" -ForegroundColor $(if ($result.TcpTestSucceeded) { "Green" } else { "Red" })
    }
}

Write-Host ""
Write-Host "=== Certificate & TLS Inspection (via openssl) ===" -ForegroundColor Cyan

# Fase 2: Inspección de certificados TLS
# Se deriva host:443 desde $endpoints automáticamente
foreach ($hostPort in ($endpoints | ForEach-Object { "$_`:443" })) {
    Write-Host ""
    Write-Host "Host: $hostPort" -ForegroundColor Yellow
    
    # Detectar si openssl está disponible en el PATH del sistema
    # openssl es más confiable para extraer el certificado completo incluyendo la cadena
    $opensslPath = (Get-Command openssl -ErrorAction SilentlyContinue).Path
    
    if ($opensslPath) {
        try {
            # Conexión TLS con openssl: envía 'Q' para cerrar la sesión inmediatamente
            # grep filtra solo las líneas relevantes del certificado (subject, issuer, fechas)
            $certOutput = echo "Q" | openssl s_client -connect $hostPort -servername $($hostPort -split ':')[0] 2>&1 | grep -E "(subject=|issuer=|Issuer:|Subject:|before|after)"
            
            if ($certOutput) {
                Write-Host "  Certificate Details (via openssl):"
                $certOutput | ForEach-Object {
                    Write-Host ("    {0}" -f $_)
                }
            }
        }
        catch {
            Write-Host "  OpenSSL extraction failed: $($_.Exception.Message)"
        }
    }
    else {
        # Fallback: extracción de certificado via PowerShell puro (sin openssl)
        # Usa SslStream para hacer handshake TLS y obtener el certificado remoto
        # El callback {$true} acepta cualquier certificado (solo para lectura, no valida confianza)
        try {
            $targetHost = $($hostPort -split ':')[0]
            $port = 443
            
            if ($ProxyServer) {
                # Conexión a través de proxy con CONNECT tunneling
                $tcpClient = New-Object System.Net.Sockets.TcpClient
                $tcpClient.Connect($ProxyServer, $ProxyPort)
                $networkStream = $tcpClient.GetStream()
                
                # Enviar CONNECT command
                $connectCmd = "CONNECT $($targetHost):$port HTTP/1.1`r`nHost: $($targetHost):$port`r`nConnection: close`r`n`r`n"
                $buffer = [System.Text.Encoding]::ASCII.GetBytes($connectCmd)
                $networkStream.Write($buffer, 0, $buffer.Length)
                $networkStream.Flush()
                
                # Leer respuesta del CONNECT
                $reader = New-Object System.IO.StreamReader($networkStream)
                $response = $reader.ReadLine()
                
                if ($response -notmatch "HTTP/1\.\d\s+200") {
                    Write-Host ("  Proxy CONNECT failed: {0}" -f $response) -ForegroundColor Red
                    $networkStream.Close()
                    $tcpClient.Close()
                    continue
                }
                
                # Leer hasta línea vacía (fin de headers)
                while ($true) {
                    $line = $reader.ReadLine()
                    if ([string]::IsNullOrEmpty($line)) { break }
                }
                
                # Ahora hacer SSL handshake sobre el tunnel del proxy
                $sslStream = New-Object System.Net.Security.SslStream($networkStream, $false, {$true})
                $sslStream.AuthenticateAsClient($targetHost)
            }
            else {
                # Conexión directa
                $tcpClient = New-Object System.Net.Sockets.TcpClient
                $tcpClient.Connect($targetHost, $port)
                $sslStream = New-Object System.Net.Security.SslStream($tcpClient.GetStream(), $false, {$true})
                $sslStream.AuthenticateAsClient($targetHost)
            }
            
            # Convertir el certificado raw a X509Certificate2 para acceder a sus propiedades
            $remoteCert = $sslStream.RemoteCertificate
            $x509Cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($remoteCert)
            
            # Mostrar datos del certificado presentado por el servidor remoto
            Write-Host "  Subject     : $($x509Cert.Subject)" -ForegroundColor Cyan
            Write-Host "  Issuer      : $($x509Cert.Issuer)" -ForegroundColor Cyan
            Write-Host "  Not Before  : $($x509Cert.NotBefore)"
            Write-Host "  Not After   : $($x509Cert.NotAfter)"
            Write-Host "  Thumbprint  : $($x509Cert.Thumbprint)"
            
            # Lógica de evaluación del emisor según el modo activo
            $issuerAlert = $false
            $alertMessage = ""
            
            if ($StrictMode) {
                # STRICT MODE: verifica que el emisor esté en la lista $authorizedIssuers
                # Cualquier otra CA (incluso legítima) generará alerta
                $isAuthorized = $false
                foreach ($authorized in $authorizedIssuers) {
                    if ($x509Cert.Issuer -match $authorized) {
                        $isAuthorized = $true
                        break
                    }
                }
                
                if (-not $isAuthorized) {
                    $issuerAlert = $true
                    $alertMessage = "STRICT MODE: Emisor no autorizado - $($x509Cert.Issuer)"
                }
            }
            else {
                # NORMAL MODE: busca en $suspiciousIssuers usando regex (-match)
                # Un match aquí indica intercepción TLS activa por proxy/appliance
                foreach ($suspicious in $suspiciousIssuers) {
                    if ($x509Cert.Issuer -match $suspicious) {
                        $issuerAlert = $true
                        $alertMessage = "TLS INTERCEPTION DETECTED: $suspicious"
                        break
                    }
                }
            }
            
            # Mostrar resultado de la evaluación del emisor
            if ($issuerAlert) {
                Write-Host ("  ⚠️  {0}" -f $alertMessage) -ForegroundColor Red
            }
            elseif (-not $StrictMode -and ($x509Cert.Issuer -match "Microsoft|DigiCert")) {
                Write-Host "  ✓ Legitimate Microsoft Certificate" -ForegroundColor Green
            }
            elseif ($StrictMode) {
                Write-Host "  ✓ Authorized Issuer (Strict Mode OK)" -ForegroundColor Green
            }
            
            # Liberar recursos de red
            $sslStream.Close()
            if ($ProxyServer) {
                # Si usamos proxy, tcpClient ya está siendo usado por el reader
                # pero stream ya está cerrado por sslStream
            } else {
                $tcpClient.Close()
            }
        }
        catch {
            Write-Host ("  Failed to retrieve certificate: {0}" -f $_.Exception.Message) -ForegroundColor Red
        }
    }
}

Write-Host ""
Write-Host "=== HTTP Headers Inspection ===" -ForegroundColor Cyan

# Fase 3: Inspección de headers HTTP
# Se deriva https://<endpoint> desde $endpoints automáticamente
# El filtro busca headers que suelen revelar presencia de proxy:
#   - Proxy-*       : headers estándar de proxy (RFC 7230)
#   - X-NanoProxy   : indicador de gateway de Microsoft Exchange Online
#   - X-*           : headers customizados que revelan infraestructura
#   - Server        : identifica el servidor o appliance que responde
foreach ($url in ($endpoints | ForEach-Object { "https://$_" })) {
    Write-Host ""
    Write-Host "URL: $url" -ForegroundColor Yellow
    try {
        # UseBasicParsing evita dependencia de Internet Explorer DOM parser
        $params = @{
            Uri = $url
            UseBasicParsing = $true
            ErrorAction = 'Stop'
        }
        
        # Si se especifica proxy, agregarlo a los parámetros
        if ($ProxyServer) {
            $proxyUri = "http://$($ProxyServer):$ProxyPort"
            $params['Proxy'] = $proxyUri
        }
        
        $response = Invoke-WebRequest @params
        Write-Host "  HTTP Status: $($response.StatusCode)" -ForegroundColor Green
        Write-Host "  Headers (proxy/server related):"
        # Filtrar solo headers relevantes para detección de proxy/infraestructura
        $response.Headers.GetEnumerator() | Where-Object { $_.Key -match '(Proxy|X-|Server)' } | ForEach-Object {
            $headerValue = if ($_.Value -is [array]) { ($_.Value | Select-Object -First 1) } else { $_.Value }
            Write-Host ("    {0}: {1}" -f $_.Key, $headerValue)
        }
    }
    catch {
        Write-Host ("  Error: {0}" -f $_.Exception.Message) -ForegroundColor Red
    }
}