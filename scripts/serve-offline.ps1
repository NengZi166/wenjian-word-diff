param([switch]$NoBrowser)

$ErrorActionPreference = 'Stop'
$siteRoot = [System.IO.Path]::GetFullPath($PSScriptRoot)
$indexPath = Join-Path $siteRoot 'index.html'

if (-not (Test-Path -LiteralPath $indexPath -PathType Leaf)) {
    Write-Host 'Offline site files are missing. Please extract the complete package.' -ForegroundColor Red
    exit 1
}

$listener = $null
$port = 8765
$started = $false

while (-not $started -and $port -le 8775) {
    try {
        $listener = [System.Net.Sockets.TcpListener]::new([System.Net.IPAddress]::Loopback, $port)
        $listener.Start()
        $started = $true
    } catch {
        if ($listener) { $listener.Stop() }
        $listener = $null
        $port += 1
    }
}

if (-not $started) {
    Write-Host 'Could not start the local server. Free one port from 8765 to 8775 and retry.' -ForegroundColor Red
    exit 1
}

$url = "http://127.0.0.1:$port/"
Write-Host ''
Write-Host 'Wenjian is running.' -ForegroundColor Green
Write-Host "Browser address: $url"
Write-Host 'Keep this window open. Closing it stops Wenjian.'
Write-Host ''
if (-not $NoBrowser) {
    try {
        Start-Process $url
    } catch {
        Write-Host 'The browser could not open automatically.' -ForegroundColor Yellow
        Write-Host "Open this address manually: $url" -ForegroundColor Yellow
        Write-Host ''
    }
}

$mimeTypes = @{
    '.html' = 'text/html; charset=utf-8'
    '.js'   = 'text/javascript; charset=utf-8'
    '.css'  = 'text/css; charset=utf-8'
    '.json' = 'application/json; charset=utf-8'
    '.wasm' = 'application/wasm'
    '.dll'  = 'application/octet-stream'
    '.dat'  = 'application/octet-stream'
    '.blat' = 'application/octet-stream'
    '.woff' = 'font/woff'
    '.woff2'= 'font/woff2'
    '.ttf'  = 'font/ttf'
    '.svg'  = 'image/svg+xml'
    '.png'  = 'image/png'
    '.ico'  = 'image/x-icon'
}

function Write-Response {
    param(
        [System.Net.Sockets.NetworkStream]$Stream,
        [int]$StatusCode,
        [string]$StatusText,
        [string]$ContentType,
        [byte[]]$Body
    )
    $header = "HTTP/1.1 $StatusCode $StatusText`r`nContent-Type: $ContentType`r`nContent-Length: $($Body.Length)`r`nCache-Control: no-store`r`nConnection: close`r`n`r`n"
    $headerBytes = [System.Text.Encoding]::ASCII.GetBytes($header)
    $Stream.Write($headerBytes, 0, $headerBytes.Length)
    if ($Body.Length -gt 0) { $Stream.Write($Body, 0, $Body.Length) }
    $Stream.Flush()
}

try {
    while ($true) {
        $client = $listener.AcceptTcpClient()
        $stream = $null
        $reader = $null
        try {
            $stream = $client.GetStream()
            $reader = [System.IO.StreamReader]::new($stream, [System.Text.Encoding]::ASCII, $false, 8192, $true)
            $requestLine = $reader.ReadLine()
            while ($null -ne ($line = $reader.ReadLine()) -and $line -ne '') { }

            if ([string]::IsNullOrWhiteSpace($requestLine)) {
                Write-Response -Stream $stream -StatusCode 400 -StatusText 'Bad Request' -ContentType 'text/plain' -Body ([byte[]]@())
                continue
            }

            $requestParts = $requestLine.Split(' ')
            if ($requestParts.Length -lt 2 -or $requestParts[0] -ne 'GET') {
                Write-Response -Stream $stream -StatusCode 405 -StatusText 'Method Not Allowed' -ContentType 'text/plain' -Body ([byte[]]@())
                continue
            }

            $rawPath = $requestParts[1].Split('?')[0]
            $relativePath = [System.Uri]::UnescapeDataString($rawPath).TrimStart('/')
            if ([string]::IsNullOrWhiteSpace($relativePath)) { $relativePath = 'index.html' }
            $candidate = [System.IO.Path]::GetFullPath((Join-Path $siteRoot $relativePath))

            if (-not $candidate.StartsWith($siteRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
                Write-Response -Stream $stream -StatusCode 403 -StatusText 'Forbidden' -ContentType 'text/plain' -Body ([byte[]]@())
                continue
            }

            if (-not (Test-Path -LiteralPath $candidate -PathType Leaf)) { $candidate = $indexPath }
            $extension = [System.IO.Path]::GetExtension($candidate).ToLowerInvariant()
            $contentType = $mimeTypes[$extension]
            if (-not $contentType) { $contentType = 'application/octet-stream' }
            $bytes = [System.IO.File]::ReadAllBytes($candidate)
            Write-Response -Stream $stream -StatusCode 200 -StatusText 'OK' -ContentType $contentType -Body $bytes
        } catch {
            if ($stream) {
                try { Write-Response -Stream $stream -StatusCode 500 -StatusText 'Internal Server Error' -ContentType 'text/plain' -Body ([byte[]]@()) } catch { }
            }
        } finally {
            if ($reader) { $reader.Dispose() }
            if ($stream) { $stream.Dispose() }
            $client.Close()
        }
    }
} finally {
    if ($listener) { $listener.Stop() }
}
