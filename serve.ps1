param(
    [int]$Port = 5500
)

$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://localhost:$Port/")
$listener.Start()
Write-Host "Local web server running at http://localhost:$Port/"

$root = $PSScriptRoot

$mimeMap = @{
    ".html" = "text/html; charset=utf-8"
    ".css"  = "text/css; charset=utf-8"
    ".js"   = "application/javascript; charset=utf-8"
    ".json" = "application/json; charset=utf-8"
    ".png"  = "image/png"
    ".jpg"  = "image/jpeg"
    ".svg"  = "image/svg+xml"
    ".pdf"  = "application/pdf"
    ".docx" = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    ".pptx" = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
}

try {
    while ($listener.IsListening) {
        $context = $listener.GetContext()
        try {
            $req = $context.Request
            $res = $context.Response

            $path = $req.Url.LocalPath
            if ($path -eq "/" -or [string]::IsNullOrWhiteSpace($path)) {
                $path = "/index.html"
            }

            $decodedPath = [System.Uri]::UnescapeDataString($path).TrimStart('/')
            $localFile = Join-Path $root $decodedPath

            if (Test-Path $localFile -PathType Leaf) {
                $ext = [System.IO.Path]::GetExtension($localFile).ToLower()
                $contentType = if ($mimeMap.ContainsKey($ext)) { $mimeMap[$ext] } else { "application/octet-stream" }
                
                $bytes = [System.IO.File]::ReadAllBytes($localFile)
                $res.ContentType = $contentType
                $res.ContentLength64 = $bytes.Length
                $res.Headers.Add("Access-Control-Allow-Origin", "*")

                if ($req.HttpMethod -ne "HEAD") {
                    $res.OutputStream.Write($bytes, 0, $bytes.Length)
                }
            } else {
                $res.StatusCode = 404
                $msg = [System.Text.Encoding]::UTF8.GetBytes("404 Not Found")
                $res.OutputStream.Write($msg, 0, $msg.Length)
            }
        } catch {
            # Client abort or write error
        } finally {
            try { $context.Response.Close() } catch {}
        }
    }
} finally {
    $listener.Stop()
}
