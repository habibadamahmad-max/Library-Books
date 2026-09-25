<#
.SYNOPSIS
  Lightweight Local Web Server & Online Book Lookup Proxy for Book Barcode Scanner System
.DESCRIPTION
  Runs an HTTP server on http://localhost:8080 using .NET HttpListener.
  Ensures browser treats the origin as a Secure Context for camera access (getUserMedia).
  Includes /api/lookup endpoint that queries Open Library, Google Books, and online web catalogs (Amazon, DirectTextbook, etc.)
  to discover metadata for any ISBN globally.
#>

param(
  [int]$Port = 8080
)

$baseDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $baseDir) {
  $baseDir = Get-Location
}

$listener = New-Object System.Net.HttpListener
$prefix = "http://localhost:$Port/"
$listener.Prefixes.Add($prefix)

try {
  $listener.Start()
} catch {
  Write-Warning "Port $Port may be in use. Trying port 8088..."
  $Port = 8088
  $prefix = "http://localhost:$Port/"
  $listener = New-Object System.Net.HttpListener
  $listener.Prefixes.Add($prefix)
  $listener.Start()
}

Write-Host ""
Write-Host "=================================================================" -ForegroundColor Cyan
Write-Host "   Book Barcode Scanner and Excel Export System Local Server    " -ForegroundColor Green
Write-Host "=================================================================" -ForegroundColor Cyan
Write-Host " [OK] Server running at: $prefix" -ForegroundColor White
Write-Host " [OK] Serving root:      $baseDir" -ForegroundColor Gray
Write-Host " [OK] Camera access:     Enabled (Secure Localhost Context)" -ForegroundColor Green
Write-Host " [OK] Online ISBN API:   Active at /api/lookup?isbn=..." -ForegroundColor Green
Write-Host "-----------------------------------------------------------------" -ForegroundColor Cyan
Write-Host " Press Ctrl+C in this console window to stop the server." -ForegroundColor Yellow
Write-Host "=================================================================" -ForegroundColor Cyan
Write-Host ""

# Open default browser
Start-Process $prefix

$mimeTypes = @{
  ".html" = "text/html; charset=utf-8"
  ".htm"  = "text/html; charset=utf-8"
  ".css"  = "text/css; charset=utf-8"
  ".js"   = "application/javascript; charset=utf-8"
  ".json" = "application/json; charset=utf-8"
  ".png"  = "image/png"
  ".jpg"  = "image/jpeg"
  ".jpeg" = "image/jpeg"
  ".gif"  = "image/gif"
  ".svg"  = "image/svg+xml"
  ".ico"  = "image/x-icon"
  ".xlsx" = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  ".csv"  = "text/csv; charset=utf-8"
  ".txt"  = "text/plain; charset=utf-8"
}

# Helper: calculate ISBN-13 check digit
function Get-ISBN13CheckDigit($first12) {
  $sum = 0
  for ($i = 0; $i -lt 12; $i++) {
    $d = [int][string]$first12[$i]
    if ($i % 2 -eq 0) { $sum += $d } else { $sum += $d * 3 }
  }
  return (10 - ($sum % 10)) % 10
}

# Helper: pad a short code into candidate ISBNs for searching
function Get-ISBN-Candidates($raw) {
  $clean = $raw -replace "[^0-9Xx]", "" 
  $candidates = [System.Collections.Generic.List[string]]::new()
  $candidates.Add($clean)
  
  if ($clean.Length -lt 10) {
    # Try SBN (9-digit) and short codes padded to ISBN-10 and ISBN-13
    $padded10 = $clean.PadLeft(10, '0')
    $sbn = '0' + $clean
    $candidates.Add($padded10)
    $candidates.Add($sbn)
    # Build ISBN-13 from padded-10
    $base978 = "978" + $padded10.Substring(0, 9)
    $check13 = Get-ISBN13CheckDigit $base978
    $candidates.Add($base978 + $check13)
    $base978sbn = "978" + $sbn.Substring(0, 9)
    $check13sbn = Get-ISBN13CheckDigit $base978sbn
    $candidates.Add($base978sbn + $check13sbn)
  }
  
  if ($clean.Length -eq 10) {
    $base978 = "978" + $clean.Substring(0, 9)
    $check13 = Get-ISBN13CheckDigit $base978
    $candidates.Add($base978 + $check13)
  }

  if ($clean.Length -eq 12) {
    $check13 = Get-ISBN13CheckDigit $clean
    $candidates.Add($clean + $check13)
  }

  if ($clean.Length -eq 13) {
    # derive ISBN-10 if starts with 978
    if ($clean.StartsWith("978")) {
      $base9 = $clean.Substring(3, 9)
      $sum = 0
      for ($i = 0; $i -lt 9; $i++) { $sum += [int][string]$base9[$i] * (10 - $i) }
      $rem = (11 - ($sum % 11)) % 11
      $check10 = if ($rem -eq 10) { 'X' } else { "$rem" }
      $candidates.Add($base9 + $check10)
    }
  }

  return ($candidates | Select-Object -Unique)
}

# Universal online book lookup engine - searches all candidate ISBNs
function Search-OnlineBook($isbn) {
  $headers = @{ "User-Agent" = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36" }
  $candidates = Get-ISBN-Candidates $isbn

  foreach ($cand in $candidates) {
    if ([string]::IsNullOrWhiteSpace($cand) -or $cand.Length -lt 4) { continue }

    # --- Open Library Search ---
    try {
      $ol = Invoke-RestMethod -Uri "https://openlibrary.org/search.json?isbn=$cand&limit=1" -Headers $headers -TimeoutSec 4
      if ($ol.docs.Count -gt 0) {
        $d = $ol.docs[0]
        $pub = if ($d.publisher) { $d.publisher[0] } else { "Not available" }
        $date = if ($d.publish_date) { $d.publish_date[0] } elseif ($d.first_publish_year) { "$($d.first_publish_year)" } else { "Not available" }
        $cat = if ($d.subject) { ($d.subject | Select-Object -First 3) -join ", " } else { "Not available" }
        $cover = if ($d.cover_i) { "https://covers.openlibrary.org/b/id/$($d.cover_i)-M.jpg" } else { "https://covers.openlibrary.org/b/isbn/$cand-M.jpg" }
        return @{
          found = $true; isbn = $cand
          title = $d.title
          author = if ($d.author_name) { $d.author_name -join ", " } else { "Not available" }
          publisher = $pub; publicationDate = $date; category = $cat; coverUrl = $cover
          source = "Open Library"
        }
      }
    } catch {}

    # --- Open Library Direct ISBN ---
    try {
      $oli = Invoke-RestMethod -Uri "https://openlibrary.org/isbn/$cand.json" -Headers $headers -TimeoutSec 4
      if ($oli.title) {
        $pub = if ($oli.publishers) { $oli.publishers[0] } else { "Not available" }
        $date = if ($oli.publish_date) { $oli.publish_date } else { "Not available" }
        return @{
          found = $true; isbn = $cand
          title = $oli.title; author = "Not available"
          publisher = $pub; publicationDate = $date; category = "Not available"
          coverUrl = "https://covers.openlibrary.org/b/isbn/$cand-M.jpg"
          source = "Open Library ISBN"
        }
      }
    } catch {}

    # --- Web Catalog Search (DuckDuckGo -> Amazon, DirectTextbook, etc.) ---
    try {
      $url = "https://html.duckduckgo.com/html/?q=$cand+book+ISBN"
      $resp = Invoke-WebRequest -Uri $url -Headers $headers -UseBasicParsing -TimeoutSec 7
      $html = $resp.Content

      # Decode HTML entities
      $html = $html -replace '&#x27;', "'" -replace '&amp;', '&' -replace '&quot;', '"' -replace '&#39;', "'"

      # Pattern A: "Title: Author: ISBN: Amazon" (book detail page)
      if ($html -match 'class="result__title"[\s\S]*?<a[\s\S]*?>([^:<]{3,80}):\s*([^:<]{3,60}):\s*[\d-]{10,17}:\s*Amazon') {
        $title = $Matches[1].Trim() -replace '<[^>]+>', ''
        $author = $Matches[2].Trim() -replace '<[^>]+>', ''
        if ($title.Length -gt 2 -and $author.Length -gt 2) {
          # Try to extract publisher from snippet
          $pub = "Not available"
          $date = "Not available"
          if ($html -match '(\d{4})') { $date = $Matches[1] }
          return @{
            found = $true; isbn = $cand
            title = $title; author = $author
            publisher = $pub; publicationDate = $date; category = "Not available"
            coverUrl = "https://covers.openlibrary.org/b/isbn/$cand-M.jpg"
            source = "Online Web Catalog (Amazon)"
          }
        }
      }

      # Pattern B: "Find <ISBN-text> <Title> by <Author> at over 30 bookstores"
      if ($html -match 'Find\s+\d+\s+([\w\s:,\.\-\&]{4,80}?)\s+by\s+([\w\s,\.\-]{3,60}?)\s+at over') {
        $title = $Matches[1].Trim()
        $author = $Matches[2].Trim()
        if ($title.Length -gt 2 -and $author.Length -gt 2) {
          $date = "Not available"
          if ($html -match '(\d{4})') { $date = $Matches[1] }
          return @{
            found = $true; isbn = $cand
            title = $title; author = $author
            publisher = "Not available"; publicationDate = $date; category = "Not available"
            coverUrl = "https://covers.openlibrary.org/b/isbn/$cand-M.jpg"
            source = "Online Web Catalog (DirectTextbook)"
          }
        }
      }

      # Pattern C: Amazon UK style "Buy <Title> by <Author> (ISBN: ...)"
      if ($html -match 'Buy\s+([\w\s:,\.\-\&]{4,80}?)\s+by\s+([\w\s,\.\-]{3,60}?)\s*\(ISBN') {
        $title = $Matches[1].Trim()
        $author = $Matches[2].Trim()
        if ($title.Length -gt 2 -and $author.Length -gt 2) {
          $date = "Not available"
          if ($html -match '(\d{4})') { $date = $Matches[1] }
          return @{
            found = $true; isbn = $cand
            title = $title; author = $author
            publisher = "Not available"; publicationDate = $date; category = "Not available"
            coverUrl = "https://covers.openlibrary.org/b/isbn/$cand-M.jpg"
            source = "Online Web Catalog (Amazon UK)"
          }
        }
      }

      # Pattern D: Generic - extract any <title> tag from results that looks like a book
      if ($html -match '<a\s[^>]*class="result__a"[^>]*>([^<]{5,100})<\/a>') {
        $t = $Matches[1].Trim()
        if ($t -notmatch '^\s*$' -and $t -notmatch '^(Amazon|Google|Wikipedia|Search)') {
          $date = "Not available"
          if ($html -match '(\d{4})') { $date = $Matches[1] }
          return @{
            found = $true; isbn = $cand
            title = $t; author = "Not available"
            publisher = "Not available"; publicationDate = $date; category = "Not available"
            coverUrl = "https://covers.openlibrary.org/b/isbn/$cand-M.jpg"
            source = "Online Web Catalog"
          }
        }
      }
    } catch {}
  }

  return @{ found = $false; isbn = $isbn }
}

try {
  while ($listener.IsListening) {
    $context = $listener.GetContext()
    $request = $context.Request
    $response = $context.Response

    $urlPath = $request.Url.LocalPath.TrimStart('/')

    # -------------------------------------------------------------------------
    # REST API Route: /api/lookup?isbn=...
    # -------------------------------------------------------------------------
    if ($urlPath -eq "api/lookup") {
      $isbnParam = $request.QueryString["isbn"]
      $response.ContentType = "application/json; charset=utf-8"
      $response.Headers.Add("Access-Control-Allow-Origin", "*")
      $response.Headers.Add("Cache-Control", "no-cache")

      if ([string]::IsNullOrWhiteSpace($isbnParam)) {
        $response.StatusCode = 400
        $errJson = @{ success = $false; error = "Missing 'isbn' query parameter" } | ConvertTo-Json
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($errJson)
        $response.OutputStream.Write($bytes, 0, $bytes.Length)
        $response.Close()
        continue
      }

      $bookResult = Search-OnlineBook $isbnParam
      $resJson = @{
        success = $bookResult.found
        book = if ($bookResult.found) { $bookResult } else { $null }
      } | ConvertTo-Json

      $bytes = [System.Text.Encoding]::UTF8.GetBytes($resJson)
      $response.StatusCode = 200
      $response.OutputStream.Write($bytes, 0, $bytes.Length)
      $response.Close()
      continue
    }

    # -------------------------------------------------------------------------
    # Static File Serving
    # -------------------------------------------------------------------------
    if ([string]::IsNullOrWhiteSpace($urlPath)) {
      $urlPath = "index.html"
    }

    $decodedPath = [System.Uri]::UnescapeDataString($urlPath)
    $filePath = Join-Path $baseDir $decodedPath

    $fullPath = [System.IO.Path]::GetFullPath($filePath)
    if (-not $fullPath.StartsWith([System.IO.Path]::GetFullPath($baseDir), [System.StringComparison]::OrdinalIgnoreCase)) {
      $response.StatusCode = 403
      $msg = [System.Text.Encoding]::UTF8.GetBytes("403 Forbidden")
      $response.OutputStream.Write($msg, 0, $msg.Length)
      $response.Close()
      continue
    }

    if (Test-Path $fullPath -PathType Leaf) {
      $ext = [System.IO.Path]::GetExtension($fullPath).ToLower()
      $mime = $mimeTypes[$ext]
      if (-not $mime) {
        $mime = "application/octet-stream"
      }

      $response.ContentType = $mime
      $response.Headers.Add("Access-Control-Allow-Origin", "*")
      $response.Headers.Add("Cache-Control", "no-cache")

      try {
        $bytes = [System.IO.File]::ReadAllBytes($fullPath)
        $response.ContentLength64 = $bytes.Length
        $response.OutputStream.Write($bytes, 0, $bytes.Length)
        $response.StatusCode = 200
      } catch {
        $response.StatusCode = 500
      }
    } else {
      $response.StatusCode = 404
      $msg = [System.Text.Encoding]::UTF8.GetBytes("404 File Not Found: $decodedPath")
      $response.ContentType = "text/plain; charset=utf-8"
      $response.OutputStream.Write($msg, 0, $msg.Length)
    }

    $response.Close()
  }
} finally {
  $listener.Stop()
  $listener.Close()
}
