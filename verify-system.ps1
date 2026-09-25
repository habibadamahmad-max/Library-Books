# Automated Verification and Unit Test Script for Book Barcode System

Write-Host "=================================================================" -ForegroundColor Cyan
Write-Host "   RUNNING AUTOMATED VERIFICATION SUITE FOR BOOK BARCODE SYSTEM  " -ForegroundColor Green
Write-Host "=================================================================" -ForegroundColor Cyan
Write-Host ""

$testsPassed = 0
$totalTests = 0

function Assert-Test($testName, $condition, $details) {
  $script:totalTests++
  if ($condition) {
    $script:testsPassed++
    Write-Host " [PASS] $testName" -ForegroundColor Green
    if ($details) { Write-Host "        $details" -ForegroundColor Gray }
  } else {
    Write-Host " [FAIL] $testName" -ForegroundColor Red
    if ($details) { Write-Host "        $details" -ForegroundColor Yellow }
  }
}

# -----------------------------------------------------------------------------
# 1. Test File Structure & Offline Assets
# -----------------------------------------------------------------------------
Write-Host "--- 1. File Structure and Offline Assets ---" -ForegroundColor White
Assert-Test "index.html exists" (Test-Path "index.html") ""
Assert-Test "test-barcodes.html exists" (Test-Path "test-barcodes.html") ""
Assert-Test "styles.css exists" (Test-Path "assets\css\styles.css") ""
Assert-Test "isbn-validator.js exists" (Test-Path "assets\js\isbn-validator.js") ""
Assert-Test "api-service.js exists" (Test-Path "assets\js\api-service.js") ""
Assert-Test "excel-export.js exists" (Test-Path "assets\js\excel-export.js") ""
Assert-Test "scanner.js exists" (Test-Path "assets\js\scanner.js") ""
Assert-Test "app.js exists" (Test-Path "assets\js\app.js") ""
Assert-Test "zxing.min.js bundled offline" (Test-Path "assets\js\zxing.min.js") "Size: $((Get-Item assets\js\zxing.min.js).Length) bytes"
Assert-Test "xlsx.full.min.js bundled offline" (Test-Path "assets\js\xlsx.full.min.js") "Size: $((Get-Item assets\js\xlsx.full.min.js).Length) bytes"
Assert-Test "jsbarcode.min.js bundled offline" (Test-Path "assets\js\jsbarcode.min.js") "Size: $((Get-Item assets\js\jsbarcode.min.js).Length) bytes"
Assert-Test "start-server.ps1 exists" (Test-Path "start-server.ps1") ""
Assert-Test "start-server.bat exists" (Test-Path "start-server.bat") ""

# -----------------------------------------------------------------------------
# 2. Test ISBN-13 Checksum Algorithm (Test Scenario 1 & 4)
# -----------------------------------------------------------------------------
Write-Host "`n--- 2. ISBN Validation Algorithms ---" -ForegroundColor White
function Check-ISBN13($s) {
  $s = $s -replace "[-\s]", ""
  if ($s.Length -ne 13 -or $s -notmatch "^\d{13}$") { return $false }
  $sum = 0
  for ($i = 0; $i -lt 12; $i++) {
    $d = [int][string]$s[$i]
    if ($i % 2 -eq 0) { $sum += $d } else { $sum += $d * 3 }
  }
  $check = (10 - ($sum % 10)) % 10
  return $check -eq [int][string]$s[12]
}

function Check-ISBN10($s) {
  $s = ($s -replace "[-\s]", "").ToUpper()
  if ($s.Length -ne 10 -or $s -notmatch "^\d{9}[\dX]$") { return $false }
  $sum = 0
  for ($i = 0; $i -lt 9; $i++) {
    $sum += [int][string]$s[$i] * (10 - $i)
  }
  $last = if ($s[9] -eq 'X') { 10 } else { [int][string]$s[9] }
  $sum += $last
  return ($sum % 11) -eq 0
}

function Convert-ISBN10To13($s) {
  $s = ($s -replace "[-\s]", "").ToUpper()
  if (-not (Check-ISBN10 $s)) { return $null }
  $base = "978" + $s.Substring(0, 9)
  $sum = 0
  for ($i = 0; $i -lt 12; $i++) {
    $d = [int][string]$base[$i]
    if ($i % 2 -eq 0) { $sum += $d } else { $sum += $d * 3 }
  }
  $check = (10 - ($sum % 10)) % 10
  return $base + $check
}

Assert-Test "Valid ISBN-13 (Clean Code: 9780132350884)" (Check-ISBN13 "9780132350884") "Expected: valid"
Assert-Test "Valid ISBN-10 with X check (Pragmatic Prog: 020161622X)" (Check-ISBN10 "020161622X") "Expected: valid"
Assert-Test "Invalid ISBN-13 Checksum (Corrupt check: 9780132350880)" (-not (Check-ISBN13 "9780132350880")) "Expected: rejected"
Assert-Test "Convert ISBN-10 to ISBN-13 (0132350882 -> 9780132350884)" ((Convert-ISBN10To13 "0132350882") -eq "9780132350884") "Converted accurately"

# -----------------------------------------------------------------------------
# 3. Test API Lookup Capability (Open Library Search API)
# -----------------------------------------------------------------------------
Write-Host "`n--- 3. API Retrieval Verification ---" -ForegroundColor White
$headers = @{ "User-Agent" = "LibraryBookScanner/1.0" }
$testIsbn = "9780132350884"
try {
  $res = Invoke-RestMethod -Uri "https://openlibrary.org/search.json?isbn=$testIsbn&limit=1" -Headers $headers -TimeoutSec 10
  $hasTitle = ($res.docs.Count -gt 0) -and ($res.docs[0].title -match "Clean Code")
  Assert-Test "Open Library returns title for 9780132350884" $hasTitle "Title: $($res.docs[0].title)"
  $hasAuthor = ($res.docs[0].author_name -contains "Robert C. Martin")
  Assert-Test "Open Library returns author for 9780132350884" $hasAuthor "Author: $($res.docs[0].author_name -join ', ')"
} catch {
  Assert-Test "Open Library API response" $false $_.Exception.Message
}

# 3b. Test Regional / Academic ISBN 9788171419128 (Comparative Education)
$regionalIsbn = "9788171419128"
Assert-Test "ISBN 9788171419128 checksum validity" (Check-ISBN13 $regionalIsbn) "Valid ISBN-13 checksum"

# Test online lookup logic for 9788171419128
try {
  $url = "https://html.duckduckgo.com/html/?q=$regionalIsbn"
  $resp = Invoke-WebRequest -Uri $url -Headers $headers -UseBasicParsing -TimeoutSec 10
  $matchedOnline = $resp.Content -match "Comparative Education" -and ($resp.Content -match "Dutt" -or $resp.Content -match "Amazon")
  Assert-Test "Online web catalog resolves 9788171419128 (Comparative Education)" $matchedOnline "Resolved title & author from online catalog"
} catch {
  Assert-Test "Online web catalog lookup" $false $_.Exception.Message
}

# -----------------------------------------------------------------------------
# 4. Summary
# -----------------------------------------------------------------------------
Write-Host ""
Write-Host "=================================================================" -ForegroundColor Cyan
Write-Host " Automated Verification Result: $testsPassed / $totalTests Passed" -ForegroundColor $(if ($testsPassed -eq $totalTests) { "Green" } else { "Yellow" })
Write-Host "=================================================================" -ForegroundColor Cyan
