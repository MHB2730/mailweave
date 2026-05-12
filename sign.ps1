# MailWeave — code-signing helper.
#
# Signs one or more .exe files with the configured certificate.
#
# Usage:
#   .\sign.ps1 path\to\file.exe [more.exe ...]
#
# Configuration:
#   The thumbprint can come from:
#     1) The MAILWEAVE_SIGN_THUMBPRINT environment variable (preferred for CI).
#     2) The fallback below (current self-signed dev cert).
#   When you receive a real OV/EV code-signing cert, install it into
#   Cert:\CurrentUser\My (or Cert:\LocalMachine\My) and update the fallback or
#   set the env var.
#
# Exits 0 on success, non-zero on failure.

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, ValueFromRemainingArguments = $true)]
    [string[]]$Files
)

$ErrorActionPreference = 'Stop'

$thumbprint = $env:MAILWEAVE_SIGN_THUMBPRINT
if (-not $thumbprint) {
    $thumbprint = '46055FEA693CD6EB3A27676231B941E7400F3B7C'  # Self-signed dev cert
}

$timestampUrl = 'http://timestamp.digicert.com'
$description  = 'MailWeave'

# Locate signtool.exe — prefer the latest x64 build under the Windows 10 SDK.
$signtool = Get-ChildItem -Path 'C:\Program Files (x86)\Windows Kits\10\bin' `
    -Filter 'signtool.exe' -Recurse -ErrorAction SilentlyContinue `
    | Where-Object { $_.FullName -match '\\x64\\signtool\.exe$' } `
    | Sort-Object FullName -Descending `
    | Select-Object -First 1 -ExpandProperty FullName

if (-not $signtool) {
    Write-Error 'signtool.exe not found. Install the Windows 10/11 SDK.'
    exit 2
}

# Verify the cert is present in the user's certificate store.
$cert = Get-ChildItem "Cert:\CurrentUser\My\$thumbprint" -ErrorAction SilentlyContinue
if (-not $cert) {
    $cert = Get-ChildItem "Cert:\LocalMachine\My\$thumbprint" -ErrorAction SilentlyContinue
}
if (-not $cert) {
    Write-Error "Signing certificate $thumbprint not found in CurrentUser\My or LocalMachine\My."
    exit 3
}

Write-Host "Signing with: $($cert.Subject) (thumbprint $thumbprint)"

$failed = 0
foreach ($file in $Files) {
    if (-not (Test-Path $file)) {
        Write-Warning "Skipping (not found): $file"
        $failed++
        continue
    }

    Write-Host "  -> $file"
    & $signtool sign `
        /sha1 $thumbprint `
        /fd SHA256 `
        /tr $timestampUrl /td SHA256 `
        /d $description `
        $file
    if ($LASTEXITCODE -ne 0) {
        Write-Warning "Signing failed for $file (exit $LASTEXITCODE)"
        $failed++
    }
}

if ($failed -gt 0) {
    Write-Error "$failed file(s) failed to sign."
    exit 1
}

Write-Host "All files signed successfully."
exit 0
