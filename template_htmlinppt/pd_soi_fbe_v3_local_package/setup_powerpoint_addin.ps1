$ErrorActionPreference = 'Stop'

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$manifest = Join-Path $scriptDir 'pd_soi_fbe_powerpoint_manifest.xml'
$cert = Join-Path $scriptDir 'localhost_selfsigned.crt'
$addinId = '5f248b07-f278-4ac4-9d1f-8d7fc55eb501'

if (-not (Test-Path -LiteralPath $manifest)) { throw "Manifest not found: $manifest" }
if (-not (Test-Path -LiteralPath $cert)) { throw "Certificate not found: $cert" }

Write-Host 'Trusting localhost certificate for current Windows user...'
& certutil -user -addstore Root $cert | Out-Host

Write-Host 'Registering the PowerPoint content add-in manifest for current user...'
Write-Host "Manifest: $manifest"
reg add 'HKCU\Software\Microsoft\Office\16.0\WEF\Developer' /v $addinId /t REG_SZ /d $manifest /f | Out-Host

Write-Host 'Done. Restart PowerPoint if it was already open.'
