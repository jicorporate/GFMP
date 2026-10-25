<#
.SYNOPSIS
    Lance le serveur de prévisualisation et de distribution mobile GFMP (SFP v1.0).
.DESCRIPTION
    Démarre un serveur local haute performance pour tester le portail mobile (PWA)
    et distribuer l'APK sur le réseau Wi-Fi local sans dépendance externe.
#>

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$MobileDir = Join-Path $ScriptDir "mobile-web"

Write-Host "==========================================================" -ForegroundColor Blue
Write-Host "   GFMP — Global Finance Management Platform (SFP v1.0)   " -ForegroundColor Yellow
Write-Host "==========================================================" -ForegroundColor Blue

# Trouver l'adresse IP locale Wi-Fi
$LocalIP = (Get-NetIPAddress -AddressFamily IPv4 -InterfaceAlias "Wi-Fi*", "Ethernet*" -ErrorAction SilentlyContinue | 
            Where-Object { $_.IPAddress -notlike "169.254*" -and $_.IPAddress -notlike "127.*" } | 
            Select-Object -First 1).IPAddress

if (-not $LocalIP) {
    $LocalIP = "localhost"
}

$Port = 8089

Write-Host ""
Write-Host "[+] Dossier servi : $MobileDir" -ForegroundColor Cyan
Write-Host "[+] Accès Local (PC)       : http://localhost:$Port" -ForegroundColor Green
Write-Host "[+] Accès Mobile (Wi-Fi)   : http://${LocalIP}:$Port" -ForegroundColor Green
Write-Host ""
Write-Host "Pour installer sur votre smartphone :" -ForegroundColor Yellow
Write-Host "  1. Connectez votre téléphone au même réseau Wi-Fi" -ForegroundColor White
Write-Host "  2. Ouvrez Chrome ou Samsung Internet à l'adresse : http://${LocalIP}:$Port" -ForegroundColor White
Write-Host "  3. Cliquez sur 'Ajouter à l'écran d'accueil' ou téléchargez l'APK" -ForegroundColor White
Write-Host ""
Write-Host "Appuyez sur Ctrl+C pour arrêter le serveur." -ForegroundColor Gray
Write-Host ""

Set-Location $MobileDir
python -m http.server $Port
