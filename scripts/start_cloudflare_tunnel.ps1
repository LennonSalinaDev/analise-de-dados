$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$Python = "python"
$Port = if ($env:PORT) { $env:PORT } else { "8000" }
$Url = "http://127.0.0.1:$Port"

Set-Location $ProjectRoot

if (-not (Get-Command cloudflared -ErrorAction SilentlyContinue)) {
    Write-Host "cloudflared não encontrado no PATH."
    Write-Host "Instale o Cloudflare Tunnel e rode este script novamente."
    Write-Host "Download: https://developers.cloudflare.com/cloudflare-one/connections/connect-networks/downloads/"
    exit 1
}

$existingApp = Get-CimInstance Win32_Process |
    Where-Object { $_.CommandLine -match "app\.py" -and $_.CommandLine -match [regex]::Escape($ProjectRoot) }

if (-not $existingApp) {
    Write-Host "Iniciando o Gerador de Orçamentos em $Url ..."
    Start-Process -FilePath $Python -ArgumentList "app.py" -WorkingDirectory $ProjectRoot -WindowStyle Hidden
    Start-Sleep -Seconds 2
}

Write-Host "Abrindo túnel Cloudflare para $Url ..."
Write-Host "Use a URL https://...trycloudflare.com que aparecer abaixo."
cloudflared tunnel --url $Url
