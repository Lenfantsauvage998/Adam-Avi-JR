# Adam Bot — NSSM Service Installer
# Run this script as Administrator

$nssm    = "C:\nssm\nssm-2.24\win64\nssm.exe"
$python  = "C:\Python313\python.exe"
$script  = "C:\Users\dani1\pc-agent\main.py"
$workdir = "C:\Users\dani1\pc-agent"
$service = "AdamBot"
$logOut  = "C:\Users\dani1\pc-agent\bot.log"
$logErr  = "C:\Users\dani1\pc-agent\bot_err.log"

Write-Host "Installing AdamBot as Windows service..." -ForegroundColor Cyan

# Remove existing service if present
$existing = Get-Service -Name $service -ErrorAction SilentlyContinue
if ($existing) {
    Write-Host "Stopping existing service..."
    & $nssm stop $service confirm
    & $nssm remove $service confirm
    Start-Sleep -Seconds 2
}

# Install service
& $nssm install $service $python $script

# Set working directory (so relative paths in main.py resolve correctly)
& $nssm set $service AppDirectory $workdir

# Redirect stdout/stderr to log files
& $nssm set $service AppStdout $logOut
& $nssm set $service AppStderr $logErr

# Append to logs instead of overwriting on each restart
& $nssm set $service AppStdoutCreationDisposition 4
& $nssm set $service AppStderrCreationDisposition 4

# Restart policy: restart after 5s on crash, then 10s, then 30s
& $nssm set $service AppRestartDelay 5000

# Start service automatically at boot
& $nssm set $service Start SERVICE_AUTO_START

# Set display name and description
& $nssm set $service DisplayName "Adam AI Bot"
& $nssm set $service Description "Adam personal AI Telegram bot — runs continuously in background"

# Start it now
Write-Host "Starting AdamBot service..." -ForegroundColor Cyan
& $nssm start $service

Start-Sleep -Seconds 3
$svc = Get-Service -Name $service -ErrorAction SilentlyContinue
if ($svc -and $svc.Status -eq "Running") {
    Write-Host "AdamBot is RUNNING as a Windows service." -ForegroundColor Green
    Write-Host "Logs: $logOut"
    Write-Host "Errors: $logErr"
    Write-Host ""
    Write-Host "Useful commands:"
    Write-Host "  Stop:    & '$nssm' stop AdamBot"
    Write-Host "  Start:   & '$nssm' start AdamBot"
    Write-Host "  Restart: & '$nssm' restart AdamBot"
    Write-Host "  Remove:  & '$nssm' remove AdamBot confirm"
    Write-Host "  Status:  Get-Service AdamBot"
} else {
    Write-Host "Service may not have started. Check logs:" -ForegroundColor Yellow
    Write-Host "  $logErr"
}
