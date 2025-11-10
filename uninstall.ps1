Param(
    [string]$InstallDir = "C:\Program Files\NataliesDownloadApproval",
    [string]$ServiceName = "DownloadWatcherService"
)

# Require admin
If (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Error "Please run this script in an elevated PowerShell (Run as administrator)."; exit 1
}

Write-Host "Stopping and removing service $ServiceName" -ForegroundColor Cyan

# Stop service and wait for it to fully stop
try {
    $svc = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    if ($svc) {
        if ($svc.Status -eq 'Running') {
            Write-Host "Stopping service..." -ForegroundColor Yellow
            Stop-Service -Name $ServiceName -Force -ErrorAction SilentlyContinue
            sc.exe stop $ServiceName | Out-Null
            
            # Wait for service to stop
            $maxWait = 15
            $waited = 0
            while ($waited -lt $maxWait) {
                Start-Sleep -Seconds 1
                $waited++
                $svc.Refresh()
                if ($svc.Status -eq 'Stopped') {
                    Write-Host "Service stopped successfully." -ForegroundColor Green
                    break
                }
            }
            
            if ($svc.Status -ne 'Stopped') {
                Write-Warning "Service may still be running. Waiting additional 5 seconds..."
                Start-Sleep -Seconds 5
            }
        }
    }
} catch {
    Write-Host "Could not stop service (may not exist): $_" -ForegroundColor Yellow
}

# Delete service
Write-Host "Deleting service..." -ForegroundColor Yellow
sc.exe delete $ServiceName | Out-Null

# Wait a bit for Windows to release file handles
Write-Host "Waiting for file handles to be released..." -ForegroundColor Gray
Start-Sleep -Seconds 3

# Remove install directory with retry logic
if (Test-Path $InstallDir) {
    Write-Host "Removing install directory: $InstallDir" -ForegroundColor Yellow
    
    # Try to remove files with retries
    $maxRetries = 5
    $retryCount = 0
    $success = $false
    
    while ($retryCount -lt $maxRetries -and -not $success) {
        try {
            # Try to remove read-only attributes first
            Get-ChildItem -Path $InstallDir -Recurse -Force | ForEach-Object {
                if ($_.Attributes -band [System.IO.FileAttributes]::ReadOnly) {
                    $_.Attributes = $_.Attributes -bxor [System.IO.FileAttributes]::ReadOnly
                }
            }
            
            Remove-Item -Path $InstallDir -Recurse -Force -ErrorAction Stop
            $success = $true
            Write-Host "Install directory removed successfully." -ForegroundColor Green
        }
        catch {
            $retryCount++
            if ($retryCount -lt $maxRetries) {
                Write-Host "Attempt $retryCount failed. Waiting 2 seconds before retry..." -ForegroundColor Yellow
                Start-Sleep -Seconds 2
            }
            else {
                Write-Warning "Could not remove all files in install directory. Some files may still be locked."
                Write-Host "You may need to restart your computer to fully remove: $InstallDir" -ForegroundColor Yellow
            }
        }
    }
}

Write-Host "Uninstall complete." -ForegroundColor Green

