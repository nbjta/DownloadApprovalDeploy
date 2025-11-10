Param(
    [string]$InstallDir = "C:\Program Files\NataliesDownloadApproval",
    [string]$ServiceName = "DownloadWatcherService",
    [string]$DownloadsPath = "$env:USERPROFILE\Downloads",
    [string]$AdminCode = "1234",
    [string]$QuarantineFolder = "",
    [switch]$KeepOnDeny = $true,
    [switch]$StartService
)

Write-Host "Installing to: $InstallDir" -ForegroundColor Cyan

# Require admin
If (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Error "Please run this script in an elevated PowerShell (Run as administrator)."; exit 1
}

$ScriptDir = $PSScriptRoot
$Root = Split-Path -Parent $PSScriptRoot
if (-not (Test-Path (Join-Path $Root "service") -PathType Container) -or -not (Test-Path (Join-Path $Root "prompt") -PathType Container)) {
    $Root = $ScriptDir
}
Set-Location $Root

# Prepare logs directory in project path
$logsDir = Join-Path $Root "logs"
New-Item -ItemType Directory -Force -Path $logsDir | Out-Null
try {
    $acl = Get-Acl $logsDir
    foreach ($identity in @("SYSTEM", "Administrators"))
    {
        $rule = New-Object System.Security.AccessControl.FileSystemAccessRule($identity, "FullControl", "ContainerInherit, ObjectInherit", "None", "Allow")
        $acl.SetAccessRule($rule)
    }
    Set-Acl -Path $logsDir -AclObject $acl
} catch {
    Write-Warning "Could not ensure permissions on logs directory ${logsDir}: $_"
}

# Determine source binaries
$sourceServiceProj = Join-Path $Root "DownloadWatcherService\DownloadWatcherService.csproj"
$sourcePromptProj = Join-Path $Root "DownloadApprovalPrompt\DownloadApprovalPrompt.csproj"
$usePublishedBinaries = (Test-Path $sourceServiceProj -PathType Leaf) -and (Test-Path $sourcePromptProj -PathType Leaf)

if ($usePublishedBinaries)
{
    Write-Host "Building service and prompt from source projects..." -ForegroundColor Cyan
    $publishService = Join-Path $env:TEMP "nda_service_$(Get-Date -Format yyyyMMddHHmmss)"
    $publishPrompt = Join-Path $env:TEMP "nda_prompt_$(Get-Date -Format yyyyMMddHHmmss)"
    dotnet publish $sourceServiceProj -c Release -r win-x64 -o $publishService | Out-Host
    dotnet publish $sourcePromptProj -c Release -r win-x64 -o $publishPrompt | Out-Host
}
else
{
    $publishService = Join-Path $Root "service"
    $publishPrompt = Join-Path $Root "prompt"
    if (-not (Test-Path $publishService -PathType Container) -or -not (Test-Path $publishPrompt -PathType Container))
    {
        Write-Error "Cannot locate pre-built binaries. Expected 'service' and 'prompt' folders alongside install.ps1"; exit 1
    }
    Write-Host "Using pre-built binaries from local 'service' and 'prompt' folders." -ForegroundColor Cyan
}

# Create install folders
$serviceDir = Join-Path $InstallDir "service"
$promptDir  = Join-Path $InstallDir "prompt"
New-Item -ItemType Directory -Force -Path $serviceDir | Out-Null
New-Item -ItemType Directory -Force -Path $promptDir | Out-Null

# Stop service if it's running before copying files
try {
    $svc = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    if ($svc) {
        if ($svc.Status -eq 'Running') {
            Write-Host "Stopping service $ServiceName..." -ForegroundColor Yellow
            # Try multiple methods to ensure service stops
            Stop-Service -Name $ServiceName -Force -ErrorAction SilentlyContinue
            sc.exe stop $ServiceName | Out-Null
            # Wait and verify service is stopped
            $maxWait = 10
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
                Write-Host "Warning: Service may still be running. Attempting to continue..." -ForegroundColor Yellow
            }
            # Additional wait to ensure files are released
            Start-Sleep -Seconds 2
        }
    }
} catch {
    Write-Host "Could not stop service (may not exist): $_" -ForegroundColor Yellow
}

# Copy outputs
try {
    Get-Process -Name "DownloadApprovalPrompt" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 1
} catch {
    Write-Host "Could not stop existing prompt processes (may not be running)." -ForegroundColor Yellow
}
Copy-Item -Recurse -Force "$publishService\*" $serviceDir
Copy-Item -Recurse -Force "$publishPrompt\*" $promptDir

# Create/overwrite appsettings.json
$appSettingsPath = Join-Path $serviceDir "appsettings.json"
if ([string]::IsNullOrWhiteSpace($DownloadsPath)) {
    Write-Host "DownloadsPath not specified - service will auto-detect per user" -ForegroundColor Yellow
    $DownloadsPath = ""
}
if ([string]::IsNullOrWhiteSpace($QuarantineFolder) -and -not [string]::IsNullOrWhiteSpace($DownloadsPath)) {
    $QuarantineFolder = Join-Path $DownloadsPath "Quarantine"
}
$keepOnDenyValue = $KeepOnDeny.IsPresent -or -not $PSBoundParameters.ContainsKey("KeepOnDeny")
$appSettings = @{
  DownloadWatcher = @{
    FolderPath = $DownloadsPath
    IncludeSubdirectories = $true
    Filter = "*.*"
    DebounceMs = 2000
    DedupSeconds = 30
    IgnoredExtensions = @()
    ExcludePaths = @()
    WhitelistedProcesses = @(
      "Cursor",
      "Code",
      "devenv",
      "MSBuild",
      "mspaint",
      "WINWORD",
      "EXCEL",
      "POWERPNT",
      "OUTLOOK",
      "notepad",
      "notepad++",
      "git",
      "dotnet",
      "msbuild",
      "node",
      "npm",
      "python",
      "javac",
      "java"
    )
    QuarantineFolder = $QuarantineFolder
    KeepOnDeny = $keepOnDenyValue
  }
  Prompt = @{
    ExecutablePath = (Join-Path $promptDir "DownloadApprovalPrompt.exe")
    Arguments = "--admin-code PLACEHOLDER"
  }
  Logs = @{
    Directory = $logsDir
    WatcherLog = "watcher.log"
    PromptLog = "prompt.log"
  }
  Logging = @{ LogLevel = @{ Default = "Information"; "Microsoft.Hosting.Lifetime" = "Information" } }
} | ConvertTo-Json -Depth 5
Set-Content -LiteralPath $appSettingsPath -Value $appSettings -Encoding UTF8

# Write shared logs configuration for prompt/service
$logsConfig = @{
  Directory = $logsDir
  WatcherLog = "watcher.log"
  PromptLog = "prompt.log"
} | ConvertTo-Json -Depth 5
Set-Content -LiteralPath (Join-Path $InstallDir "logs.config") -Value $logsConfig -Encoding UTF8

# Create admin code file (hidden, not in git)
$codeFilePath = Join-Path $serviceDir "system.cfg"
Set-Content -LiteralPath $codeFilePath -Value $AdminCode -Encoding UTF8 -Force
try {
    $file = Get-Item $codeFilePath -Force
    $file.Attributes = $file.Attributes -bor [System.IO.FileAttributes]::Hidden
} catch {
    # Ignore if we can't set hidden attribute
}

Write-Host "Configuration:" -ForegroundColor Cyan
Write-Host "  Downloads folder: $DownloadsPath" -ForegroundColor Gray
Write-Host "  Quarantine folder: $QuarantineFolder" -ForegroundColor Gray
Write-Host "  Admin code: $AdminCode (stored in system.cfg)" -ForegroundColor Gray
Write-Host "  Keep files on deny: $keepOnDenyValue" -ForegroundColor Gray
Write-Host "  Whitelisted processes: Cursor, Code, git, dotnet, node, etc." -ForegroundColor Gray
Write-Host "  Whitelisted paths: .cursor folders, Cursor AppData folders" -ForegroundColor Gray

# Install/Update Windows Service
$svcExe = Join-Path $serviceDir "DownloadWatcherService.exe"
if (-not (Test-Path $svcExe)) { Write-Error "Service executable not found: $svcExe"; exit 1 }

# Check if service exists
sc.exe query $ServiceName | Out-Null
if ($LASTEXITCODE -eq 0) {
  Write-Host "Service exists. Updating binary path and resetting to auto." -ForegroundColor Yellow
  # Service should already be stopped from above, but ensure it's stopped
  $svc = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
  if ($svc -and $svc.Status -eq 'Running') {
    Write-Host "Service still running, stopping before update..." -ForegroundColor Yellow
    Stop-Service -Name $ServiceName -Force -ErrorAction SilentlyContinue
    sc.exe stop $ServiceName | Out-Null
    Start-Sleep -Seconds 5
    $svc.Refresh()
    if ($svc.Status -ne 'Stopped') {
      Write-Warning "Service may not have stopped completely. Continuing anyway..."
    }
  }
  # sc.exe config requires space after = and separate arguments
  $configResult = sc.exe config $ServiceName binPath= "`"$svcExe`"" start= auto 2>&1
  if ($LASTEXITCODE -ne 0) {
    Write-Error "Failed to update service configuration: $configResult"; exit 1
  }
  # Always update description and display name with version
  $version = "1.1.12"
  $description = "Watcher that opens approval prompt for new downloads (v$version)"
  $displayName = "Download Watcher Service (v$version)"
  sc.exe description $ServiceName $description | Out-Null
  Set-Service -Name $ServiceName -DisplayName $displayName -ErrorAction SilentlyContinue
  Write-Host "Service description updated: $description" -ForegroundColor Gray
  Write-Host "Service display name updated: $displayName" -ForegroundColor Gray
} else {
  Write-Host "Creating service $ServiceName" -ForegroundColor Cyan
  # sc.exe create requires space after = and separate arguments
  $createResult = sc.exe create $ServiceName binPath= "`"$svcExe`"" start= auto 2>&1
  if ($LASTEXITCODE -ne 0) {
    Write-Error "Failed to create service: $createResult"; exit 1
  }
  $version = "1.1.12"
  $description = "Watcher that opens approval prompt for new downloads (v$version)"
  $displayName = "Download Watcher Service (v$version)"
  sc.exe description $ServiceName $description | Out-Null
  Set-Service -Name $ServiceName -DisplayName $displayName -ErrorAction SilentlyContinue
  Write-Host "Service created successfully." -ForegroundColor Green
  Write-Host "Service version: $version" -ForegroundColor Gray
  Write-Host "Service display name: $displayName" -ForegroundColor Gray
}

if ($StartService) {
  Start-Sleep -Seconds 1
  $startResult = sc.exe start $ServiceName 2>&1
  if ($LASTEXITCODE -eq 0) {
    Write-Host "Service started successfully." -ForegroundColor Green
  } else {
    Write-Warning "Service may have failed to start: $startResult"
  }
} else {
  Write-Host "Service installed. Start with: sc.exe start $ServiceName" -ForegroundColor Green
}

Write-Host ""
Write-Host "Installation complete." -ForegroundColor Green
Write-Host ""
Write-Host "Usage examples:" -ForegroundColor Cyan
Write-Host "  .\install.ps1 -AdminCode 'MySecret123' -StartService" -ForegroundColor Gray
Write-Host "  .\install.ps1 -DownloadsPath 'D:\Downloads' -QuarantineFolder 'D:\Quarantine' -StartService" -ForegroundColor Gray
Write-Host "  .\install.ps1 -AdminCode '1234' -KeepOnDeny:`$false -StartService" -ForegroundColor Gray

