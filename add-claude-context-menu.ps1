#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Adds a "Claude Code" dropdown to the Windows 11 modern (top-level) context menu.
.DESCRIPTION
    Compiles a native COM DLL, creates a signed sparse AppX package, and registers
    it so the dropdown appears in the top-level right-click menu for directory
    backgrounds.

    Submenu items:
      - Open Claude Code Here        (launches claude in the current folder)
      - YOLO Claude, YOLO!           (launches with --dangerously-skip-permissions)

    Run with -Uninstall to cleanly remove everything.
.PARAMETER ClaudePath
    Optional. Full path to claude.exe. Auto-detected if omitted.
#>
param(
    [switch]$Uninstall,
    [string]$ClaudePath
)

$ErrorActionPreference = "Stop"

# --- Configuration ---
$clsid         = "E3C26D71-5A2F-4B89-9C7E-A1D3F6B84E52"
$packageName   = "ClaudeCode.ContextMenu"
$publisher     = "CN=ClaudeCodeDev"
$installDir    = Join-Path $env:ProgramFiles "ClaudeCodeContextMenu"
$certFriendly  = "ClaudeCode Context Menu Dev Cert"
$shellKey      = "Registry::HKEY_CLASSES_ROOT\Directory\Background\shell\ClaudeCodeCLI"

# Auto-detect Windows SDK
function Find-SdkBin {
    $kitsRoot = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows Kits\Installed Roots" -ErrorAction SilentlyContinue).KitsRoot10
    if (-not $kitsRoot) { Write-Error "Windows SDK not found. Install the Windows 10/11 SDK." }
    $ver = Get-ChildItem (Join-Path $kitsRoot "bin") -Directory |
        Where-Object { $_.Name -match '^\d+\.' } |
        Sort-Object Name -Descending | Select-Object -First 1
    if (-not $ver) { Write-Error "No SDK version found under $kitsRoot\bin" }
    return Join-Path $kitsRoot "bin\$($ver.Name)\x64"
}

# ── Uninstall ────────────────────────────────────────────────────────────────
if ($Uninstall) {
    Write-Host "Uninstalling..." -ForegroundColor Cyan

    $pkg = Get-AppxPackage -Name $packageName -ErrorAction SilentlyContinue
    if ($pkg) { Remove-AppxPackage $pkg; Write-Host "  Removed AppX package." -ForegroundColor Green }

    if (Test-Path $shellKey) {
        Remove-Item $shellKey -Recurse -Force
        Write-Host "  Removed registry shell key." -ForegroundColor Green
    }

    $clsidKey = "Registry::HKEY_CLASSES_ROOT\CLSID\{$clsid}"
    if (Test-Path $clsidKey) {
        Remove-Item $clsidKey -Recurse -Force
        Write-Host "  Removed COM registration." -ForegroundColor Green
    }

    Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2

    if (Test-Path $installDir) {
        Remove-Item $installDir -Recurse -Force
        Write-Host "  Deleted $installDir." -ForegroundColor Green
    }

    foreach ($store in @("Cert:\CurrentUser\My", "Cert:\LocalMachine\TrustedPeople")) {
        Get-ChildItem $store | Where-Object { $_.FriendlyName -eq $certFriendly } |
            ForEach-Object { Remove-Item $_.PSPath -Force; Write-Host "  Removed cert from $store." -ForegroundColor Green }
    }

    Start-Process explorer.exe
    Write-Host "Done." -ForegroundColor Cyan
    return
}

# ── Install ──────────────────────────────────────────────────────────────────

$sdkBin   = Find-SdkBin
$makeAppx = Join-Path $sdkBin "makeappx.exe"
$signTool = Join-Path $sdkBin "signtool.exe"
foreach ($tool in @($makeAppx, $signTool)) {
    if (-not (Test-Path $tool)) { Write-Error "Missing: $tool" }
}

# 1. Locate claude.exe
if (-not $ClaudePath) {
    $ClaudePath = (Get-Command claude.exe -ErrorAction SilentlyContinue).Source
}
if (-not $ClaudePath) {
    $candidate = Join-Path $env:USERPROFILE ".local\bin\claude.exe"
    if (Test-Path $candidate) { $ClaudePath = $candidate }
}
if (-not $ClaudePath) {
    Write-Error "claude.exe not found. Install Claude Code first, or pass -ClaudePath."
}
$ClaudePath = [System.IO.Path]::GetFullPath($ClaudePath)
Write-Host "Using claude.exe at: $ClaudePath" -ForegroundColor Cyan

# 2. Generate claude_path.h with C-escaped backslashes
$escaped = [string]::Join("\\", $ClaudePath.Split([char]0x5C))
$headerPath = Join-Path $PSScriptRoot "claude_path.h"
$line = '#define CLAUDE_EXE_PATH L"' + $escaped + '"'
[IO.File]::WriteAllText($headerPath, $line)

# 3. Build native COM DLL
$buildBat = Join-Path $PSScriptRoot "build.bat"
$srcDll   = Join-Path $PSScriptRoot "ClaudeCodeContextMenu.dll"

Write-Host "Building native COM DLL..." -ForegroundColor Cyan
$prevEAP = $ErrorActionPreference
$ErrorActionPreference = "Continue"
$buildOut = cmd /c $buildBat 2>&1
$ErrorActionPreference = $prevEAP
if (-not (Test-Path $srcDll)) {
    $buildOut | Write-Host
    Write-Error "Build failed. Ensure Visual Studio C++ desktop workload is installed."
}
Write-Host "  Build succeeded." -ForegroundColor Green

# 4. Stop Explorer if DLL is locked
if (Test-Path $installDir) {
    $dllDest = Join-Path $installDir "ClaudeCodeContextMenu.dll"
    if (Test-Path $dllDest) {
        try { [IO.File]::OpenWrite($dllDest).Close() }
        catch {
            Write-Host "  DLL locked by Explorer. Restarting Explorer..." -ForegroundColor Yellow
            Stop-Process -Name explorer -Force
            Start-Sleep -Seconds 2
        }
    }
}

# 5. Copy DLL to install directory
if (-not (Test-Path $installDir)) { New-Item -ItemType Directory -Path $installDir -Force | Out-Null }
Copy-Item $srcDll $installDir -Force
Write-Host "  Installed to $installDir" -ForegroundColor Green

# 6. Placeholder logo PNG
$logoDest = Join-Path $installDir "logo.png"
if (-not (Test-Path $logoDest)) {
    Add-Type -AssemblyName System.Drawing
    $bmp = New-Object System.Drawing.Bitmap(44, 44)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.Clear([System.Drawing.Color]::Transparent)
    $g.Dispose()
    $bmp.Save($logoDest, [System.Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose()
}

# 7. Stub exe (AppX manifest requires an .exe)
$stubPath = Join-Path $installDir "Stub.exe"
if (-not (Test-Path $stubPath)) {
    $fwDir = [System.Runtime.InteropServices.RuntimeEnvironment]::GetRuntimeDirectory()
    $csc = Join-Path $fwDir "csc.exe"
    $stubSrc = Join-Path $env:TEMP "stub.cs"
    Set-Content -Path $stubSrc -Value "class S{static void Main(){}}"
    & $csc /nologo /target:exe /out:$stubPath $stubSrc 2>&1 | Out-Null
}

# 8. Register COM class
Write-Host "Registering COM class..." -ForegroundColor Cyan
$clsidKey = "Registry::HKEY_CLASSES_ROOT\CLSID\{$clsid}"
New-Item -Path "$clsidKey\InprocServer32" -Force | Out-Null
Set-ItemProperty -Path $clsidKey -Name "(Default)" -Value "ClaudeCodeCommand"
Set-ItemProperty -Path "$clsidKey\InprocServer32" -Name "(Default)" -Value (Join-Path $installDir "ClaudeCodeContextMenu.dll")
Set-ItemProperty -Path "$clsidKey\InprocServer32" -Name "ThreadingModel" -Value "Both"
Write-Host "  Registered CLSID {$clsid}" -ForegroundColor Green

# 9. Shell verb with ExplorerCommandHandler
New-Item -Path $shellKey -Force | Out-Null
Set-ItemProperty -Path $shellKey -Name "ExplorerCommandHandler" -Value "{$clsid}"
Write-Host "  Created shell verb." -ForegroundColor Green

# 10. AppxManifest.xml
$manifestPath = Join-Path $installDir "AppxManifest.xml"
@"
<?xml version="1.0" encoding="utf-8"?>
<Package
  xmlns="http://schemas.microsoft.com/appx/manifest/foundation/windows10"
  xmlns:uap="http://schemas.microsoft.com/appx/manifest/uap/windows10"
  xmlns:uap10="http://schemas.microsoft.com/appx/manifest/uap/windows10/10"
  xmlns:com="http://schemas.microsoft.com/appx/manifest/com/windows10"
  xmlns:desktop4="http://schemas.microsoft.com/appx/manifest/desktop/windows10/4"
  xmlns:desktop5="http://schemas.microsoft.com/appx/manifest/desktop/windows10/5"
  xmlns:desktop6="http://schemas.microsoft.com/appx/manifest/desktop/windows10/6"
  xmlns:rescap="http://schemas.microsoft.com/appx/manifest/foundation/windows10/restrictedcapabilities"
  IgnorableNamespaces="uap uap10 com desktop4 desktop5 desktop6 rescap">

  <Identity Name="$packageName"
            Publisher="$publisher"
            Version="1.0.0.0"
            ProcessorArchitecture="x64" />

  <Properties>
    <DisplayName>ClaudeCode Context Menu</DisplayName>
    <PublisherDisplayName>ClaudeCode</PublisherDisplayName>
    <Logo>logo.png</Logo>
    <uap10:AllowExternalContent>true</uap10:AllowExternalContent>
    <desktop6:RegistryWriteVirtualization>disabled</desktop6:RegistryWriteVirtualization>
    <desktop6:FileSystemWriteVirtualization>disabled</desktop6:FileSystemWriteVirtualization>
  </Properties>

  <Dependencies>
    <TargetDeviceFamily Name="Windows.Desktop"
                        MinVersion="10.0.19041.0"
                        MaxVersionTested="10.0.26100.0" />
  </Dependencies>

  <Resources>
    <Resource Language="en-us" />
  </Resources>

  <Applications>
    <Application Id="App"
                 Executable="Stub.exe"
                 uap10:TrustLevel="mediumIL"
                 uap10:RuntimeBehavior="win32App">
      <uap:VisualElements
        DisplayName="ClaudeCode"
        Description="ClaudeCode Context Menu"
        BackgroundColor="transparent"
        Square150x150Logo="logo.png"
        Square44x44Logo="logo.png"
        AppListEntry="none" />
      <Extensions>
        <desktop4:Extension Category="windows.fileExplorerContextMenus">
          <desktop4:FileExplorerContextMenus>
            <desktop5:ItemType Type="Directory\Background">
              <desktop5:Verb Id="ClaudeCodeCLI" Clsid="$clsid" />
            </desktop5:ItemType>
          </desktop4:FileExplorerContextMenus>
        </desktop4:Extension>
        <com:Extension Category="windows.comServer">
          <com:ComServer>
            <com:SurrogateServer DisplayName="ClaudeCode Context Menu Handler">
              <com:Class Id="$clsid"
                         Path="ClaudeCodeContextMenu.dll"
                         ThreadingModel="Both" />
            </com:SurrogateServer>
          </com:ComServer>
        </com:Extension>
      </Extensions>
    </Application>
  </Applications>

  <Capabilities>
    <rescap:Capability Name="runFullTrust" />
    <rescap:Capability Name="unvirtualizedResources" />
  </Capabilities>
</Package>
"@ | Set-Content -Path $manifestPath -Encoding UTF8
Write-Host "  Generated AppxManifest.xml" -ForegroundColor Green

# 11. Self-signed certificate
Write-Host "Setting up signing certificate..." -ForegroundColor Cyan
$cert = Get-ChildItem Cert:\CurrentUser\My |
    Where-Object { $_.Subject -eq $publisher -and $_.FriendlyName -eq $certFriendly } |
    Select-Object -First 1

if (-not $cert) {
    $cert = New-SelfSignedCertificate `
        -Type Custom -Subject $publisher `
        -KeyUsage DigitalSignature `
        -FriendlyName $certFriendly `
        -CertStoreLocation "Cert:\CurrentUser\My" `
        -TextExtension @("2.5.29.37={text}1.3.6.1.5.5.7.3.3", "2.5.29.19={text}")
    Write-Host "  Created signing certificate." -ForegroundColor Green
} else {
    Write-Host "  Reusing existing certificate." -ForegroundColor Green
}

$trusted = Get-ChildItem Cert:\LocalMachine\TrustedPeople | Where-Object { $_.Thumbprint -eq $cert.Thumbprint }
if (-not $trusted) {
    $tmpCer = Join-Path $env:TEMP "claudecode_dev.cer"
    Export-Certificate -Cert $cert -FilePath $tmpCer | Out-Null
    Import-Certificate -FilePath $tmpCer -CertStoreLocation "Cert:\LocalMachine\TrustedPeople" | Out-Null
    Remove-Item $tmpCer -Force
    Write-Host "  Certificate trusted." -ForegroundColor Green
}

# 12. Remove previous package
$existingPkg = Get-AppxPackage -Name $packageName -ErrorAction SilentlyContinue
if ($existingPkg) {
    Remove-AppxPackage $existingPkg
    Write-Host "  Removed previous package." -ForegroundColor Green
}

# 13. Create and sign MSIX
Write-Host "Creating MSIX package..." -ForegroundColor Cyan
$msixPath = Join-Path $env:TEMP "ClaudeCodeContextMenu.msix"
if (Test-Path $msixPath) { Remove-Item $msixPath -Force }

$packArgs = @("pack", "/d", "`"$installDir`"", "/p", "`"$msixPath`"", "/nv", "/o")
$packResult = & $makeAppx @packArgs 2>&1
if ($LASTEXITCODE -ne 0) { $packResult | Write-Host; Write-Error "MakeAppx failed." }
Write-Host "  Package created." -ForegroundColor Green

$signResult = & $signTool sign /fd SHA256 /a /sha1 $cert.Thumbprint "`"$msixPath`"" 2>&1
if ($LASTEXITCODE -ne 0) { $signResult | Write-Host; Write-Error "SignTool failed." }
Write-Host "  Package signed." -ForegroundColor Green

# 14. Register sparse package
Write-Host "Registering sparse AppX package..." -ForegroundColor Cyan
Add-AppxPackage -Path $msixPath -ExternalLocation $installDir
Write-Host "  Package registered." -ForegroundColor Green

# 15. Restart Explorer
Write-Host ""
Write-Host "All done! Restarting Explorer..." -ForegroundColor Cyan
Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1
Start-Process explorer.exe
Write-Host "Right-click inside any folder to see the 'Claude Code' dropdown." -ForegroundColor Green
