param(
  [string]$ExtensionId = "com.nester.ai.cepstarter"
)

$ErrorActionPreference = "Stop"

$sourcePath = Join-Path $PSScriptRoot "extension"
$targetRoot = Join-Path $env:APPDATA "Adobe\CEP\extensions"
$targetPath = Join-Path $targetRoot $ExtensionId

if (-not (Test-Path $sourcePath)) {
  throw "Source folder not found: $sourcePath"
}

if (Get-Process -Name "Illustrator" -ErrorAction SilentlyContinue) {
  throw "Illustrator is running. Close Illustrator first, then run install-dev.ps1 again."
}

New-Item -Path $targetRoot -ItemType Directory -Force | Out-Null

if (Test-Path $targetPath) {
  Get-ChildItem -Path $targetPath -Force | Remove-Item -Recurse -Force
} else {
  New-Item -Path $targetPath -ItemType Directory -Force | Out-Null
}

Copy-Item -Path (Join-Path $sourcePath "*") -Destination $targetPath -Recurse -Force

foreach ($version in 8..13) {
  $keyPath = "HKCU:\Software\Adobe\CSXS.$version"
  New-Item -Path $keyPath -Force | Out-Null
  New-ItemProperty -Path $keyPath -Name "PlayerDebugMode" -Value "1" -PropertyType String -Force | Out-Null
}

Write-Host "Installed/Updated CEP extension at: $targetPath"
Write-Host "PlayerDebugMode=1 set for CSXS.8 to CSXS.13"
Write-Host "Restart Illustrator, then open: Window > Extensions (Legacy) > Nester"
