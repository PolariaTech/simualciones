$ErrorActionPreference = "Stop"

param(
  [string]$FrontUrl = "http://localhost:3001",
  [string]$Downloads = "$env:USERPROFILE\Downloads"
)

Write-Host "== Polaria UI Runner =="
Write-Host "Front: $FrontUrl"
Write-Host "Downloads: $Downloads"

if (!(Test-Path $Downloads)) {
  throw "No existe carpeta de descargas: $Downloads"
}

npm install
node .\run.mjs --downloads "$Downloads" --front-url "$FrontUrl" --headed
