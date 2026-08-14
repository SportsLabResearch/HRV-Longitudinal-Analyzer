Set-Location -Path $PSScriptRoot
Write-Host "============================================================"
Write-Host "HRV-LONGITUDINAL-ANALYZER - INICIO LOCAL"
Write-Host "============================================================"
py launcher_hrv.py
if ($LASTEXITCODE -ne 0) {
    python launcher_hrv.py
}
Read-Host "Pulsa ENTER para salir"
