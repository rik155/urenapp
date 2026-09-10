param(
    [string]$ServerUrl = "https://uren.93-119-6-183.sslip.io/api/excel",
    [string]$TargetDir = "$env:USERPROFILE\Documents\Urenregistratie",
    [int]$IntervalSeconds = 60
)

$ErrorActionPreference = "Stop"
$TargetFile = Join-Path $TargetDir "urenregistratie.xlsx"
$BackupDir = Join-Path $TargetDir "Backups"
$StatusFile = Join-Path $TargetDir "sync-status.txt"
$TempFile = Join-Path $TargetDir "urenregistratie.download.xlsx"

New-Item -ItemType Directory -Force -Path $TargetDir | Out-Null
New-Item -ItemType Directory -Force -Path $BackupDir | Out-Null

function Write-Status([string]$Message) {
    $stamp = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
    "$stamp - $Message" | Set-Content -Path $StatusFile -Encoding UTF8
    Write-Host "$stamp - $Message"
}

function Test-FileLocked([string]$Path) {
    if (-not (Test-Path $Path)) { return $false }
    try {
        $stream = [System.IO.File]::Open($Path, 'Open', 'ReadWrite', 'None')
        $stream.Close()
        return $false
    } catch {
        return $true
    }
}

function Sync-Workbook {
    try {
        Invoke-WebRequest -Uri $ServerUrl -OutFile $TempFile -UseBasicParsing -TimeoutSec 30

        if ((Get-Item $TempFile).Length -lt 1000) {
            throw "Gedownload bestand is onverwacht klein."
        }

        if (Test-FileLocked $TargetFile) {
            Remove-Item $TempFile -Force -ErrorAction SilentlyContinue
            Write-Status "Wachten: Excelbestand is geopend op kantoor."
            return
        }

        $replace = $true
        if (Test-Path $TargetFile) {
            $oldHash = (Get-FileHash $TargetFile -Algorithm SHA256).Hash
            $newHash = (Get-FileHash $TempFile -Algorithm SHA256).Hash
            if ($oldHash -eq $newHash) { $replace = $false }
        }

        if ($replace) {
            if (Test-Path $TargetFile) {
                $backupName = "urenregistratie_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
                Copy-Item $TargetFile (Join-Path $BackupDir $backupName) -Force
            }
            Move-Item $TempFile $TargetFile -Force
            Write-Status "Bijgewerkt: nieuwste urenbestand staat op kantoor."
        } else {
            Remove-Item $TempFile -Force -ErrorAction SilentlyContinue
            Write-Status "Geen wijzigingen."
        }

        Get-ChildItem $BackupDir -Filter "*.xlsx" | Sort-Object LastWriteTime -Descending | Select-Object -Skip 30 | Remove-Item -Force -ErrorAction SilentlyContinue
    } catch {
        Remove-Item $TempFile -Force -ErrorAction SilentlyContinue
        Write-Status "FOUT: $($_.Exception.Message)"
    }
}

Write-Status "UrenApp sync gestart. Controle elke $IntervalSeconds seconden."
while ($true) {
    Sync-Workbook
    Start-Sleep -Seconds $IntervalSeconds
}
