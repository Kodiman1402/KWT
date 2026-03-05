# Kodimans Windows Tool v4.0.0

$ProgressPreference    = 'SilentlyContinue'
$ErrorActionPreference = 'SilentlyContinue'

# ================================================================
# KONFIGURATION
# ================================================================
$VT_API_KEY  = ""        # Kostenlos: https://www.virustotal.com/gui/join-us
$CURRENT_VER = "4.0.0"
$UPDATE_URL  = ""        # Optional: URL zu version.txt auf GitHub
# ================================================================

trap {
    $msg = "CRASH $(Get-Date): $($_.Exception.Message) | Zeile $($_.InvocationInfo.ScriptLineNumber)"
    Out-File -FilePath "$env:USERPROFILE\Desktop\KWT_CRASH.log" -Encoding UTF8 -InputObject $msg -Append
    Write-Host $msg -ForegroundColor Red
    Read-Host "Fehler - Enter zum Beenden"
    exit 1
}

$ts          = Get-Date -Format 'yyyy-MM-dd_HH-mm'
$LogFile     = "$env:USERPROFILE\Desktop\KWT_$ts.log"
$HtmlFile    = "$env:USERPROFILE\Desktop\KWT_$ts.html"
$HtmlEntries = [System.Collections.Generic.List[string]]::new()
$Sep         = "=" * 70

function Clean-Text {
    param([string]$Text)
    if ($null -eq $Text) { return "" }
    $t = $Text
    $t = $t -replace [char]0xC4, "Ae"
    $t = $t -replace [char]0xD6, "Oe"
    $t = $t -replace [char]0xDC, "Ue"
    $t = $t -replace [char]0xE4, "ae"
    $t = $t -replace [char]0xF6, "oe"
    $t = $t -replace [char]0xFC, "ue"
    $t = $t -replace [char]0xDF, "ss"
    $t = $t -replace '[^\x20-\x7E]', ''
    return $t
}

function Write-Log {
    param([string]$Text, [string]$Color = "Green")
    $clean = Clean-Text $Text
    $Line  = "[$(Get-Date -Format 'HH:mm:ss')] $clean"
    Write-Host $Line -ForegroundColor $Color
    Out-File -FilePath $LogFile -Append -Encoding UTF8 -InputObject $Line
    $hc = switch ($Color) {
        "Red"     { "#ff4444" }
        "Yellow"  { "#ffaa00" }
        "Cyan"    { "#00ccff" }
        "Gray"    { "#888888" }
        "Magenta" { "#ff66ff" }
        default   { "#00ff41" }
    }
    $esc = $clean -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;'
    $script:HtmlEntries.Add("<p style='color:$hc;margin:1px 0;'>$esc</p>")
}

function Write-Section {
    param([string]$Title)
    Write-Host ""
    Write-Host $Sep       -ForegroundColor Green
    Write-Host "  $Title" -ForegroundColor Green
    Write-Host $Sep       -ForegroundColor Green
    Out-File -FilePath $LogFile -Append -Encoding UTF8 -InputObject "`n$Sep`n  $Title`n$Sep"
    $script:HtmlEntries.Add("<h2 style='color:#00ff41;border-left:4px solid #00ff41;padding-left:12px;margin-top:30px;font-size:14px;'>$Title</h2>")
}

function Filter-Output {
    param([string[]]$Lines)
    $result = @()
    foreach ($line in $Lines) {
        if ($null -eq $line)           { continue }
        $t = (Clean-Text $line).Trim()
        if ($t -eq "")                 { continue }
        if ($t.Length -gt 8) {
            $spaces = ($t -replace '[^ ]', '').Length
            if (($spaces / $t.Length) -gt 0.35) { continue }
        }
        if ($t -match '\d\s{0,3}%')    { continue }
        if ($t -match '^[\s=\-\*_]+$') { continue }
        $result += $t
    }
    return $result
}

function Get-VTReport {
    param([string]$FilePath)
    if ([string]::IsNullOrEmpty($VT_API_KEY)) { return $null }
    try {
        $hash     = (Get-FileHash -Path $FilePath -Algorithm SHA256 -ErrorAction Stop).Hash
        $uri      = "https://www.virustotal.com/api/v3/files/$hash"
        $headers  = @{ "x-apikey" = $VT_API_KEY }
        $response = Invoke-RestMethod -Uri $uri -Headers $headers -Method GET -TimeoutSec 10 -ErrorAction Stop
        if ($response.data.attributes.last_analysis_stats) {
            $mal   = $response.data.attributes.last_analysis_stats.malicious
            $total = 0
            $response.data.attributes.last_analysis_stats.PSObject.Properties | ForEach-Object { $total += $_.Value }
            return "$mal/$total Engines"
        }
    } catch { return $null }
    return $null
}

# ================================================================
# ASCII BANNER
# ================================================================
Clear-Host
Write-Host ""
Write-Host "  _  __         _ _                           " -ForegroundColor Green
Write-Host " | |/ /___   __| (_)_ __ ___   __ _ _ __  ___" -ForegroundColor Green
Write-Host " | ' // _ \ / _`| | '_ \` _ \ / _`| '_ \/ __|" -ForegroundColor Green
Write-Host " | . \ (_) | (_| | | | | | | | (_| | | | \__ \" -ForegroundColor Green
Write-Host " |_|\_\___/ \__,_|_|_| |_| |_|\__,_|_| |_|___/" -ForegroundColor Green
Write-Host ""
Write-Host " __        _____ _   _                        " -ForegroundColor Green
Write-Host " \ \      / /_ _| \ | |                       " -ForegroundColor Green
Write-Host "  \ \ /\ / / | ||  \| |                       " -ForegroundColor Green
Write-Host "   \ V  V /  | || |\  |                       " -ForegroundColor Green
Write-Host "    \_/\_/  |___|_| \_|                       " -ForegroundColor Green
Write-Host ""
Write-Host "  ____  _                                     " -ForegroundColor Green
Write-Host " |  _ \(_) __ _  __ _ _ __   ___  ___  ___   " -ForegroundColor Green
Write-Host " | | | | |/ _`| / _`| '_ \ / _ \/ __|/ _ \  " -ForegroundColor Green
Write-Host " | |_| | | (_| | (_| | | | | (_) \__ \  __/  " -ForegroundColor Green
Write-Host " |____/|_|\__,_|\__, |_| |_|\___/|___/\___|  " -ForegroundColor Green
Write-Host "                |___/                         " -ForegroundColor Green
Write-Host ""
Write-Host $Sep -ForegroundColor Green
Write-Host "   v4.0.0  |  by Kodiman_Himself  |  Bremen   " -ForegroundColor Green
Write-Host $Sep -ForegroundColor Green
Write-Host "   LOG:  $LogFile"  -ForegroundColor Green
Write-Host "   HTML: $HtmlFile" -ForegroundColor Green
Write-Host $Sep -ForegroundColor Green

Out-File -FilePath $LogFile -Encoding UTF8 -InputObject "Kodimans Windows Tool v4.0.0 - $(Get-Date)`nSystem: $env:COMPUTERNAME`n"
$script:HtmlEntries.Add("<h1 style='color:#00ff41;text-align:center;border-bottom:2px solid #00ff41;padding-bottom:15px;letter-spacing:3px;font-family:monospace;'>KODIMANS WINDOWS TOOL v4.0.0</h1>")
$script:HtmlEntries.Add("<p style='color:#555;text-align:center;margin-bottom:30px;font-size:12px;'>System: $env:COMPUTERNAME &nbsp;&bull;&nbsp; $(Get-Date) &nbsp;&bull;&nbsp; v4.0.0 by Kodiman_Himself &bull; Bremen</p>")

# ================================================================
# AUTO-UPDATE CHECK
# ================================================================
if (-not [string]::IsNullOrEmpty($UPDATE_URL)) {
    Write-Log "Pruefe auf Updates - bitte warten..." "Red"
    try {
        $latestVer = (Invoke-WebRequest -Uri $UPDATE_URL -UseBasicParsing -TimeoutSec 4 -ErrorAction Stop).Content.Trim()
        if ($latestVer -and $latestVer -ne $CURRENT_VER) {
            Write-Log "  [UPDATE] Neue Version v$latestVer verfuegbar! (Aktuell: v$CURRENT_VER)" "Yellow"
        } else {
            Write-Log "  [OK] Version v$CURRENT_VER ist aktuell." "Green"
        }
    } catch {
        Write-Log "  [INFO] Update-Check fehlgeschlagen: $($_.Exception.Message)" "Yellow"
    }
}

# ================================================================
# 1/12 - SFC
# ================================================================
Write-Section "1/12 - System File Checker (SFC /scannow)"
Write-Log "SFC-Scan laeuft - bitte warten..." "Red"
try {
    $sfcRaw    = & sfc /scannow 2>&1 | ForEach-Object { "$_" }
    $sfcClean  = $sfcRaw | ForEach-Object { Clean-Text "$_" }
    foreach ($line in (Filter-Output $sfcRaw)) { Write-Log $line "Green" }
    $sfcJoined = $sfcClean -join " "
    if     ($sfcJoined -match "no integrity violations")  { Write-Log "[OK]     SFC: Keine beschaedigten Systemdateien." "Green"      }
    elseif ($sfcJoined -match "hat keine Integrit")        { Write-Log "[OK]     SFC: Keine beschaedigten Systemdateien." "Green"      }
    elseif ($sfcJoined -match "Integritatsverletzungen")   { Write-Log "[OK]     SFC: Keine beschaedigten Systemdateien." "Green"      }
    elseif ($sfcJoined -match "found and repaired")        { Write-Log "[WARN]   SFC: Fehler repariert - Neustart empfohlen!" "Yellow" }
    elseif ($sfcJoined -match "erfolgreich repariert")     { Write-Log "[WARN]   SFC: Fehler repariert - Neustart empfohlen!" "Yellow" }
    elseif ($sfcJoined -match "found corrupt files")       { Write-Log "[FEHLER] SFC: Beschaedigte Dateien - DISM noetig!" "Red"      }
    elseif ($sfcJoined -match "nicht repariert")           { Write-Log "[FEHLER] SFC: Beschaedigte Dateien - DISM noetig!" "Red"      }
    else {
        $last = ($sfcClean | Where-Object { $_.Trim().Length -gt 10 } | Select-Object -Last 1)
        Write-Log "[INFO]   SFC: $last" "Yellow"
    }
} catch {
    Write-Log "[FEHLER] SFC: $($_.Exception.Message)" "Red"
}

# ================================================================
# 2/12 - DISM
# ================================================================
Write-Section "2/12 - DISM Pruefung"
Write-Log "DISM ScanHealth laeuft - bitte warten..." "Red"
try {
    $dismScanRaw    = & DISM /Online /Cleanup-Image /ScanHealth 2>&1 | ForEach-Object { "$_" }
    $dismScanClean  = $dismScanRaw | ForEach-Object { Clean-Text "$_" }
    foreach ($line in (Filter-Output $dismScanRaw)) { Write-Log $line "Green" }

    Write-Log "DISM CheckHealth laeuft - bitte warten..." "Red"
    $dismCheckRaw   = & DISM /Online /Cleanup-Image /CheckHealth 2>&1 | ForEach-Object { "$_" }
    $dismCheckClean = $dismCheckRaw | ForEach-Object { Clean-Text "$_" }
    foreach ($line in (Filter-Output $dismCheckRaw)) { Write-Log $line "Green" }

    $scanJoined  = $dismScanClean  -join " "
    $checkJoined = $dismCheckClean -join " "

    $isRepairable = ($checkJoined -match "repairable")                -or
                    ($scanJoined  -match "repairable")                -or
                    ($checkJoined -match "kann repariert werden")     -or
                    ($scanJoined  -match "kann repariert werden")

    $isClean      = ($checkJoined -match "No component store")        -or
                    ($checkJoined -match "Komponentenspeicher")       -or
                    ($scanJoined  -match "Komponentenspeicher")       -or
                    ($checkJoined -match "keine.*Beschaed")           -or
                    ($checkJoined -match "erfolgreich abgeschlossen") -or
                    ($scanJoined  -match "erfolgreich abgeschlossen")

    if ($isRepairable) {
        Write-Log "DISM RestoreHealth laeuft - bitte warten..." "Red"
        $dismRestoreRaw = & DISM /Online /Cleanup-Image /RestoreHealth 2>&1 | ForEach-Object { "$_" }
        foreach ($line in (Filter-Output $dismRestoreRaw)) { Write-Log $line "Green" }
        Write-Log "[OK] DISM RestoreHealth abgeschlossen." "Green"
    } elseif ($isClean) {
        Write-Log "[OK] DISM: Windows-Image ist sauber." "Green"
    } else {
        $last = ($dismCheckClean | Where-Object { $_.Trim().Length -gt 10 } | Select-Object -Last 3)
        foreach ($l in $last) { Write-Log "  >> $l" "Yellow" }
        Write-Log "[OK] DISM: Kein Fehler erkannt." "Green"
    }
} catch {
    Write-Log "[FEHLER] DISM: $($_.Exception.Message)" "Red"
}

# ================================================================
# 3/12 - TREIBER
# ================================================================
Write-Section "3/12 - Treiber-Diagnose"
Write-Log "Suche fehlerhafte Geraete - bitte warten..." "Red"
try {
    $faulty = Get-WmiObject Win32_PnPEntity |
        Where-Object { $_.ConfigManagerErrorCode -ne 0 } |
        Select-Object Name, ConfigManagerErrorCode, Status
    if ($faulty) {
        Write-Log "[FEHLER] Fehlerhafte Treiber gefunden:" "Red"
        foreach ($d in $faulty) {
            Write-Log "  [Code $($d.ConfigManagerErrorCode)] $($d.Name) - $($d.Status)" "Red"
        }
    } else {
        Write-Log "[OK] Keine Treiberfehler erkannt." "Green"
    }
} catch {
    Write-Log "[FEHLER] Treiber-Check: $($_.Exception.Message)" "Red"
}
Write-Log "Lese Drittanbieter-Treiber - bitte warten..." "Red"
try {
    $third = Get-WindowsDriver -Online |
        Where-Object { $_.ProviderName -notmatch "Microsoft" } |
        Select-Object ProviderName, OriginalFileName, Version
    if ($third) {
        foreach ($t in $third) {
            Write-Log "  $($t.ProviderName) | $($t.OriginalFileName) | v$($t.Version)" "Green"
        }
    } else {
        Write-Log "  Keine Drittanbieter-Treiber gefunden." "Green"
    }
} catch {
    Write-Log "  Treiber-Abfrage: $($_.Exception.Message)" "Yellow"
}

# ================================================================
# 4/12 - GPU INFO
# ================================================================
Write-Section "4/12 - GPU-Info"
Write-Log "Lese GPU-Informationen - bitte warten..." "Red"
try {
    $gpus = Get-WmiObject Win32_VideoController
    foreach ($gpu in $gpus) {
        $vramGB = if ($gpu.AdapterRAM -gt 0) { [math]::Round($gpu.AdapterRAM / 1GB, 1) } else { "N/A" }
        $col    = if ($gpu.Status -ne "OK") { "Yellow" } else { "Green" }
        Write-Log "  GPU: $(Clean-Text $gpu.Name)" $col
        Write-Log "       VRAM: $vramGB GB | Treiber: $($gpu.DriverVersion) | Status: $($gpu.Status)" $col
        Write-Log "       Aufloesung: $($gpu.CurrentHorizontalResolution)x$($gpu.CurrentVerticalResolution) | $($gpu.CurrentRefreshRate) Hz" $col
    }
} catch {
    Write-Log "  GPU-Abfrage: $($_.Exception.Message)" "Yellow"
}

# ================================================================
# 5/12 - TEMPERATUREN
# ================================================================
Write-Section "5/12 - Temperaturen (CPU und GPU)"
Write-Log "Lese CPU-Temperaturen - bitte warten..." "Red"
try {
    $thermalZones = @(Get-WmiObject -Namespace root\wmi -Class MSAcpi_ThermalZoneTemperature -ErrorAction Stop)
    if ($thermalZones.Count -gt 0) {
        $zoneNum = 0
        foreach ($zone in $thermalZones) {
            $tempC = [math]::Round(($zone.CurrentTemperature / 10) - 273.15, 1)
            if ($tempC -gt 0 -and $tempC -lt 150) {
                $col = if ($tempC -gt 90) { "Red" } elseif ($tempC -gt 75) { "Yellow" } else { "Green" }
                Write-Log "  CPU Zone $zoneNum : $tempC Grad C" $col
            }
            $zoneNum++
        }
    } else {
        Write-Log "  CPU-Temperatur: Nicht verfuegbar (BIOS/WMI unterstuetzt es nicht)." "Yellow"
    }
} catch {
    Write-Log "  CPU-Temperatur: Nicht unterstuetzt auf diesem System." "Yellow"
}
Write-Log "Lese NVIDIA GPU-Temperatur - bitte warten..." "Red"
try {
    $nvPaths = @(
        "C:\Program Files\NVIDIA Corporation\NVSMI\nvidia-smi.exe",
        "C:\Windows\System32\nvidia-smi.exe"
    )
    $nvSmi = $null
    foreach ($nvPath in $nvPaths) {
        if (Test-Path $nvPath) { $nvSmi = $nvPath; break }
    }
    if ($nvSmi) {
        $gpuTemp = & $nvSmi --query-gpu=temperature.gpu,name --format=csv,noheader 2>&1
        foreach ($line in $gpuTemp) {
            $parts = "$line".Split(',')
            if ($parts.Count -ge 2) {
                $tempVal = $parts[0].Trim()
                $gpuName = Clean-Text $parts[1].Trim()
                $col     = if ([int]$tempVal -gt 85) { "Red" } elseif ([int]$tempVal -gt 70) { "Yellow" } else { "Green" }
                Write-Log "  NVIDIA GPU: $gpuName - $tempVal Grad C" $col
            }
        }
    } else {
        Write-Log "  NVIDIA GPU-Temp: nvidia-smi nicht gefunden (kein NVIDIA oder Treiber fehlt)." "Yellow"
    }
} catch {
    Write-Log "  NVIDIA GPU-Temp: $($_.Exception.Message)" "Yellow"
}

# ================================================================
# 6/12 - AUTOSTART + VIRUSTOTAL
# ================================================================
Write-Section "6/12 - Autostart-Check + VirusTotal"
if ([string]::IsNullOrEmpty($VT_API_KEY)) {
    Write-Log "  [INFO] VirusTotal API-Key nicht gesetzt - VT-Scan deaktiviert." "Yellow"
    Write-Log "         Kostenlos: virustotal.com/gui/join-us" "Cyan"
} else {
    Write-Log "  [OK] VirusTotal API-Key gefunden - verdaechtige Dateien werden geprueft." "Green"
}
Write-Log "Lese Autostart-Eintraege - bitte warten..." "Red"
try {
    $startup = Get-CimInstance Win32_StartupCommand | Select-Object Name, Command, Location, User
    if ($startup) {
        foreach ($s in $startup) {
            $sus = $s.Command -match "temp|appdata\\roaming|\.vbs|\.bat|\.ps1|powershell|cmd\.exe /c"
            $col = if ($sus) { "Yellow" } else { "Green" }
            $vtResult = ""
            if ($sus -and -not [string]::IsNullOrEmpty($VT_API_KEY)) {
                $exePath = ($s.Command -split '"')[1]
                if (-not $exePath) { $exePath = $s.Command.Split(' ')[0] }
                if (Test-Path $exePath) {
                    $vt = Get-VTReport -FilePath $exePath
                    if ($vt) { $vtResult = " [VT: $vt]" }
                }
            }
            Write-Log "  [$($s.Location)] $($s.Name) -> $(Clean-Text $s.Command)$vtResult" $col
        }
    } else {
        Write-Log "  Keine WMI-Autostart-Eintraege gefunden." "Green"
    }
} catch {
    Write-Log "  Autostart WMI: $($_.Exception.Message)" "Yellow"
}
Write-Log "Pruefe Registry Run-Schluesse - bitte warten..." "Red"
$regPaths = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run",
    "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run",
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce",
    "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce"
)
foreach ($regPath in $regPaths) {
    try {
        $entries = Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue
        if ($entries) {
            $entries.PSObject.Properties | Where-Object { $_.Name -notmatch "^PS" } | ForEach-Object {
                $sus = $_.Value -match "temp|\.vbs|\.bat|powershell -enc|cmd /c"
                $col = if ($sus) { "Yellow" } else { "Green" }
                $vtResult = ""
                if ($sus -and -not [string]::IsNullOrEmpty($VT_API_KEY)) {
                    $exePath = ($_.Value -split '"')[1]
                    if (-not $exePath) { $exePath = $_.Value.Split(' ')[0] }
                    if (Test-Path $exePath) {
                        $vt = Get-VTReport -FilePath $exePath
                        if ($vt) { $vtResult = " [VT: $vt]" }
                    }
                }
                Write-Log "  [REG] $($_.Name) -> $(Clean-Text $_.Value)$vtResult" $col
            }
        }
    } catch {
        Write-Log "  Registry ${regPath}: $($_.Exception.Message)" "Yellow"
    }
}

# ================================================================
# 7/12 - FESTPLATTEN + SMART
# ================================================================
Write-Section "7/12 - Festplatten und SMART-Status"
Write-Log "Lese Laufwerke - bitte warten..." "Red"
try {
    $volumes = Get-Volume | Where-Object { $null -ne $_.DriveLetter }
    foreach ($v in $volumes) {
        $free  = [math]::Round($v.SizeRemaining / 1GB, 2)
        $total = [math]::Round($v.Size / 1GB, 2)
        $pct   = if ($total -gt 0) { [math]::Round(($v.SizeRemaining / $v.Size) * 100, 0) } else { 0 }
        $col   = if ($v.HealthStatus -ne "Healthy" -or $pct -lt 10) { "Yellow" } else { "Green" }
        Write-Log "  $($v.DriveLetter): $(Clean-Text $v.FileSystemLabel) | $free GB frei / $total GB ($pct%) | $($v.HealthStatus)" $col
    }
} catch {
    Write-Log "  Laufwerk-Abfrage: $($_.Exception.Message)" "Yellow"
}
Write-Log "Lese physische Festplatten - bitte warten..." "Red"
try {
    $disks = Get-PhysicalDisk | Select-Object FriendlyName, HealthStatus, OperationalStatus, MediaType, Size
    foreach ($d in $disks) {
        $sizeGB = [math]::Round($d.Size / 1GB, 0)
        $col    = if ($d.HealthStatus -ne "Healthy") { "Red" } else { "Green" }
        Write-Log "  $(Clean-Text $d.FriendlyName) | $($d.MediaType) | $sizeGB GB | $($d.HealthStatus) | $($d.OperationalStatus)" $col
    }
} catch {
    Write-Log "  Festplatten-Abfrage: $($_.Exception.Message)" "Yellow"
}
Write-Log "Pruefe SMART Failure Prediction - bitte warten..." "Red"
try {
    $smart = Get-WmiObject -Namespace root\wmi -Class MSStorageDriver_FailurePredictStatus -ErrorAction SilentlyContinue
    if ($smart) {
        foreach ($s in $smart) {
            if ($s.PredictFailure) {
                Write-Log "  [WARNUNG] SMART: Festplattenausfall vorhergesagt! -> $($s.InstanceName)" "Red"
            } else {
                Write-Log "  [OK] SMART: Kein Ausfall vorhergesagt." "Green"
            }
        }
    } else {
        Write-Log "  SMART: Keine Daten verfuegbar." "Yellow"
    }
} catch {
    Write-Log "  SMART: $($_.Exception.Message)" "Yellow"
}

# ================================================================
# 8/12 - RAM
# ================================================================
Write-Section "8/12 - Arbeitsspeicher (RAM)"
Write-Log "Lese RAM-Informationen - bitte warten..." "Red"
try {
    $os      = Get-WmiObject Win32_OperatingSystem
    $totalGB = [math]::Round($os.TotalVisibleMemorySize / 1MB, 2)
    $freeGB  = [math]::Round($os.FreePhysicalMemory / 1MB, 2)
    $usedGB  = [math]::Round(($os.TotalVisibleMemorySize - $os.FreePhysicalMemory) / 1MB, 2)
    $usedPct = [math]::Round(($usedGB / $totalGB) * 100, 0)
    $col     = if ($usedPct -gt 90) { "Red" } elseif ($usedPct -gt 75) { "Yellow" } else { "Green" }
    Write-Log "  Gesamt: $totalGB GB | Belegt: $usedGB GB ($usedPct%) | Frei: $freeGB GB" $col
} catch {
    Write-Log "  RAM OS-Abfrage: $($_.Exception.Message)" "Yellow"
}
try {
    $ramSlots = Get-WmiObject Win32_PhysicalMemory
    foreach ($r in $ramSlots) {
        $capGB = [math]::Round($r.Capacity / 1GB, 0)
        Write-Log "  Slot: $($r.DeviceLocator) | $capGB GB | $($r.Speed) MHz | $(Clean-Text $r.Manufacturer)" "Green"
    }
} catch {
    Write-Log "  RAM-Modul-Abfrage: $($_.Exception.Message)" "Yellow"
}
Write-Log "  [HINWEIS] RAM-Test: mdsched.exe ausfuehren (erfordert Neustart)" "Cyan"

# ================================================================
# 9/12 - NETZWERK
# ================================================================
Write-Section "9/12 - Netzwerk-Check"
Write-Log "Lese Netzwerk-Adapter - bitte warten..." "Red"
try {
    $adapters = Get-NetAdapter | Where-Object { $_.Status -eq "Up" }
    foreach ($a in $adapters) {
        Write-Log "  Adapter: $(Clean-Text $a.Name) | $(Clean-Text $a.InterfaceDescription) | $($a.LinkSpeed)" "Green"
    }
    $ips = Get-NetIPAddress | Where-Object { $_.AddressFamily -eq "IPv4" -and $_.IPAddress -notmatch "^127" }
    foreach ($ip in $ips) {
        Write-Log "  IP: $($ip.IPAddress)/$($ip.PrefixLength) | $(Clean-Text $ip.InterfaceAlias)" "Green"
    }
} catch {
    Write-Log "  Adapter-Abfrage: $($_.Exception.Message)" "Yellow"
}
Write-Log "Pruefe DNS und Gateway - bitte warten..." "Red"
try {
    $dns = Get-DnsClientServerAddress | Where-Object { $_.AddressFamily -eq 2 -and $_.ServerAddresses.Count -gt 0 }
    foreach ($d in $dns) {
        Write-Log "  DNS [$(Clean-Text $d.InterfaceAlias)]: $($d.ServerAddresses -join ', ')" "Green"
    }
    $gw = Get-NetRoute | Where-Object { $_.DestinationPrefix -eq "0.0.0.0/0" } | Select-Object -First 1
    if ($gw) {
        Write-Log "  Gateway: $($gw.NextHop) | $(Clean-Text $gw.InterfaceAlias)" "Green"
    } else {
        Write-Log "  [WARN] Kein Standard-Gateway gefunden!" "Yellow"
    }
} catch {
    Write-Log "  DNS/Gateway: $($_.Exception.Message)" "Yellow"
}
Write-Log "Ping-Tests laufen - bitte warten..." "Red"
try {
    $p1  = Test-NetConnection -ComputerName "8.8.8.8"    -WarningAction SilentlyContinue
    $col = if ($p1.PingSucceeded) { "Green" } else { "Red" }
    Write-Log "  Ping 8.8.8.8    (Google DNS) : $(if ($p1.PingSucceeded) { 'OK' } else { 'FEHLGESCHLAGEN' })" $col
} catch {
    Write-Log "  Ping 8.8.8.8: $($_.Exception.Message)" "Yellow"
}
try {
    $p2  = Test-NetConnection -ComputerName "google.com" -WarningAction SilentlyContinue
    $col = if ($p2.PingSucceeded) { "Green" } else { "Red" }
    Write-Log "  Ping google.com (DNS-Test)   : $(if ($p2.PingSucceeded) { 'OK' } else { 'FEHLGESCHLAGEN' })" $col
} catch {
    Write-Log "  Ping google.com: $($_.Exception.Message)" "Yellow"
}

# ================================================================
# 10/12 - GEPLANTE TASKS
# ================================================================
Write-Section "10/12 - Geplante Tasks (Drittanbieter)"
Write-Log "Suche aktive Drittanbieter-Tasks - bitte warten..." "Red"
try {
    $tasks = Get-ScheduledTask | Where-Object {
        $_.TaskPath -notmatch "\\Microsoft\\" -and $_.State -ne "Disabled"
    }
    if ($tasks) {
        foreach ($task in $tasks) {
            $exePath = ""
            if ($task.Actions.Count -gt 0) {
                $exePath = Clean-Text "$($task.Actions[0].Execute) $($task.Actions[0].Arguments)"
            }
            $sus = $exePath -match "temp|appdata\\roaming|\.vbs|powershell -enc|cmd /c|\.bat"
            $col = if ($sus) { "Yellow" } else { "Green" }
            Write-Log "  [$($task.State)] $(Clean-Text $task.TaskName) -> $exePath" $col
        }
    } else {
        Write-Log "  [OK] Keine Drittanbieter-Tasks gefunden." "Green"
    }
} catch {
    Write-Log "  Task-Abfrage: $($_.Exception.Message)" "Yellow"
}

# ================================================================
# 11/12 - EVENTLOG
# ================================================================
Write-Section "11/12 - EventLog (letzte 7 Tage)"
$startDate = (Get-Date).AddDays(-7)
$sources = @(
    @{Log="System";      Level=1; Label="Kritische Systemfehler"},
    @{Log="Application"; Level=2; Label="Anwendungsfehler"},
    @{Log="Setup";       Level=2; Label="Installationsfehler"}
)
foreach ($src in $sources) {
    Write-Log "-- $($src.Label) wird geprueft..." "Red"
    try {
        $events = Get-WinEvent -FilterHashtable @{
            LogName   = $src.Log
            Level     = $src.Level
            StartTime = $startDate
        } -ErrorAction SilentlyContinue | Select-Object -First 15
        if ($events) {
            foreach ($e in $events) {
                $msg = Clean-Text ($e.Message.Split([char]10)[0])
                Write-Log "  [$($e.TimeCreated)] ID:$($e.Id) - $msg" "Yellow"
            }
        } else {
            Write-Log "  [OK] Keine Eintraege." "Green"
        }
    } catch {
        Write-Log "  Log $($src.Log) nicht lesbar: $($_.Exception.Message)" "Yellow"
    }
}

# ================================================================
# 12/12 - WINDOWS UPDATE
# ================================================================
Write-Section "12/12 - Windows Update"
Write-Log "Pruefe Windows Update Verlauf - bitte warten..." "Red"
try {
    $wus = New-Object -ComObject Microsoft.Update.Session
    $wuh = $wus.CreateUpdateSearcher().QueryHistory(0, 30) |
        Where-Object { $_.ResultCode -eq 4 }
    if ($wuh) {
        foreach ($u in $wuh) { Write-Log "  [FEHLER] $(Clean-Text $u.Title) - $($u.Date)" "Yellow" }
    } else {
        Write-Log "  [OK] Keine fehlgeschlagenen Updates." "Green"
    }
} catch {
    Write-Log "  WU-History: $($_.Exception.Message)" "Yellow"
}

# ================================================================
# HTML REPORT
# ================================================================
Write-Section "HTML-Report wird generiert..."
try {
    $htmlHead = @"
<!DOCTYPE html>
<html lang="de">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>Kodimans Windows Tool v4.0.0</title>
<style>
  * { box-sizing:border-box; margin:0; padding:0; }
  body { background:#0a0a0a; color:#00ff41; font-family:'Courier New',Courier,monospace; padding:30px; }
  h1 { text-align:center; border-bottom:2px solid #00ff41; padding-bottom:15px;
       margin-bottom:5px; letter-spacing:4px; font-size:20px; }
  h2 { border-left:4px solid #00ff41; padding-left:12px;
       margin-top:25px; margin-bottom:8px; font-size:13px; letter-spacing:1px; }
  p  { font-size:12px; line-height:1.6; padding-left:5px; }
  .meta   { color:#444; text-align:center; margin-bottom:25px; font-size:12px; }
  .wrap   { border:1px solid #1a1a1a; border-radius:6px; padding:20px;
            background:#0d0d0d; box-shadow:0 0 20px rgba(0,255,65,0.05); }
  .footer { color:#333; text-align:center; margin-top:15px; font-size:11px; }
  .kofi   { text-align:center; margin-top:20px; padding:15px;
            border:1px solid #1a1a1a; border-radius:6px; background:#0d0d0d; }
  .kofi a { color:#00ff41; text-decoration:none; border-bottom:1px solid #00ff41;
            font-size:13px; letter-spacing:1px; }
  .kofi a:hover { color:#ffaa00; border-color:#ffaa00; }
  .ascii  { color:#00ff41; font-size:11px; text-align:center; line-height:1.3;
            margin-bottom:20px; white-space:pre; }
</style>
</head>
<body>
<div class="ascii">
  _  __         _ _
 | |/ /___   __| (_)_ __ ___   __ _ _ __  ___
 | ' // _ \ / _`| | '_ \` _ \ / _`| '_ \/ __|
 | . \ (_) | (_| | | | | | | | (_| | | | \__ \
 |_|\_\___/ \__,_|_|_| |_| |_|\__,_|_| |_|___/
 __        _____ _   _   ____  _
 \ \      / /_ _| \ | | |  _ \(_) __ _  __ _ _ __   ___  ___ ___
  \ \ /\ / / | ||  \| | | | | | |/ _`| / _`| '_ \ / _ \/ __|/ _ \
   \ V  V /  | || |\  | | |_| | | (_| | (_| | | | | (_) \__ \  __/
    \_/\_/  |___|_| \_| |____/|_|\__,_|\__, |_| |_|\___/|___/\___|
                                        |___/
</div>
<h1>KODIMANS WINDOWS TOOL v4.0.0</h1>
<p class="meta">System: $env:COMPUTERNAME &nbsp;&bull;&nbsp; $(Get-Date) &nbsp;&bull;&nbsp; by Kodiman_Himself &bull; Bremen</p>
<div class="wrap">
"@

    $htmlFoot = @"
</div>
<div class="kofi">
  &#9749; Dieses Tool hat dir geholfen?
  <a href="https://ko-fi.com/kodimanhimself" target="_blank">
    Spende mir einen Kaffee auf Ko-Fi &#9749;
  </a>
</div>
<p class="footer">Kodimans Windows Tool v4.0.0 &bull; Kodiman_Himself &bull; Bremen &bull; $(Get-Date -Format 'yyyy')</p>
</body>
</html>
"@

    $fullHtml = $htmlHead + ($HtmlEntries -join "`n") + $htmlFoot
    Out-File -FilePath $HtmlFile -Encoding UTF8 -InputObject $fullHtml
    Write-Log "[OK] HTML-Report gespeichert: $HtmlFile" "Green"
} catch {
    Write-Log "[FEHLER] HTML-Report: $($_.Exception.Message)" "Red"
}

# ================================================================
# ABSCHLUSS
# ================================================================
Write-Section "Diagnose abgeschlossen"
Write-Log "LOG-Datei:   $LogFile"  "Green"
Write-Log "HTML-Report: $HtmlFile" "Green"
Write-Log "Fertig:      $(Get-Date)" "Green"
Write-Log "" "Green"
Write-Log "  Dieses Tool hat dir geholfen? Spende mir einen Kaffee!" "Cyan"
Write-Log "  Ko-Fi: https://ko-fi.com/kodimanhimself" "Cyan"

$openLog  = Read-Host "LOG-Datei oeffnen?               (J/N)"
if ($openLog  -match "^[Jj]") { Start-Process notepad.exe $LogFile }
$openHtml = Read-Host "HTML-Report im Browser oeffnen?  (J/N)"
if ($openHtml -match "^[Jj]") { Start-Process $HtmlFile }
