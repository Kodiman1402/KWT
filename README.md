
![Version](https://img.shields.io/badge/Version-4.0.0-00ff41?style=flat-square&labelColor=0a0a0a)
![PowerShell](https://img.shields.io/badge/PowerShell-5.1+-00ff41?style=flat-square&labelColor=0a0a0a&logo=powershell)
![Platform](https://img.shields.io/badge/Platform-Windows%2010%2F11-00ff41?style=flat-square&labelColor=0a0a0a&logo=windows)
![License](https://img.shields.io/badge/License-MIT-00ff41?style=flat-square&labelColor=0a0a0a)
[![Ko-Fi](https://img.shields.io/badge/Ko--Fi-Spende%20einen%20Kaffee-ff5f5f?style=flat-square&labelColor=0a0a0a&logo=ko-fi)](https://ko-fi.com/kodimanhimself)

**Kodimans Windows Tool** ist ein vollstaendiges Windows-Diagnosetool als PowerShell-Skript und EXE.  
Entwickelt von **Kodiman_Himself** aus Bremen.

</div>

---

## Features

| Modul | Beschreibung |
|---|---|
| **1/12 SFC** | System File Checker - prueft Systemdateien |
| **2/12 DISM** | Windows-Image Pruefung und Reparatur |
| **3/12 Treiber** | Fehlerhafte Treiber und Drittanbieter |
| **4/12 GPU** | GPU-Name, VRAM, Treiber, Aufloesung |
| **5/12 Temperaturen** | CPU-Zonen + NVIDIA GPU via nvidia-smi |
| **6/12 Autostart** | WMI + Registry + optionaler VirusTotal-Scan |
| **7/12 Festplatten** | Laufwerke, SMART Failure Prediction |
| **8/12 RAM** | Belegung, Slots, Hersteller, Takt |
| **9/12 Netzwerk** | Adapter, IP, DNS, Gateway, Ping |
| **10/12 Tasks** | Aktive Drittanbieter-Tasks |
| **11/12 EventLog** | Kritische Fehler der letzten 7 Tage |
| **12/12 Windows Update** | Fehlgeschlagene Updates |

### Output
- **LOG-Datei** `.log` auf dem Desktop
- **HTML-Report** `.html` im Matrix Dark-Mode Design

---

## Installation und Start

### Als EXE (empfohlen)
1. [Neueste Version herunterladen](releases/)
2. Doppelklick auf `Kodimans_Windows_Tool.exe`
3. UAC-Prompt mit **Ja** bestaetigen
4. Fertig!

### Als PowerShell-Skript
```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\Kodimans_Windows_Tool.ps1

Hinweis: Das Tool benoetigt Administratorrechte.
