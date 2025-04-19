# 🖥️ FleetADScan - Escaneo Paralelo de Usuarios en Equipos (PowerShell GUI)

**FleetADScan** es una herramienta desarrollada en PowerShell con interfaz gráfica (WinForms) que permite escanear remotamente los equipos de un centro de salud o unidad organizativa (OU), detectando qué usuarios están logueados, si el equipo está libre o inaccesible, y mostrando nombre y apellidos extraídos de Active Directory.

---

## 📋 Características

- Escaneo **paralelo** de múltiples equipos mediante `RunspacePool`.
- Detección de equipos:
  - Ocupados (usuario conectado)
  - Libres (último usuario conocido)
  - Inaccesibles o sin permisos
- Extracción de nombre y apellidos desde **Active Directory** (`Get-ADUser`).
- Lectura desde fichero CSV con relación **Centro → Equipos**.
- Interfaz WinForms moderna y práctica.
- **Exportación de resultados** a CSV o TXT.
- Soporte para ejecución local o remota usando **PsExec**.

---

## ⚙️ Requisitos

- PowerShell 5.1 (o superior en Windows)
- Módulo de Active Directory (`RSAT: Active Directory`)
- Permisos administrativos para acceder remotamente a equipos
- PsExec (incluido en la ruta: `.\NRC_APP\tools\PsExec.exe`)
- Archivo CSV con formato:

