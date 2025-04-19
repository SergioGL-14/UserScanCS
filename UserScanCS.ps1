# ================================================================
# UserScanCS Refactorizado - Escaneo paralelo de usuarios en equipos
# Versión: 1.1
# ================================================================

#region Parámetros y Constantes
# Ruta al CSV de configuración
$csvPath = "..\..\data\csv\COPIA_SERGAS_CUENTAS_AD_FULL.csv"
# Herramienta PsExec
$psexecPath = "..\..\resources\tools\PsExec.exe"
# Archivo de log
$logFile = "$PSScriptRoot\UserScanCS.log"
#endregion

# Iniciar registro de ejecución
try {
    Start-Transcript -Path $logFile -Append -ErrorAction Stop
} catch {
    Write-Warning "No se pudo iniciar el registro: $_"
}

#region Funciones Principales
function Load-Configuration {
    <#
    .SYNOPSIS
        Carga centros y equipos desde el CSV.
    .OUTPUTS
        Hashtable con clave Centro y valor array de Equipos.
    #>
    if (-not (Test-Path $csvPath)) {
        Throw "Archivo de configuración no encontrado: $csvPath"
    }
    $table = @{}
    Import-Csv -Path $csvPath -Delimiter ';' | ForEach-Object {
        $key = $_.OU.Trim()
        $machine = $_.Equipo.Trim()
        if (-not $table.ContainsKey($key)) { $table[$key] = @() }
        if ($machine -and -not $table[$key].Contains($machine)) {
            $table[$key] += $machine
        }
    }
    return $table
}

function Get-LastLoggedOnUser {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ComputerName
    )
    try {
        $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $ComputerName)
        $key = $reg.OpenSubKey('SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Authentication\\LogonUI')
        return $key?.GetValue('LastLoggedOnUser') -or ''
    } catch {
        return ''
    }
}

function Get-ADNameSurname {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SamAccountName
    )
    try {
        if (-not (Get-Module ActiveDirectory)) { Import-Module ActiveDirectory -ErrorAction Stop }
        $user = Get-ADUser -Filter { SamAccountName -eq $SamAccountName } -Properties GivenName, Surname -ErrorAction Stop
        return @{ GivenName = $user.GivenName; Surname = $user.Surname }
    } catch {
        return @{ GivenName = ''; Surname = '' }
    }
}

function Scan-Computer {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Computer,
        [Parameter(Mandatory)][string]$PsexecPath
    )
    $result = [PSCustomObject]@{
        Computer    = $Computer
        Status      = 'Inaccesible'
        Login       = ''
        GivenName   = ''
        Surname     = ''
    }
    # Verificar conexión
    if (-not (Test-Connection -ComputerName $Computer -Count 1 -Quiet)) { return $result }

    # Intentar quser remota
    $user = ''
    try {
        if (Test-Path $PsexecPath) {
            $output = & $PsexecPath -nobanner "\\$Computer" quser 2>&1
            foreach ($line in $output) {
                if ($line -match '^\s*(\S+)') {
                    $user = $matches[1]; break
                }
            }
        }
    } catch { }

    if ($user) {
        $result.Status = 'Ocupado'
    } else {
        $user = Get-LastLoggedOnUser -ComputerName $Computer
        $result.Status = if ($user) { 'Libre' } else { 'Libre (SIN ACCESO)' }
    }
    $result.Login = $user

    if ($user) {
        $adInfo = Get-ADNameSurname -SamAccountName ($user.Split('\\')[-1])
        $result.GivenName = $adInfo.GivenName
        $result.Surname   = $adInfo.Surname
    }
    return $result
}
#endregion

#region Interfaz Gráfica - Diseño Moderno
# Carga ensamblados
Add-Type -AssemblyName System.Windows.Forms, System.Drawing

# Form principal
$form = New-Object System.Windows.Forms.Form -Property @{ 
    Text            = 'UserScanCS - Escaneo de Usuarios'
    Size            = [System.Drawing.Size]::new(820, 700)
    StartPosition   = 'CenterScreen'
    BackColor       = [System.Drawing.Color]::FromArgb(245,245,245)
    Font            = 'Segoe UI,10'
}

# ComboBox de Centros
$comboCenters = [System.Windows.Forms.ComboBox]::new()
$comboCenters.SetBounds(20,20,450,30)
$comboCenters.DropDownStyle     = 'DropDown'
$comboCenters.AutoCompleteMode   = 'SuggestAppend'
$comboCenters.AutoCompleteSource = 'ListItems'
$form.Controls.Add($comboCenters)

# Botón Escanear
$btnScan = [System.Windows.Forms.Button]::new()
$btnScan.Text      = 'Escanear'
$btnScan.SetBounds(490,18,120,30)
$btnScan.FlatStyle = 'Flat'
$btnScan.BackColor  = [System.Drawing.Color]::FromArgb(0,120,215)
$btnScan.ForeColor  = 'White'
$form.Controls.Add($btnScan)

# ProgressBar de escaneo
$progress = [System.Windows.Forms.ProgressBar]::new()
$progress.SetBounds(20,60,760,20)
$progress.Style = 'Continuous'
$form.Controls.Add($progress)

# ListView de Resultados
$listView = [System.Windows.Forms.ListView]::new()
$listView.SetBounds(20,100,760,440)
$listView.View         = 'Details'
$listView.FullRowSelect= $true
$listView.GridLines    = $true
$listView.Columns.Add('Equipo',150)
$listView.Columns.Add('Estado',120)
$listView.Columns.Add('Login',120)
$listView.Columns.Add('Nombre',180)
$listView.Columns.Add('Apellidos',180)
$form.Controls.Add($listView)

# Label de Estado\ `$lblStatus = [System.Windows.Forms.Label]::new()
$lblStatus.SetBounds(20,560,500,25)
$lblStatus.Text       = 'Listo.'
$lblStatus.ForeColor  = 'DarkSlateGray'
$form.Controls.Add($lblStatus)

# Botón Exportar
$btnExport = [System.Windows.Forms.Button]::new()
$btnExport.Text      = 'Exportar CSV'
$btnExport.SetBounds(620,555,160,30)
$btnExport.FlatStyle = 'Flat'
$btnExport.BackColor  = [System.Drawing.Color]::FromArgb(0,120,215)
$btnExport.ForeColor  = 'White'
$form.Controls.Add($btnExport)
#endregion

#region Eventos
# Cargar configuración al iniciar
try {
    $config = Load-Configuration
    $comboCenters.Items.AddRange($config.Keys | Sort-Object)
} catch {
    [System.Windows.Forms.MessageBox]::Show("Error al cargar configuración:`n$_","Error","OK","Error")
    Exit
}

# Manejar clic Escanear
$btnScan.Add_Click({
    $listView.Items.Clear()
    $lblStatus.Text   = 'Iniciando escaneo...'
    $progress.Value   = 0
    $form.Refresh()

    $center = $comboCenters.Text.Trim()
    if (-not $config.ContainsKey($center)) {
        [System.Windows.Forms.MessageBox]::Show('Centro no válido.','Error','OK','Error')
        return
    }

    $machines = $config[$center]
    $count    = $machines.Count; $i=0
    $runspacePool = [runspacefactory]::CreateRunspacePool(1,10)
    $runspacePool.Open()
    $jobs = @()

    foreach ($machine in $machines) {
        $psInstance = [powershell]::Create()
        $psInstance.RunspacePool = $runspacePool
        $psInstance.AddScript(${function:Scan-Computer}).AddArgument($machine).AddArgument($psexecPath) | Out-Null
        $jobs += @{ Instance = $psInstance; Async = $psInstance.BeginInvoke() }
    }

    foreach ($job in $jobs) {
        $res = $job.Instance.EndInvoke($job.Async)
        $job.Instance.Dispose()
        $item = [System.Windows.Forms.ListViewItem]::new($res.Computer)
        $item.SubItems.Add($res.Status)
        $item.SubItems.Add($res.Login)
        $item.SubItems.Add($res.GivenName)
        $item.SubItems.Add($res.Surname)
        switch ($res.Status) {
            'Ocupado'            { $item.ForeColor = 'DarkOrange' }
            'Libre'              { $item.ForeColor = 'DarkGreen' }
            'Libre (SIN ACCESO)' { $item.ForeColor = 'Gray' }
            default              { $item.ForeColor = 'DarkRed' }
        }
        $listView.Items.Add($item)

        # Actualizar progreso
        $i++;
        $progress.Value    = [int](( $i / $count ) * 100)
        $lblStatus.Text    = "Escaneado $i de $count equipos"
        $form.Refresh()
    }

    $runspacePool.Close(); $runspacePool.Dispose()
    $lblStatus.Text = "Escaneo completado: $center"
})

# Manejar clic Exportar
$btnExport.Add_Click({
    if ($listView.Items.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('Nada para exportar.','Aviso','OK','Information')
        return
    }
    $dlg = [System.Windows.Forms.SaveFileDialog]::new()
    $dlg.Filter   = 'CSV (*.csv)|*.csv'
    $dlg.FileName = "Resultados_UserScanCS_$(Get-Date -Format yyyyMMdd).csv"
    if ($dlg.ShowDialog() -eq 'OK') {
        $file = $dlg.FileName
        $listView.Items | ForEach-Object {
            ($_.SubItems | ForEach-Object { $_.Text }) -join ';'
        } | Set-Content -Path $file -Encoding UTF8
        [System.Windows.Forms.MessageBox]::Show('Exportación exitosa.','Éxito','OK','Information')
    }
})

# Mostrar formulario
$form.Add_Shown({$form.Activate()})
[void] $form.ShowDialog()

# Finalizar registro
try {
    Stop-Transcript
} catch {
}
