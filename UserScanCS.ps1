# UserScanCS - Escaneo paralelo de usuarios en equipos con RunspacePool

##############################################################################
# BLOQUE 1: Carga de CSV y diccionario Centro → Equipos
##############################################################################
$RutaCSV = ".\csv\COPIA_SERGAS_CUENTAS_AD_FULL.csv"
$CentrosEquipos = @{}

if (Test-Path $RutaCSV) {
    try {
        $Datos = Import-Csv -Path $RutaCSV -Delimiter ';'
        foreach ($linea in $Datos) {
            $centro = $linea.OU.Trim()
            $equipo = $linea.Equipo.Trim()
            if (-not $CentrosEquipos.ContainsKey($centro)) {
                $CentrosEquipos[$centro] = @()
            }
            $CentrosEquipos[$centro] += $equipo
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error al cargar el CSV: $_", "Error", "OK", "Error")
        return
    }
} else {
    [System.Windows.Forms.MessageBox]::Show("Archivo CSV no encontrado en la ruta:`n$RutaCSV", "Error", "OK", "Error")
    return
}

##############################################################################
# BLOQUE 2: Interfaz gráfica
##############################################################################
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
$form = New-Object System.Windows.Forms.Form
$form.Text = "UserScanCS - Escaneo de Usuarios en Equipos"
$form.Size = New-Object System.Drawing.Size(800,660)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::FromArgb(245,245,245)
$form.Font = "Segoe UI,9"

$comboCentros = New-Object System.Windows.Forms.ComboBox
$comboCentros.Location = New-Object System.Drawing.Point(20,20)
$comboCentros.Size = New-Object System.Drawing.Size(400,30)
$comboCentros.DropDownStyle = 'DropDown'
$comboCentros.AutoCompleteMode = 'SuggestAppend'
$comboCentros.AutoCompleteSource = 'ListItems'
$comboCentros.Items.AddRange(($CentrosEquipos.Keys | Sort-Object))
$form.Controls.Add($comboCentros)

$btnEscanear = New-Object System.Windows.Forms.Button
$btnEscanear.Text = "Escanear Equipos"
$btnEscanear.Location = New-Object System.Drawing.Point(440, 18)
$btnEscanear.Size = New-Object System.Drawing.Size(150,30)
$btnEscanear.BackColor = [System.Drawing.Color]::FromArgb(0,120,215)
$btnEscanear.ForeColor = "White"
$btnEscanear.FlatStyle = "Flat"
$form.Controls.Add($btnEscanear)

$listEquipos = New-Object System.Windows.Forms.ListView
$listEquipos.Location = New-Object System.Drawing.Point(20, 70)
$listEquipos.Size = New-Object System.Drawing.Size(760, 440)
$listEquipos.View = 'Details'
$listEquipos.FullRowSelect = $true
$listEquipos.GridLines = $true
$listEquipos.Columns.Add("Equipo", 140)
$listEquipos.Columns.Add("Estado", 140)
$listEquipos.Columns.Add("Login", 120)
$listEquipos.Columns.Add("Nombre", 140)
$listEquipos.Columns.Add("Apellidos", 180)
$form.Controls.Add($listEquipos)

$lblEstado = New-Object System.Windows.Forms.Label
$lblEstado.Location = New-Object System.Drawing.Point(20,520)
$lblEstado.Size = New-Object System.Drawing.Size(600,20)
$lblEstado.Text = "Listo para escanear..."
$lblEstado.ForeColor = "DarkSlateGray"
$form.Controls.Add($lblEstado)

$btnExportar = New-Object System.Windows.Forms.Button
$btnExportar.Text = "Exportar Resultados"
$btnExportar.Location = New-Object System.Drawing.Point(620, 520)
$btnExportar.Size = New-Object System.Drawing.Size(160, 30)
$btnExportar.BackColor = [System.Drawing.Color]::FromArgb(0,120,215)
$btnExportar.ForeColor = "White"
$btnExportar.FlatStyle = "Flat"
$form.Controls.Add($btnExportar)

##############################################################################
# BLOQUE 3: Escaneo paralelo usando RunspacePool
##############################################################################
$psexecPath = ".\\NRC_APP\\tools\\PsExec.exe"
$btnEscanear.Add_Click({
    $listEquipos.Items.Clear()
    $lblEstado.Text = "Iniciando escaneo..."
    $form.Refresh()

    $centroSeleccionado = $comboCentros.Text
    if (-not $CentrosEquipos.ContainsKey($centroSeleccionado)) {
        [System.Windows.Forms.MessageBox]::Show("Centro no válido o no encontrado.","Error","OK","Error")
        return
    }

    $equipos = $CentrosEquipos[$centroSeleccionado]
    $runspacePool = [runspacefactory]::CreateRunspacePool(1,10)
    $runspacePool.Open()
    $runspaces = @()

    foreach ($equipo in $equipos) {
        $ps = [powershell]::Create()
        $ps.RunspacePool = $runspacePool

        $script = {
            param($equipo, $psexecPath)

            function Get-ADNombreApellidos {
                param ([string]$login)
                if (-not $login) { return @("","") }
                try {
                    $samAccountName = $login.Split('\')[-1]
                    $user = Get-ADUser -Filter { SamAccountName -eq $samAccountName } -Properties GivenName, Surname
                    if ($user) {
                        return @($user.GivenName, $user.Surname)
                    }
                } catch {
                    return @("","")
                }
                return @("","")
            }

            function Get-LastLoggedUser {
                param ([string]$Equipo)
                try {
                    $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $Equipo)
                    $key = $reg.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Authentication\\LogonUI")
                    if ($key) {
                        $valor = $key.GetValue("LastLoggedOnUser")
                        if ($valor) {
                            return $valor
                        }
                    }
                } catch {
                    return ""
                }
                return ""
            }

            $resultado = @{ Equipo=$equipo; Estado='Inaccesible'; Login=''; Nombre=''; Apellidos='' }

            if (Test-Connection -ComputerName $equipo -Count 1 -Quiet -ErrorAction SilentlyContinue) {
                $usuario = ""

                try {
                    if (Test-Path $psexecPath) {
                        $quserOutput = & $psexecPath -nobanner "\\$equipo" quser 2>&1
                        foreach ($line in $quserOutput) {
                            if ($line -match "^\s*(\S+)\s+[\w\s]+\s+(Activo|Active)") {
                                $usuario = $matches[1]
                                break
                            }
                        }
                    }
                } catch {}

                if ($usuario) {
                    $resultado.Estado = "Ocupado"
                } else {
                    $usuario = Get-LastLoggedUser -Equipo $equipo
                    if ($usuario) {
                        $resultado.Estado = "Libre"
                    } else {
                        $resultado.Estado = "Libre (SIN ACCESO)"
                    }
                }

                $resultado.Login = $usuario
                if ($usuario) {
                    $res = Get-ADNombreApellidos -login $usuario
                    $resultado.Nombre = $res[0]
                    $resultado.Apellidos = $res[1]
                }
            }

            return $resultado
        }

        $ps.AddScript($script).AddArgument($equipo).AddArgument($psexecPath)
        $async = $ps.BeginInvoke()
        $runspaces += @{ Pipe=$ps; Handle=$async }
    }

    foreach ($r in $runspaces) {
        $resultado = $r.Pipe.EndInvoke($r.Handle)[0]
        $r.Pipe.Dispose()

        $item = New-Object System.Windows.Forms.ListViewItem($resultado.Equipo)
        $item.SubItems.Add($resultado.Estado)
        $item.SubItems.Add($resultado.Login)
        $item.SubItems.Add($resultado.Nombre)
        $item.SubItems.Add($resultado.Apellidos)

        switch ($resultado.Estado) {
            "Ocupado"               { $item.ForeColor = "DarkOrange" }
            "Libre"                 { $item.ForeColor = "DarkGreen" }
            "Libre (SIN ACCESO)"    { $item.ForeColor = "Gray" }
            default                 { $item.ForeColor = "DarkRed" }
        }

        $listEquipos.Items.Add($item)
        $form.Refresh()
    }

    $runspacePool.Close()
    $runspacePool.Dispose()
    $lblEstado.Text = "Escaneo completado para centro: $centroSeleccionado"
})

##############################################################################
# BLOQUE 4: Exportación de resultados
##############################################################################
$btnExportar.Add_Click({
    if ($listEquipos.Items.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No hay resultados para exportar.","Aviso","OK","Information")
        return
    }

    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.Filter = "CSV Files (*.csv)|*.csv|Text Files (*.txt)|*.txt"
    $saveDialog.Title = "Guardar resultados"
    $saveDialog.FileName = "Resultados_UserScanCS.csv"

    if ($saveDialog.ShowDialog() -eq "OK") {
        $ruta = $saveDialog.FileName
        $contenido = @()

        foreach ($item in $listEquipos.Items) {
            $linea = @()
            foreach ($subItem in $item.SubItems) {
                $linea += $subItem.Text
            }
            $contenido += ($linea -join ";")
        }

        try {
            $contenido | Set-Content -Path $ruta -Encoding UTF8
            [System.Windows.Forms.MessageBox]::Show("Exportación completada con éxito.","Éxito","OK","Information")
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error al guardar archivo: $_","Error","OK","Error")
        }
    }
})

##############################################################################
# Final: Mostrar formulario
##############################################################################
$form.Add_Shown({ $form.Activate() })
$form.ShowDialog()
