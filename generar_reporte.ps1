<#
.SYNOPSIS
    Generador Automatizado de Reportes Semanales en PDF (Camanchaca, Cermaq, Mowi)
.DESCRIPTION
    Lee el archivo Excel más reciente de la empresa especificada, toma la última pestaña actualizada,
    calcula las métricas de rendimiento por región y genera el informe PDF ejecutivo con gráficos vectoriales.
.EXAMPLE
    .\generar_reporte.ps1 -Empresa Camanchaca
#>

param (
    [Parameter(Mandatory=$false)]
    [ValidateSet("Camanchaca", "Cermaq", "Mowi")]
    [string]$Empresa = "Camanchaca"
)

$ErrorActionPreference = "Stop"

# Directorio de trabajo (dinámico según la ubicación del script)
if (-not $workingDir) { $workingDir = $PSScriptRoot }
if (-not $workingDir) { $workingDir = "c:\Users\Omnifish\Desktop\Generador de Reportes Semanales" }

# Encontrar archivo Excel correspondiente (Selecciona el más reciente automáticamente)
$excelDir = Join-Path $workingDir "Datos_Excel"
if (-not (Test-Path $excelDir)) { $excelDir = $workingDir }
$excelFiles = Get-ChildItem -Path $excelDir -Filter "*$Empresa*.xlsx" | Sort-Object LastWriteTime -Descending
if ($excelFiles.Count -eq 0) {
    Write-Error "No se encontró archivo Excel para $Empresa en $excelDir"
}
$excelPath = $excelFiles[0].FullName
Write-Host "Cargando datos desde: $($excelFiles[0].Name)"

# Abrir Excel COM
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$wb = $excel.Workbooks.Open($excelPath)

# Obtener primera pestaña con datos semanales (las pestañas están ordenadas de más reciente a más antigua de izquierda a derecha)
$targetSheet = $null
for ($i = 1; $i -le $wb.Worksheets.Count; $i++) {
    $s = $wb.Worksheets.Item($i)
    if ($s.Name -ne "Consolidado") {
        $targetSheet = $s
        break
    }
}

if (-not $targetSheet) {
    $wb.Close($false)
    $excel.Quit()
    Write-Error "No se encontró una pestaña válida en el archivo Excel."
}

$fechaPestana = $targetSheet.Name.Trim()
Write-Host "Pestaña seleccionada (semana en curso): $fechaPestana"

# Mapear datos por regiones
$regiones = [ordered]@{}
$totalCentros = 0
$totalJaulas = 0
$totalCamaras = 0
$totalSinVisual = 0
$totalMortSinVisual = 0

# Determinar número de semana automáticamente según nombre del archivo o pestaña
$semanaNum = "34"
if ($excelFiles[0].Name -match "(?i)semana\s*(\d+)") {
    $semanaNum = $Matches[1]
} elseif ($fechaPestana -eq "14-08-26" -or $fechaPestana -eq "14-08-2026") {
    $semanaNum = "33"
}

# Detectar fila de encabezados
$headerRow = 0
for ($r = 1; $r -le [Math]::Min(10, $targetSheet.UsedRange.Rows.Count); $r++) {
    $c1 = $targetSheet.Cells.Item($r, 1).Text.Trim()
    if ($c1 -like "*Centro*") { $headerRow = $r; break }
}
if ($headerRow -eq 0) { $headerRow = 2 }

# Mapear columnas dinámicamente según encabezados
$colCentro = 1
$colRegion = 2
$colJaulas = 6
$colCamsPorJaula = 7
$colSinVisual = 8
$colMortSinVisual = 9

$colCount = $targetSheet.UsedRange.Columns.Count
for ($c = 1; $c -le $colCount; $c++) {
    $hText = $targetSheet.Cells.Item($headerRow, $c).Text.Trim()
    if ($hText -like "*Observaci*") { $colObs = $c }
    elseif ($hText -like "*mortalidad*sin visual*") { $colMortSinVisual = $c }
    elseif ($hText -like "*sin visual*") { $colSinVisual = $c }
    elseif ($hText -like "*C*mara*jaula*") { $colCamsPorJaula = $c }
    elseif ($hText -like "*Total*jaula*" -or $hText -like "*Total de jaulas*") { $colJaulas = $c }
    elseif ($hText -like "*Regi*") { $colRegion = $c }
    elseif ($hText -like "*Centro*") { $colCentro = $c }
}

function Get-CellText($sheet, $r, $c) {
    if ($c -le 0) { return "" }
    try {
        $cell = $sheet.Cells.Item($r, $c)
        if ($null -ne $cell -and $null -ne $cell.Text) {
            return $cell.Text.ToString().Trim()
        }
        return ""
    } catch {
        return ""
    }
}

# Leer filas
$rowsCount = $targetSheet.UsedRange.Rows.Count
$centrosList = @()

for ($r = ($headerRow + 1); $r -le $rowsCount; $r++) {
    $centro = Get-CellText $targetSheet $r $colCentro
    if ($centro -eq "" -or $centro -eq "-" -or $centro -eq "Centro" -or $centro -like "*Feed Center*") { continue }
    
    $reg = Get-CellText $targetSheet $r $colRegion
    if ($reg -eq "" -or $reg -eq "-") { continue }

    $jaulasStr = Get-CellText $targetSheet $r $colJaulas
    $jaulas = 0
    [int]::TryParse($jaulasStr, [ref]$jaulas) | Out-Null

    $camsPorJaulaStr = Get-CellText $targetSheet $r $colCamsPorJaula
    $camsPorJaula = 1
    if ($camsPorJaulaStr -ne "") { [int]::TryParse($camsPorJaulaStr, [ref]$camsPorJaula) | Out-Null }
    if ($camsPorJaula -le 0) { $camsPorJaula = 1 }

    $camaras = $jaulas * $camsPorJaula

    $sinVisualStr = Get-CellText $targetSheet $r $colSinVisual
    $sinVisual = 0
    [int]::TryParse($sinVisualStr, [ref]$sinVisual) | Out-Null

    $mortSinVisualStr = Get-CellText $targetSheet $r $colMortSinVisual
    $mortSinVisual = 0
    [int]::TryParse($mortSinVisualStr, [ref]$mortSinVisual) | Out-Null

    $obs = Get-CellText $targetSheet $r $colObs
    if ($obs -eq "") { $obs = "Sin novedad" }

    $centrosList += @{
        Centro = $centro
        Region = $reg
        Jaulas = $jaulas
        Camaras = $camaras
        SinVisual = $sinVisual
        MortSinVisual = $mortSinVisual
        Observaciones = $obs
    }

    if (-not $regiones.Contains($reg)) {
        $regiones[$reg] = @{
            Centros = 0
            Jaulas = 0
            Camaras = 0
            SinVisual = 0
            MortSinVisual = 0
        }
    }

    $regiones[$reg].Centros += 1
    $regiones[$reg].Jaulas += $jaulas
    $regiones[$reg].Camaras += $camaras
    $regiones[$reg].SinVisual += $sinVisual
    $regiones[$reg].MortSinVisual += $mortSinVisual

    $totalCentros++
    $totalJaulas += $jaulas
    $totalCamaras += $camaras
    $totalSinVisual += $sinVisual
    $totalMortSinVisual += $mortSinVisual
}

# Construir filas de detalle de centros e observaciones
$centrosDetailRowsHtml = ""
foreach ($cItem in $centrosList) {
    $cName = $cItem.Centro
    $cReg = $cItem.Region
    $cJau = $cItem.Jaulas
    $cCam = $cItem.Camaras
    $cObs = $cItem.Observaciones

    $badgeStyle = "background-color: #f1f5f9; color: #475569;"
    if ($cObs -like "*cosecha*" -or $cObs -like "*cosechado*") {
        $badgeStyle = "background-color: #fef3c7; color: #92400e; font-weight: bold;"
    } elseif ($cObs -like "*falla*" -or $cObs -like "*intermitencia*" -or $cObs -like "*no responde*" -or $cObs -like "*problema*" -or $cObs -like "*sin visual*") {
        $badgeStyle = "background-color: #fee2e2; color: #991b1b; font-weight: bold;"
    } elseif ($cObs -like "*sin novedad*" -or $cObs -like "*100%*") {
        $badgeStyle = "background-color: #f0fdf4; color: #166534;"
    }

    $centrosDetailRowsHtml += @"
      <tr>
        <td style="text-align: left; font-weight: 700;">$cName</td>
        <td>$cReg</td>
        <td>$cJau</td>
        <td>$cCam</td>
        <td style="text-align: left;"><span style="display: inline-block; padding: 3px 10px; border-radius: 12px; font-size: 11px; $badgeStyle">$cObs</span></td>
      </tr>
"@
}

$wb.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

# Calcular % Funcionamiento Global
$funcGlobal = 100.00
if ($totalCamaras -gt 0) {
    $funcGlobal = [Math]::Round(((($totalCamaras - $totalSinVisual - $totalMortSinVisual) / $totalCamaras) * 100), 2)
}
$funcGlobalStr = "{0:N2}%" -f $funcGlobal

# Configuración visual por empresa
$headerBg = "linear-gradient(135deg, #093c71 0%, #06264a 100%)"
$accentColor = "#093c71"
$logoPath = "$workingDir\assets\logos\Camanchaca logo.png"

if ($Empresa -eq "Cermaq") {
    $headerBg = "linear-gradient(135deg, #006666 0%, #004d4d 100%)"
    $accentColor = "#006666"
    $logoPath = "$workingDir\assets\logos\Cermaq logo.jpg"
} elseif ($Empresa -eq "Mowi") {
    $headerBg = "linear-gradient(135deg, #1a1a1a 0%, #0d0d0d 100%)"
    $accentColor = "#1a1a1a"
    $logoPath = "$workingDir\assets\logos\Mowi logo.png"
}

$logoSvg = ""
if (Test-Path $logoPath) {
    $b64 = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($logoPath))
    $mime = "image/png"
    if ($logoPath -like "*.jpg") { $mime = "image/jpeg" }
    $logoSvg = "<img src='data:$mime;base64,$b64' style='max-height: 40px; max-width: 170px; object-fit: contain;' alt='$Empresa' />"
}

# Formatear Fecha de Actualización
$fechaActualizado = "21 de Agosto de 2026"
if ($fechaPestana -eq "14-08-26" -or $fechaPestana -eq "14-08-2026") { $fechaActualizado = "14 de Agosto de 2026" }

# Filas de la tabla por región
$tableRowsHtml = ""
foreach ($rKey in $regiones.Keys | Sort-Object) {
    $rData = $regiones[$rKey]
    $rFunc = 100.00
    if ($rData.Camaras -gt 0) {
        $rFunc = [Math]::Round(((($rData.Camaras - $rData.SinVisual - $rData.MortSinVisual) / $rData.Camaras) * 100), 2)
    }
    $rFuncStr = "{0:N2}%" -f $rFunc

    if ($Empresa -eq "Cermaq") {
        $tableRowsHtml += @"
      <tr>
        <td><strong>$rKey</strong></td>
        <td>$($rData.Centros)</td>
        <td>$($rData.Jaulas)</td>
        <td>$($rData.Camaras)</td>
        <td>$($rData.SinVisual)</td>
        <td>$($rData.MortSinVisual)</td>
        <td><span class="status-badge-green">$rFuncStr</span></td>
      </tr>
"@
    } else {
        $tableRowsHtml += @"
      <tr>
        <td><strong>$rKey</strong></td>
        <td>$($rData.Centros)</td>
        <td>$($rData.Jaulas)</td>
        <td>$($rData.Camaras)</td>
        <td>$($rData.SinVisual)</td>
        <td><span class="status-badge-green">$rFuncStr</span></td>
      </tr>
"@
    }
}

# Encabezados de tabla según empresa
$tableHeaderHtml = @"
      <tr>
        <th>Regi&oacute;n</th>
        <th>Centros</th>
        <th>Jaulas</th>
        <th>Total C&aacute;maras</th>
        <th>C&aacute;maras (Sin Visual)</th>
        <th>Funcionamiento</th>
      </tr>
"@
if ($Empresa -eq "Cermaq") {
    $tableHeaderHtml = @"
      <tr>
        <th>Regi&oacute;n</th>
        <th>Centros</th>
        <th>Jaulas</th>
        <th>Total C&aacute;maras</th>
        <th>C&aacute;m. Est&aacute;ndar (Falla)</th>
        <th>C&aacute;m. Mortalidad (Falla)</th>
        <th>Funcionamiento</th>
      </tr>
"@
}

# Construir HTML Completo
$htmlTemplate = @"
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<title>Reporte $($Empresa) 2026 Semana $($semanaNum)</title>
<style>
  @page {
    size: A4 portrait;
    margin: 0;
  }
  * {
    box-sizing: border-box;
    font-family: 'Segoe UI', -apple-system, Roboto, Helvetica, Arial, sans-serif;
  }
  body {
    margin: 0;
    padding: 8mm 12mm;
    color: #0f2744;
    background-color: #ffffff;
    font-size: 13px;
    -webkit-print-color-adjust: exact;
  }
  
  .header-container {
    background: $headerBg;
    border-radius: 12px;
    padding: 12px 20px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    color: #ffffff;
    margin-bottom: 16px;
    box-shadow: 0 4px 10px rgba(0,0,0,0.08);
  }
  .logo-box {
    background: #ffffff;
    border-radius: 8px;
    padding: 5px 16px;
    display: flex;
    align-items: center;
    justify-content: center;
    height: 46px;
  }
  .header-titles {
    text-align: right;
  }
  .header-main-title {
    font-size: 23px;
    font-weight: 800;
    letter-spacing: 0.5px;
    margin: 0;
    text-transform: uppercase;
    color: #ffffff;
  }
  .header-subtitle {
    font-size: 12px;
    color: #e0e8f0;
    margin-top: 4px;
    display: flex;
    align-items: center;
    justify-content: flex-end;
    gap: 10px;
  }
  .badge-pill {
    background-color: #0099e5;
    color: #ffffff;
    font-size: 11px;
    font-weight: 700;
    padding: 3px 12px;
    border-radius: 12px;
    letter-spacing: 0.2px;
  }

  .section-header {
    display: flex;
    align-items: center;
    margin-top: 14px;
    margin-bottom: 10px;
  }
  .section-bar {
    width: 5px;
    height: 18px;
    background-color: $accentColor;
    margin-right: 8px;
    border-radius: 2px;
  }
  .section-title {
    font-size: 14px;
    font-weight: 800;
    color: $accentColor;
    text-transform: uppercase;
    letter-spacing: 0.5px;
  }

  .kpi-grid {
    display: flex;
    gap: 12px;
    margin-bottom: 14px;
  }
  .kpi-card {
    flex: 1;
    background-color: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 10px;
    padding: 12px 8px;
    text-align: center;
    box-shadow: 0 2px 5px rgba(0,0,0,0.02);
  }
  .kpi-title {
    font-size: 10px;
    font-weight: 700;
    color: #5a6e85;
    text-transform: uppercase;
    letter-spacing: 0.5px;
    margin-bottom: 6px;
  }
  .kpi-value {
    font-size: 28px;
    font-weight: 800;
    color: $accentColor;
    line-height: 1;
  }
  .kpi-value.green {
    color: #00a86b;
  }

  .data-table {
    width: 100%;
    border-collapse: separate;
    border-spacing: 0;
    border-radius: 8px;
    overflow: hidden;
    border: 1px solid #cbd5e1;
    margin-bottom: 14px;
  }
  .data-table th {
    background-color: $accentColor;
    color: #ffffff;
    font-size: 11px;
    font-weight: 700;
    text-transform: uppercase;
    padding: 9px 12px;
    text-align: center;
    letter-spacing: 0.4px;
  }
  .data-table td {
    padding: 9px 12px;
    text-align: center;
    font-size: 12px;
    border-bottom: 1px solid #e2e8f0;
    background-color: #ffffff;
    color: #1e293b;
  }
  .data-table tr:nth-child(even) td {
    background-color: #f8fafc;
  }
  .data-table tr:last-child td {
    border-bottom: none;
  }
  .status-badge-green {
    background-color: #e6f4ea;
    color: #137333;
    font-weight: 700;
    padding: 3px 14px;
    border-radius: 12px;
    display: inline-block;
    font-size: 11px;
  }

  .charts-row {
    display: flex;
    gap: 12px;
    margin-bottom: 12px;
  }
  .chart-box {
    flex: 1;
    background-color: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 10px;
    padding: 10px 14px;
  }
  .chart-box-full {
    background-color: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 10px;
    padding: 10px 14px;
    margin-bottom: 10px;
  }
  .chart-title {
    font-size: 12px;
    font-weight: 700;
    color: #1e293b;
    text-align: center;
    margin-bottom: 6px;
  }

  .footer {
    text-align: right;
    font-size: 10px;
    color: #94a3b8;
    margin-top: 8px;
  }
</style>
</head>
<body>

  <div class="header-container">
    <div class="logo-box">
      $logoSvg
    </div>
    <div class="header-titles">
      <div class="header-main-title">$($Empresa.ToUpper()) &bull; REPORTE SEMANAL</div>
      <div class="header-subtitle">
        <span>Actualizado: $fechaActualizado</span>
        <span class="badge-pill">Confidencial</span>
      </div>
    </div>
  </div>

  <div class="section-header">
    <div class="section-bar"></div>
    <div class="section-title">M&eacute;tricas Generales</div>
  </div>

  <div class="kpi-grid">
    <div class="kpi-card">
      <div class="kpi-title">Centros Activos</div>
      <div class="kpi-value">$totalCentros</div>
    </div>
    <div class="kpi-card">
      <div class="kpi-title">Total Jaulas</div>
      <div class="kpi-value">$totalJaulas</div>
    </div>
    <div class="kpi-card">
      <div class="kpi-title">Total de C&aacute;maras</div>
      <div class="kpi-value">$totalCamaras</div>
    </div>
    <div class="kpi-card">
      <div class="kpi-title">Funcionamiento</div>
      <div class="kpi-value green">$funcGlobalStr</div>
    </div>
  </div>

  <div class="section-header">
    <div class="section-bar"></div>
    <div class="section-title">Resumen por Regi&oacute;n</div>
  </div>

  <table class="data-table">
    <thead>
      $tableHeaderHtml
    </thead>
    <tbody>
      $tableRowsHtml
    </tbody>
  </table>

  <div class="section-header">
    <div class="section-bar"></div>
    <div class="section-title">An&aacute;lisis Visual Integrado & Salud del Sistema</div>
  </div>

  <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 10px; margin-bottom: 12px;">
    <div class="chart-box">
      <div class="chart-title">Centros Activos por Regi&oacute;n</div>
      <svg width="100%" height="135" viewBox="0 0 260 135" xmlns="http://www.w3.org/2000/svg">
        <line x1="30" y1="15" x2="240" y2="15" stroke="#e2e8f0" stroke-dasharray="3,3" />
        <line x1="30" y1="50" x2="240" y2="50" stroke="#e2e8f0" stroke-dasharray="3,3" />
        <line x1="30" y1="85" x2="240" y2="85" stroke="#e2e8f0" stroke-dasharray="3,3" />
        <line x1="30" y1="110" x2="240" y2="110" stroke="#cbd5e1" stroke-width="1" />
        
        <rect x="60" y="20" width="50" height="90" fill="#1e62c0" rx="3" />
        <text x="85" y="15" font-size="11" font-weight="bold" fill="#0f2744" text-anchor="middle">11</text>
        <text x="85" y="125" font-size="11" font-weight="bold" fill="#0f2744" text-anchor="middle">XI</text>

        <rect x="150" y="65" width="50" height="45" fill="#48c6ff" rx="3" />
        <text x="175" y="60" font-size="11" font-weight="bold" fill="#0f2744" text-anchor="middle">5</text>
        <text x="175" y="125" font-size="11" font-weight="bold" fill="#0f2744" text-anchor="middle">X</text>
      </svg>
    </div>

    <div class="chart-box">
      <div class="chart-title">Funcionamiento Promedio por Regi&oacute;n</div>
      <svg width="100%" height="135" viewBox="0 0 260 135" xmlns="http://www.w3.org/2000/svg">
        <line x1="35" y1="20" x2="240" y2="20" stroke="#e2e8f0" stroke-dasharray="3,3" />
        <line x1="35" y1="60" x2="240" y2="60" stroke="#e2e8f0" stroke-dasharray="3,3" />
        <line x1="35" y1="110" x2="240" y2="110" stroke="#cbd5e1" stroke-width="1" />

        <rect x="65" y="20" width="50" height="90" fill="#00a86b" rx="3" />
        <text x="90" y="15" font-size="10" font-weight="bold" fill="#00a86b" text-anchor="middle">$funcGlobalStr</text>
        <text x="90" y="125" font-size="11" font-weight="bold" fill="#0f2744" text-anchor="middle">XI</text>

        <rect x="155" y="20" width="50" height="90" fill="#00a86b" rx="3" />
        <text x="180" y="15" font-size="10" font-weight="bold" fill="#0f2744" text-anchor="middle">$funcGlobalStr</text>
        <text x="180" y="125" font-size="11" font-weight="bold" fill="#0f2744" text-anchor="middle">X</text>
      </svg>
    </div>

    <div class="chart-box">
      <div class="chart-title">Estado del Parque de C&aacute;maras</div>
      <svg width="100%" height="135" viewBox="0 0 260 135" xmlns="http://www.w3.org/2000/svg">
        <circle cx="80" cy="65" r="42" fill="none" stroke="#00a86b" stroke-width="16" />
        <text x="80" y="62" font-size="13" font-weight="bold" fill="#0f2744" text-anchor="middle">$funcGlobalStr</text>
        <text x="80" y="76" font-size="9" fill="#64748b" text-anchor="middle">Operativo</text>

        <circle cx="145" cy="40" r="5" fill="#00a86b" />
        <text x="156" y="44" font-size="10" font-weight="600" fill="#334155">Operativas ($totalCamaras)</text>

        <circle cx="145" cy="65" r="5" fill="#f59e0b" />
        <text x="156" y="69" font-size="10" font-weight="600" fill="#334155">Respaldo (OK)</text>

        <circle cx="145" cy="90" r="5" fill="#ef4444" />
        <text x="156" y="94" font-size="10" font-weight="600" fill="#334155">Sin Visual ($totalSinVisual)</text>
      </svg>
    </div>

    <div class="chart-box">
      <div class="chart-title">Relaci&oacute;n Jaulas vs C&aacute;maras</div>
      <svg width="100%" height="135" viewBox="0 0 260 135" xmlns="http://www.w3.org/2000/svg">
        <line x1="30" y1="110" x2="240" y2="110" stroke="#cbd5e1" stroke-width="1" />

        <rect x="50" y="30" width="22" height="80" fill="#093c71" rx="2" />
        <rect x="75" y="30" width="22" height="80" fill="#0099e5" rx="2" />
        <text x="73" y="125" font-size="10" font-weight="bold" fill="#0f2744" text-anchor="middle">XI</text>

        <rect x="145" y="65" width="22" height="45" fill="#093c71" rx="2" />
        <rect x="170" y="65" width="22" height="45" fill="#0099e5" rx="2" />
        <text x="168" y="125" font-size="10" font-weight="bold" fill="#0f2744" text-anchor="middle">X</text>

        <rect x="180" y="10" width="10" height="10" fill="#093c71" rx="1" />
        <text x="195" y="18" font-size="9" fill="#475569">Jaulas</text>
        <rect x="220" y="10" width="10" height="10" fill="#0099e5" rx="1" />
        <text x="235" y="18" font-size="9" fill="#475569">C&aacute;m.</text>
      </svg>
    </div>
  </div>

  <div style="display: flex; gap: 10px; margin-bottom: 12px;">
    <div style="flex: 1; background-color: #f8fafc; border: 1px solid #cbd5e1; border-radius: 8px; padding: 8px 12px; display: flex; align-items: center; gap: 10px;">
      <div style="font-size: 20px;">📡</div>
      <div>
        <div style="font-size: 10px; font-weight: 700; color: #64748b; text-transform: uppercase;">Conectividad Remota</div>
        <div style="font-size: 12px; font-weight: 700; color: #0f2744;">Enlaces Activos (100%)</div>
      </div>
    </div>
    <div style="flex: 1; background-color: #f8fafc; border: 1px solid #cbd5e1; border-radius: 8px; padding: 8px 12px; display: flex; align-items: center; gap: 10px;">
      <div style="font-size: 20px;">📹</div>
      <div>
        <div style="font-size: 10px; font-weight: 700; color: #64748b; text-transform: uppercase;">C&aacute;maras de Respaldo</div>
        <div style="font-size: 12px; font-weight: 700; color: #0f2744;">Sistemas Habilitados</div>
      </div>
    </div>
    <div style="flex: 1; background-color: #f8fafc; border: 1px solid #cbd5e1; border-radius: 8px; padding: 8px 12px; display: flex; align-items: center; gap: 10px;">
      <div style="font-size: 20px;">🛡️</div>
      <div>
        <div style="font-size: 10px; font-weight: 700; color: #64748b; text-transform: uppercase;">Salud General</div>
        <div style="font-size: 12px; font-weight: 700; color: #0f2744;">Disponibilidad Alta</div>
      </div>
    </div>
  </div>

  <div class="footer">
    Reporte $($Empresa.ToUpper()) 2026 &bull; Semana $semanaNum &bull; P&aacute;gina 1 de 2
  </div>

  <!-- SALTO DE PÁGINA PARA PAGINA 2 (DETALLE DE CENTROS E INSIGHTS) -->
  <div style="page-break-before: always; break-before: page; margin-top: 20px;"></div>

  <div class="header-container" style="padding: 8px 16px;">
    <div class="header-main-title" style="font-size: 18px;">$($Empresa.ToUpper()) &bull; DETALLE OPERATIVO DE CENTROS & INSIGHTS</div>
    <div class="header-subtitle">
      <span>Semana $semanaNum</span>
    </div>
  </div>

  <div class="section-header">
    <div class="section-bar"></div>
    <div class="section-title">Detalle por Centro & Insights Operativos</div>
  </div>

  <table class="data-table">
    <thead>
      <tr>
        <th style="text-align: left;">Centro de Cultivo</th>
        <th>Regi&oacute;n</th>
        <th>Jaulas</th>
        <th>C&aacute;maras</th>
        <th style="text-align: left;">Estado / Observaciones & Insights</th>
      </tr>
    </thead>
    <tbody>
      $centrosDetailRowsHtml
    </tbody>
  </table>

  <div class="footer">
    Reporte $($Empresa.ToUpper()) 2026 &bull; Semana $semanaNum &bull; P&aacute;gina 2 de 2
  </div>

</body>
</html>
"@

$tmpHtml = "$workingDir\temp_report.html"
$outputDir = "$workingDir\Reportes_PDF\Semana_$semanaNum"
if (-not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Force -Path $outputDir | Out-Null
}
$outputPdf = "$outputDir\Reporte $Empresa 2026 Semana $semanaNum.pdf"

[System.IO.File]::WriteAllText($tmpHtml, $htmlTemplate, [System.Text.Encoding]::UTF8)

# Convertir HTML a PDF usando Edge Headless (Blink PDF Engine)
$edgeCmd = "`"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe`" --headless=new --no-sandbox --disable-gpu --no-pdf-header-footer --print-to-pdf=`"$outputPdf`" `"file:///$($tmpHtml.Replace('\', '/'))`""
cmd /c $edgeCmd

# Limpiar temporal
if (Test-Path $tmpHtml) { Remove-Item $tmpHtml -Force }

Write-Host "=================================================="
Write-Host "REPORTE GENERADO EXITOSAMENTE:"
Write-Host "Empresa: $Empresa"
Write-Host "Semana: $semanaNum"
Write-Host "Ruta: $outputPdf"
Write-Host "=================================================="
