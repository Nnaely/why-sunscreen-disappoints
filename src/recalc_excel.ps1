param(
  [string]$Path = (Join-Path $PSScriptRoot "..\Portofolio_Ulasan_Skincare_Sephora.xlsx"),
  [string]$Pdf = "",
  [string]$Skin = "",
  [string]$Tone = ""
)
$Path = (Resolve-Path $Path).Path
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
  $wb = $xl.Workbooks.Open($Path)
  $dash = $wb.Worksheets.Item("Dashboard")
  $testMode = ($Skin -ne "" -or $Tone -ne "")
  if ($Skin -ne "") { $dash.Range("B8").Value2 = $Skin }
  if ($Tone -ne "") { $dash.Range("E8").Value2 = $Tone }
  $xl.CalculateFull()

  # #N/A di tabel tahunan Analisis sengaja dipakai agar titik tidak digambar; dihitung terpisah.
  $errors = @(); $na = 0; $formulas = 0
  foreach ($ws in $wb.Worksheets) {
    $used = $ws.UsedRange
    try { $cells = $used.SpecialCells(-4123) } catch { continue }   # xlCellTypeFormulas
    foreach ($c in $cells) {
      $formulas++
      $t = $c.Text
      if ($t -eq "#N/A" -and $ws.Name -eq "Analisis") { $na++; continue }
      if ($c.Value2 -is [int] -and $c.Value2 -lt -2146820000) { $errors += "$($ws.Name)!$($c.Address(0,0)) = $t" }
      elseif ($t -match '^#(REF!|NAME\?|VALUE!|DIV/0!|NUM!|NULL!|N/A)') { $errors += "$($ws.Name)!$($c.Address(0,0)) = $t" }
    }
  }
  $k = [ordered]@{}
  foreach ($a in "L8","B13","D21","B25","D33","E25","G33","H25","J33","B40","D51") { $k[$a] = $dash.Range($a).Text }

  if ($Pdf -ne "") { $dash.ExportAsFixedFormat(0, $Pdf) }
  if (-not $testMode) { $wb.Save() }
  $wb.Close($false)
  [pscustomobject]@{ formulas = $formulas; errors = $errors.Count; intended_na = $na; error_cells = ($errors | Select-Object -First 30); kpi = $k } | ConvertTo-Json -Depth 4
}
finally {
  $xl.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
