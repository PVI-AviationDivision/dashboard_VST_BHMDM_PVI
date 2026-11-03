# ------------------------------------------------------------------
#  Build vst.js from Viettel Store monthly reports ("Quan ly ho so bao hiem PVI")
#  - Reads every .xlsx in BaoCao\VST (monthly files are small: read all, dedupe by Ma ho so)
#  - Keeps NO customer data (name, phone, address, IMEI are dropped)
#  - Output: per-record compact rows [date, hour, store, staff, status, apiError]
#  ASCII-only file (Windows PowerShell 5.1 safe).
# ------------------------------------------------------------------
param([string]$Root = (Split-Path -Parent $PSScriptRoot), [string]$DataDir = 'BaoCao\VST')
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'xlsx-lib.ps1')

$vstPath = Join-Path $Root $DataDir
if (-not (Test-Path $vstPath)) { New-Item -ItemType Directory -Path $vstPath | Out-Null }
$files = @(Get-ChildItem -Path $vstPath -Filter '*.xlsx' -File | Where-Object { $_.Name -notlike '~$*' -and $_.Name -notlike 'danh-muc*' } | Sort-Object LastWriteTime)

# ---- device catalogue: BaoCao\VST\danh-muc-may.xlsx (Ma thiet bi -> Dong may, Dung luong) ----
$catalog = @{}
$catPath = Join-Path $vstPath 'danh-muc-may.xlsx'
if (Test-Path $catPath) {
  $crows = Read-XlsxRows $catPath
  if ($crows.Count -gt 0) {
    $ch = @($crows[0] | ForEach-Object { Normalize $_ })
    $iS = [array]::IndexOf($ch, 'ma thiet bi'); $iM = [array]::IndexOf($ch, 'dong may'); $iC = [array]::IndexOf($ch, 'dung luong')
    for ($r = 1; $r -lt $crows.Count; $r++) {
      $x = $crows[$r]; if ($iS -lt 0 -or $iS -ge $x.Length -or [string]::IsNullOrWhiteSpace($x[$iS])) { continue }
      $sku = ([string]$x[$iS]).Trim()
      $m = $(if ($iM -ge 0 -and $iM -lt $x.Length) { ([string]$x[$iM]).Trim() } else { '' })
      $c = $(if ($iC -ge 0 -and $iC -lt $x.Length) { ([string]$x[$iC]).Trim() } else { '' })
      $catalog[$sku] = @{ m = $m; c = $c }
    }
  }
  Write-Host ("  Danh muc may: " + $catalog.Count + " ma thiet bi")
}
if ($files.Count -eq 0) { Write-Host "  Chua co file bao cao VST trong $DataDir - bo qua."; exit 0 }

function To-DateTime([string]$s) {
  if ([string]::IsNullOrWhiteSpace($s)) { return $null }
  $s = $s.Trim(); $d = [datetime]::MinValue
  $fmts = [string[]]@('dd/MM/yyyy HH:mm:ss', 'dd/MM/yyyy H:mm:ss', 'dd/MM/yyyy HH:mm', 'dd/MM/yyyy', 'd/M/yyyy H:mm:ss', 'yyyy-MM-dd HH:mm:ss', 'yyyy-MM-ddTHH:mm:ss')
  if ([datetime]::TryParseExact($s, $fmts, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$d)) { return $d }
  $n = 0.0
  if ([double]::TryParse($s, [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$n) -and $n -gt 20000 -and $n -lt 80000) {
    return ([datetime]'1899-12-30').AddDays($n)
  }
  return $null
}
function JS([string]$v) {
  if ($null -eq $v) { return '""' }
  return '"' + ($v.Replace('\', '\\').Replace('"', '\"') -replace "[\x00-\x1f]", ' ') + '"'
}

$want = [ordered]@{ id = 'ma ho so'; imei = 'imei thiet bi'; gd = 'ma giao dich'; sku = 'ma thiet bi'; gia = 'gia tri thiet bi'; phi = 'phi bao hiem'; tt = 'trang thai'; ngay = 'ngay tao ho so'; kho = 'kho'; nv = 'nguoi tao ho so'; api = 'trang thai api'; msg = 'mo ta trang thai api' }
$recs = [ordered]@{}
$stores = @{}; $provs = @{}; $skus = @{}
foreach ($f in $files) {
  Write-Host ("  - Doc file VST: " + $f.Name)
  $rows = Read-XlsxRows $f.FullName
  $hdr = -1; $map = @{}
  for ($r = 0; $r -lt [math]::Min(15, $rows.Count); $r++) {
    $norm = @($rows[$r] | ForEach-Object { Normalize $_ })
    if ($norm -contains 'ma ho so' -and $norm -contains 'kho') { $hdr = $r; foreach ($k in $want.Keys) { $map[$k] = [array]::IndexOf($norm, $want[$k]) }; break }
  }
  if ($hdr -lt 0) { Write-Warning "    Bo qua (khong thay cot 'Ma ho so' / 'Kho')"; continue }
  $cnt = 0
  for ($r = $hdr + 1; $r -lt $rows.Count; $r++) {
    $a = $rows[$r]
    function V($k) { $i = $map[$k]; if ($i -ge 0 -and $i -lt $a.Length -and $null -ne $a[$i]) { return ([string]$a[$i]).Trim() } else { return '' } }
    $id = V 'id'; if ($id -eq '') { continue }
    # same device can appear in several monthly files with different 'Ma ho so' -> dedupe by IMEI
    $key = V 'imei'; if ($key -eq '') { $key = V 'gd' }; if ($key -eq '') { $key = 'ID' + $id }
    $dt = To-DateTime (V 'ngay'); if ($null -eq $dt) { continue }
    $kho = V 'kho'
    # "333294 - HNI01 - HNI, P, So 45A Hang Dau, Phuong ..., Thanh pho Ha Noi"  /  "329874 - DNG-CH01 - DNG , P, ..."
    $parts = $kho -split '\s+-\s+|\s+\u2013\s+', 3
    $st = ([regex]::Match($kho, '\b([A-Z]{3}(?:-CH)?\d{2,3})\b')).Groups[1].Value
    if ($st -eq '') { $st = $(if ($parts.Count -ge 2) { $parts[1].Trim() } else { 'KHAC' }) }
    $pv = ([regex]::Match($st, '^[A-Z]+')).Value
    if (-not $stores.ContainsKey($st)) {
      $addr = ''
      if ($parts.Count -ge 3) {
        $seg = @($parts[2] -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' })
        if ($seg.Count -ge 3) { $addr = $seg[2] }              # street, e.g. "So 45A Hang Dau"
        $last = $seg[$seg.Count - 1]
        if ($pv -and -not $provs.ContainsKey($pv) -and $last -match '^(Th\S+nh\s+ph\S+|T\S+nh|TP\.?)\s+(.+)$') { $provs[$pv] = $Matches[2].Trim() }
        if ($seg.Count -ge 2) { $ward = $seg[$seg.Count - 2]; if ($ward -notmatch '^(P|X)$') { $addr = ($addr + ', ' + $ward).Trim(', ') } }
      }
      $stores[$st] = @{ pv = $pv; ad = $addr }
    }
    $tt = Normalize (V 'tt')
    $s = $(if ($tt -like 'hieu luc*') { 'L' } elseif ($tt -like '*huy*') { 'H' } elseif ($tt -like 'khoi tao*') { 'K' } else { 'X' })
    $err = ''
    if ((Normalize (V 'api')) -notlike 'thanh cong*') { $err = V 'msg'; if ($err -eq '') { $err = 'Loi khong ro' } }
    $sku = V 'sku'
    $gia = To-Number (V 'gia'); if ($gia -le 0 -or $gia -eq (To-Number (V 'phi'))) { $gia = 0 }   # some VST rows carry the premium instead of the device value
    if ($sku -ne '') {
      if (-not $skus.ContainsKey($sku)) { $skus[$sku] = @{ g = 0 } }
      if ($gia -gt 0) { $skus[$sku].g = $gia }
    }
    $recs[$key] = @($dt.ToString('yyyy-MM-dd'), $dt.Hour, $st, (V 'nv'), $s, $err, $sku)
    $cnt++
  }
  Write-Host ("    -> " + $cnt + " ho so")
}

$sb = New-Object Text.StringBuilder
[void]$sb.Append("/* Tu dong sinh boi scripts/build-vst.ps1 - khong chua thong tin khach hang */`nwindow.VST_DATA = {")
[void]$sb.Append('"generatedAt":' + (JS (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss')) + ',"sourceFiles":[' + (($files | ForEach-Object { JS $_.Name }) -join ',') + '],')
[void]$sb.Append('"provinces":{' + (($provs.Keys | Sort-Object | ForEach-Object { (JS $_) + ':' + (JS $provs[$_]) }) -join ',') + '},')
[void]$sb.Append('"stores":{' + (($stores.Keys | Sort-Object | ForEach-Object { (JS $_) + ':{"pv":' + (JS $stores[$_].pv) + ',"ad":' + (JS $stores[$_].ad) + '}' }) -join ',') + '},')
[void]$sb.Append('"models":{' + (($skus.Keys | Sort-Object | ForEach-Object { $cm = $catalog[$_]; (JS $_) + ':{"m":' + (JS $(if ($cm) { $cm.m } else { '' })) + ',"c":' + (JS $(if ($cm) { $cm.c } else { '' })) + ',"g":' + ([double]$skus[$_].g).ToString('0', [Globalization.CultureInfo]::InvariantCulture) + '}' }) -join ',') + '},')
[void]$sb.Append('"cols":["ngay","gio","ch","nv","tt","loi","sku"],"rows":[')
$first = $true
foreach ($r in $recs.Values) {
  if (-not $first) { [void]$sb.Append(',') }; $first = $false
  [void]$sb.Append('[' + (JS $r[0]) + ',' + $r[1] + ',' + (JS $r[2]) + ',' + (JS $r[3]) + ',' + (JS $r[4]) + ',' + (JS $r[5]) + ',' + (JS $r[6]) + ']')
}
[void]$sb.Append("]};`n")
$out = Join-Path $Root 'vst.js'
[IO.File]::WriteAllText($out, $sb.ToString(), (New-Object Text.UTF8Encoding($false)))
$L = @($recs.Values | Where-Object { $_[4] -eq 'L' }).Count; $H = @($recs.Values | Where-Object { $_[4] -eq 'H' }).Count; $E = @($recs.Values | Where-Object { $_[5] -ne '' }).Count
Write-Host ("  VST: " + $recs.Count + " ho so | hieu luc " + $L + " | da huy " + $H + " | loi API " + $E + " | " + $stores.Count + " cua hang")
$missing = @($skus.Keys | Where-Object { -not $catalog.ContainsKey($_) -or $catalog[$_].m -eq '' } | Sort-Object)
if ($missing.Count -gt 0) { Write-Host ("  Luu y: " + $missing.Count + " ma thiet bi chua co ten dong may trong danh-muc-may.xlsx: " + (($missing | Select-Object -First 8) -join ', ') + $(if ($missing.Count -gt 8) { ' ...' } else { '' })) }
Write-Host ("  Da ghi: " + $out)
