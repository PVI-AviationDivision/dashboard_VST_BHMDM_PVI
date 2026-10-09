# ------------------------------------------------------------------
#  Build data.js for the "BH Mot doi mot - Viettel Store" dashboard
#  - Reads every bhkh-*.xlsx in the project folder (no Excel/Python needed)
#  - Merges & de-duplicates by policy number (newest file wins)
#  - Masks customer names before writing data.js (safe to publish)
#  Compatible with Windows PowerShell 5.1 and PowerShell 7+.
#  File kept ASCII-only on purpose (PS 5.1 encoding safety).
# ------------------------------------------------------------------
param(
  [string]$Root = (Split-Path -Parent $PSScriptRoot),
  [string]$Pattern = 'bhkh-*.xlsx',
  [string]$DataDir = 'BaoCao'
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Normalize([string]$s) {
  if ($null -eq $s) { return '' }
  $s = $s.Trim().ToLowerInvariant()
  $s = $s.Replace([string][char]0x0111, 'd')   # d-stroke
  $f = $s.Normalize([Text.NormalizationForm]::FormD)
  $sb = New-Object Text.StringBuilder
  foreach ($ch in $f.ToCharArray()) {
    if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch) -ne [Globalization.UnicodeCategory]::NonSpacingMark) { [void]$sb.Append($ch) }
  }
  return (($sb.ToString()) -replace '\s+', ' ')
}

function ColIndex([string]$ref) {
  $letters = ($ref -replace '[0-9]', '').ToUpperInvariant()
  $n = 0
  foreach ($c in $letters.ToCharArray()) { $n = $n * 26 + ([int]$c - 64) }
  return $n - 1
}

function ReadEntry($zip, [string]$name) {
  $e = $zip.GetEntry($name)
  if ($null -eq $e) { return $null }
  $sr = New-Object IO.StreamReader($e.Open(), [Text.Encoding]::UTF8)
  try { return $sr.ReadToEnd() } finally { $sr.Close() }
}

function Read-XlsxRows([string]$path) {
  # copy first: the file may be open/locked in Excel
  $tmp = [IO.Path]::Combine([IO.Path]::GetTempPath(), [Guid]::NewGuid().ToString() + '.xlsx')
  $fsIn = New-Object IO.FileStream($path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
  $fsOut = [IO.File]::Create($tmp)
  try { $fsIn.CopyTo($fsOut) } finally { $fsIn.Close(); $fsOut.Close() }
  $zip = [IO.Compression.ZipFile]::OpenRead($tmp)
  try {
    $shared = @()
    $sstXml = ReadEntry $zip 'xl/sharedStrings.xml'
    if ($sstXml) {
      $sst = New-Object Xml.XmlDocument; $sst.LoadXml($sstXml)
      foreach ($si in $sst.DocumentElement.ChildNodes) {
        if ($si.LocalName -ne 'si') { continue }
        $txt = ''
        foreach ($t in $si.GetElementsByTagName('t', $si.NamespaceURI)) {
          if ($t.ParentNode.LocalName -ne 'rPh') { $txt += $t.InnerText }
        }
        $shared += ,$txt
      }
    }
    $sheetName = 'xl/worksheets/sheet1.xml'
    if ($null -eq $zip.GetEntry($sheetName)) {
      $sheetName = ($zip.Entries | Where-Object { $_.FullName -like 'xl/worksheets/*.xml' } | Select-Object -First 1).FullName
    }
    $doc = New-Object Xml.XmlDocument; $doc.LoadXml((ReadEntry $zip $sheetName))
    $rows = New-Object System.Collections.ArrayList
    foreach ($row in $doc.GetElementsByTagName('row', $doc.DocumentElement.NamespaceURI)) {
      $cells = @{}
      $max = -1
      foreach ($c in $row.ChildNodes) {
        if ($c.LocalName -ne 'c') { continue }
        $i = ColIndex $c.GetAttribute('r')
        $t = $c.GetAttribute('t')
        $val = ''
        if ($t -eq 'inlineStr') { $val = $c.InnerText }
        else {
          $v = $null
          foreach ($ch in $c.ChildNodes) { if ($ch.LocalName -eq 'v') { $v = $ch.InnerText } }
          if ($null -ne $v) { if ($t -eq 's') { $val = $shared[[int]$v] } else { $val = $v } }
        }
        $cells[$i] = $val
        if ($i -gt $max) { $max = $i }
      }
      $arr = New-Object string[] ($max + 1)
      foreach ($k in $cells.Keys) { $arr[$k] = $cells[$k] }
      [void]$rows.Add($arr)
    }
    return ,$rows
  } finally { $zip.Dispose(); Remove-Item $tmp -Force -ErrorAction SilentlyContinue }
}

function To-IsoDate([string]$s) {
  if ([string]::IsNullOrWhiteSpace($s)) { return $null }
  $s = $s.Trim()
  $d = [datetime]::MinValue
  $fmts = [string[]]@('dd/MM/yyyy', 'd/M/yyyy', 'dd/MM/yyyy HH:mm:ss', 'yyyy-MM-dd', 'yyyy-MM-ddTHH:mm:ss')
  if ([datetime]::TryParseExact($s, $fmts, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$d)) { return $d.ToString('yyyy-MM-dd') }
  $n = 0.0
  if ([double]::TryParse($s, [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$n) -and $n -gt 20000 -and $n -lt 80000) {
    return ([datetime]'1899-12-30').AddDays([math]::Floor($n)).ToString('yyyy-MM-dd')
  }
  return $null
}

function To-Number([string]$s) {
  if ([string]::IsNullOrWhiteSpace($s)) { return 0 }
  $clean = ($s -replace '[^0-9\.\-]', '')
  $n = 0.0
  if ([double]::TryParse($clean, [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$n)) { return $n }
  return 0
}

function Mask-Name([string]$s) {
  if ([string]::IsNullOrWhiteSpace($s)) { return '' }
  $parts = $s.Trim() -split '\s+'
  $out = @()
  foreach ($p in $parts) { if ($p.Length -gt 0) { $out += $p.Substring(0, 1).ToUpperInvariant() + '.' } }
  return ($out -join '')
}

function Short-Hash([string]$s) {
  $sha = [Security.Cryptography.SHA256]::Create()
  $bytes = $sha.ComputeHash([Text.Encoding]::UTF8.GetBytes('mdm-vst|' + (Normalize $s)))
  return (($bytes[0..4] | ForEach-Object { $_.ToString('x2') }) -join '')
}

# column keys -> normalized header text
$want = [ordered]@{
  soct = 'so ct'; ngay = 'ngay ct'; sodon = 'so don bh'; kh = 'khach hang'; masp = 'ma sp'
  stbh = 'so tien bh'; phi = 'phi bh'; bd = 'ngay bat dau'; kt = 'ngay ket thuc'
  cb = 'cb khai thac'; file = 'sl file dinh kem'; tt = 'trang thai'
}

# reports live in <Root>\BaoCao (files left in the root folder are still read)
$dataPath = Join-Path $Root $DataDir
if (-not (Test-Path $dataPath)) { New-Item -ItemType Directory -Path $dataPath | Out-Null }
$files = @(Get-ChildItem -Path $dataPath, $Root -Filter $Pattern -File | Where-Object { $_.Name -notlike '~$*' } | Sort-Object @{ Expression = { $m = [regex]::Match($_.Name, '\d{4}-\d{2}-\d{2}'); if ($m.Success) { $m.Value } else { $_.LastWriteTime.ToString('yyyy-MM-dd') } } }, LastWriteTime)
if ($files.Count -eq 0) { throw "Khong tim thay file $Pattern trong thu muc $DataDir" }

# reads column $k of current row $a using header map $map (dynamic scope)
function G($k) { $i = $map[$k]; if ($i -ge 0 -and $i -lt $a.Length) { return [string]$a[$i] } else { return '' } }

$byPolicy = [ordered]@{}
$sourceFiles = @()
$prevMax = ''   # latest Ngay CT already reported by earlier files
$prevMin = ''   # earliest Ngay CT covered by earlier files
foreach ($f in $files) {
  $fm = [regex]::Match($f.Name, '\d{4}-\d{2}-\d{2}')
  $fileDate = $(if ($fm.Success) { $fm.Value } else { $f.LastWriteTime.ToString('yyyy-MM-dd') })
  $fileMax = $prevMax
  $fileMin = $prevMin
  Write-Host ("  - Doc file: " + $f.Name)
  $rows = Read-XlsxRows $f.FullName
  $hdrIdx = -1; $map = @{}
  for ($r = 0; $r -lt [math]::Min(15, $rows.Count); $r++) {
    $norm = @($rows[$r] | ForEach-Object { Normalize $_ })
    if ($norm -contains 'so don bh') {
      $hdrIdx = $r
      foreach ($k in $want.Keys) { $map[$k] = [array]::IndexOf($norm, $want[$k]) }
      break
    }
  }
  if ($hdrIdx -lt 0) { Write-Warning ("    Bo qua (khong thay cot 'So don BH'): " + $f.Name); continue }
  $cnt = 0
  for ($r = $hdrIdx + 1; $r -lt $rows.Count; $r++) {
    $a = $rows[$r]
    $sodon = (G 'sodon').Trim()
    if ($sodon -eq '') { continue }
    $isAdj = $sodon -match '/E\d+'
    $base = ($sodon -split '/E\d+')[0].Trim()
    $sdbs = ''
    if ($isAdj) { $sdbs = ([regex]::Match($sodon, '/(E\d+.*)$')).Groups[1].Value.Trim() }
    $name = G 'kh'
    $rec = [ordered]@{
      ct    = (G 'soct').Trim()
      ngay  = To-IsoDate (G 'ngay')
      don   = $sodon
      goc   = $base
      loai  = $(if ($isAdj) { 'H' } else { 'G' })
      sdbs  = $sdbs
      kh    = Mask-Name $name
      kid   = Short-Hash $name
      sp    = (G 'masp').Trim()
      stbh  = To-Number (G 'stbh')
      phi   = To-Number (G 'phi')
      bd    = To-IsoDate (G 'bd')
      kt    = To-IsoDate (G 'kt')
      cb    = (G 'cb').Trim()
      tt    = (G 'tt').Trim()
      nap   = $fileDate
      lui   = $false
    }
    if ($byPolicy.Contains($sodon)) {
      # keep when the policy was first seen
      $rec.nap = $byPolicy[$sodon].nap
      $rec.lui = $byPolicy[$sodon].lui
    } elseif ($prevMax -ne '' -and $rec.ngay -and $rec.ngay -lt $prevMax -and $rec.ngay -ge $prevMin) {
      # new policy dated inside a period already reported -> back-dated entry
      # (dates before the earliest reported day just mean the export range was widened)
      $rec.lui = $true
    }
    if ($rec.ngay -and $rec.ngay -gt $fileMax) { $fileMax = $rec.ngay }
    if ($rec.ngay -and ($fileMin -eq '' -or $rec.ngay -lt $fileMin)) { $fileMin = $rec.ngay }
    $byPolicy[$sodon] = $rec
    $cnt++
  }
  Write-Host ("    -> " + $cnt + " dong")
  $sourceFiles += $f.Name
  $prevMax = $fileMax
  $prevMin = $fileMin
}

$list = @($byPolicy.Values)
$payload = [ordered]@{
  generatedAt = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss')
  sourceFiles = $sourceFiles
  rows        = $list
}
$json = $payload | ConvertTo-Json -Depth 5 -Compress
$js = "/* Tu dong sinh boi scripts/build-data.ps1 - khong sua tay */`nwindow.MDM_DATA = " + $json + ";`n"
$outPath = Join-Path $Root 'data.js'
[IO.File]::WriteAllText($outPath, $js, (New-Object Text.UTF8Encoding($false)))

$g = @($list | Where-Object { $_.loai -eq 'G' }).Count
$h = @($list | Where-Object { $_.loai -eq 'H' }).Count
Write-Host ""
Write-Host ("  Ghi nhan: " + $g + "  |  Huy (SDBS): " + $h + "  |  Thuc te: " + ($g - $h))
Write-Host ("  Da ghi: " + $outPath)
