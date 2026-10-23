# ------------------------------------------------------------------
#  Build data.js for the "BH Mot doi mot - Viettel Store" dashboard
#  - Incremental: keeps a running store (BaoCao\_tonghop.csv); each run reads
#    only NEW bhkh-*.xlsx files in BaoCao, then moves them to BaoCao\DaNap
#  - Merges & de-duplicates by policy number (newest file wins)
#  - Use -Rebuild to re-read everything (BaoCao + BaoCao\DaNap) from scratch
#  - Masks customer names before writing data.js (safe to publish)
#  Compatible with Windows PowerShell 5.1 and PowerShell 7+.
#  File kept ASCII-only on purpose (PS 5.1 encoding safety).
# ------------------------------------------------------------------
param(
  [string]$Root = (Split-Path -Parent $PSScriptRoot),
  [string]$Pattern = 'bhkh-*.xlsx',
  [string]$DataDir = 'BaoCao',
  # Chi lay san pham BH Mot doi mot (Ma SP 020413). File xuat "BH Khac" co the lan san pham khac.
  [string[]]$Products = @('020413'),
  # Ngay bat dau tinh (yyyy-MM-dd). De trong = lay tat ca.
  [string]$FromDate = '',
  # Doc lai toan bo file (ke ca BaoCao\DaNap) va tao lai file tong hop
  [switch]$Rebuild
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

# ---------------- paths ----------------
$dataPath  = Join-Path $Root $DataDir
$donePath  = Join-Path $dataPath 'DaNap'
$storePath = Join-Path $dataPath '_tonghop.csv'
foreach ($d in @($dataPath, $donePath)) { if (-not (Test-Path $d)) { New-Item -ItemType Directory -Path $d | Out-Null } }

$cols = @('ct','ngay','don','goc','loai','sdbs','kh','kid','sp','stbh','phi','bd','kt','cb','tt','nap','lui')
$byPolicy = New-Object System.Collections.Specialized.OrderedDictionary
$prevMax = ''; $prevMin = ''
$fullRun = $Rebuild -or -not (Test-Path $storePath)

if (-not $fullRun) {
  foreach ($r in (Import-Csv -Path $storePath -Encoding UTF8)) {
    $rec = [ordered]@{}
    foreach ($c in $cols) { $rec[$c] = $r.$c }
    $rec.stbh = [double]$rec.stbh; $rec.phi = [double]$rec.phi; $rec.lui = ($rec.lui -eq 'True')
    $byPolicy[$rec.don] = $rec
    if ($rec.ngay) {
      if ($prevMax -eq '' -or $rec.ngay -gt $prevMax) { $prevMax = $rec.ngay }
      if ($prevMin -eq '' -or $rec.ngay -lt $prevMin) { $prevMin = $rec.ngay }
    }
  }
  Write-Host ("  Da co san " + $byPolicy.Count + " dong trong file tong hop.")
}

$dirs = @($dataPath, $Root)
if ($fullRun) { $dirs = @($donePath, $dataPath, $Root); Write-Host "  Che do doc lai toan bo file." }
$files = @(Get-ChildItem -Path $dirs -Filter $Pattern -File | Where-Object { $_.Name -notlike '~$*' } |
  Sort-Object @{ Expression = { $m = [regex]::Match($_.Name, '\d{4}-\d{2}-\d{2}'); if ($m.Success) { $m.Value } else { $_.LastWriteTime.ToString('yyyy-MM-dd') } } }, LastWriteTime)

if ($files.Count -eq 0 -and $byPolicy.Count -eq 0) { throw "Khong tim thay file $Pattern trong thu muc $DataDir" }
if ($files.Count -eq 0) { Write-Host "  Khong co file bao cao moi trong thu muc $DataDir - giu nguyen so lieu." }

# reads column $k of current row $a using header map $map (dynamic scope)
function G($k) { $i = $map[$k]; if ($i -ge 0 -and $i -lt $a.Length) { return [string]$a[$i] } else { return '' } }

$skipProd = 0; $skipDate = 0
$readFiles = @()
foreach ($f in $files) {
  Write-Host ("  - Doc file: " + $f.Name)
  $fm = [regex]::Match($f.Name, '\d{4}-\d{2}-\d{2}')
  $fileDate = $(if ($fm.Success) { $fm.Value } else { $f.LastWriteTime.ToString('yyyy-MM-dd') })
  $fileMax = $prevMax; $fileMin = $prevMin
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
  $cnt = 0; $new = 0
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
      ct = (G 'soct').Trim(); ngay = To-IsoDate (G 'ngay'); don = $sodon; goc = $base
      loai = $(if ($isAdj) { 'H' } else { 'G' }); sdbs = $sdbs
      kh = Mask-Name $name; kid = Short-Hash $name; sp = (G 'masp').Trim()
      stbh = To-Number (G 'stbh'); phi = To-Number (G 'phi')
      bd = To-IsoDate (G 'bd'); kt = To-IsoDate (G 'kt'); cb = (G 'cb').Trim(); tt = (G 'tt').Trim()
      nap = $fileDate; lui = $false
    }
    if ($Products.Count -gt 0 -and $Products -notcontains $rec.sp) { $skipProd++; continue }
    if ($FromDate -ne '' -and $rec.ngay -and $rec.ngay -lt $FromDate) { $skipDate++; continue }
    if ($byPolicy.Contains($sodon)) {
      $rec.nap = $byPolicy[$sodon].nap; $rec.lui = $byPolicy[$sodon].lui
    } else {
      $new++
      if ($prevMax -ne '' -and $rec.ngay -and $rec.ngay -lt $prevMax -and $rec.ngay -ge $prevMin) { $rec.lui = $true }
    }
    if ($rec.ngay -and $rec.ngay -gt $fileMax) { $fileMax = $rec.ngay }
    if ($rec.ngay -and ($fileMin -eq '' -or $rec.ngay -lt $fileMin)) { $fileMin = $rec.ngay }
    $byPolicy[$sodon] = $rec
    $cnt++
  }
  Write-Host ("    -> " + $cnt + " dong (" + $new + " don moi)")
  $readFiles += $f
  $prevMax = $fileMax; $prevMin = $fileMin
}

# ---------------- write store (CSV) ----------------
$list = @($byPolicy.Values)
$list | ForEach-Object { New-Object PSObject -Property $_ } | Select-Object $cols |
  Export-Csv -Path $storePath -NoTypeInformation -Encoding UTF8

# ---------------- write data.js (hand-built JSON: no size limit in PS 5.1) ----------------
function JS([string]$v) {
  if ($null -eq $v) { return 'null' }
  if ($v -notmatch '["\\\x00-\x1f]') { return '"' + $v + '"' }   # fast path
  $sb = New-Object Text.StringBuilder
  [void]$sb.Append('"')
  foreach ($ch in $v.ToCharArray()) {
    switch ($ch) {
      '"' { [void]$sb.Append('\"') } '\' { [void]$sb.Append('\\') }
      "`n" { [void]$sb.Append('\n') } "`r" { [void]$sb.Append('\r') } "`t" { [void]$sb.Append('\t') }
      default { if ([int]$ch -lt 32) { [void]$sb.Append(('\u{0:x4}' -f [int]$ch)) } else { [void]$sb.Append($ch) } }
    }
  }
  [void]$sb.Append('"'); return $sb.ToString()
}
$inv = [Globalization.CultureInfo]::InvariantCulture
$all = @(Get-ChildItem -Path @($donePath, $dataPath, $Root) -Filter $Pattern -File | Where-Object { $_.Name -notlike '~$*' } | Sort-Object Name | ForEach-Object { $_.Name })
$out = New-Object Text.StringBuilder
[void]$out.Append("/* Tu dong sinh boi scripts/build-data.ps1 - khong sua tay */`nwindow.MDM_DATA = {")
[void]$out.Append('"generatedAt":' + (JS (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss')) + ',"sourceFiles":[' + (($all | ForEach-Object { JS $_ }) -join ',') + '],"rows":[')
$first = $true
foreach ($r in $list) {
  if (-not $first) { [void]$out.Append(',') }; $first = $false
  [void]$out.Append('{"ct":' + (JS $r.ct) + ',"ngay":' + (JS $r.ngay) + ',"don":' + (JS $r.don) + ',"goc":' + (JS $r.goc) +
    ',"loai":' + (JS $r.loai) + ',"sdbs":' + (JS $r.sdbs) + ',"kh":' + (JS $r.kh) + ',"kid":' + (JS $r.kid) + ',"sp":' + (JS $r.sp) +
    ',"stbh":' + ([double]$r.stbh).ToString('0.##', $inv) + ',"phi":' + ([double]$r.phi).ToString('0.##', $inv) +
    ',"bd":' + (JS $r.bd) + ',"kt":' + (JS $r.kt) + ',"cb":' + (JS $r.cb) + ',"tt":' + (JS $r.tt) + ',"nap":' + (JS $r.nap) +
    ',"lui":' + $(if ($r.lui) { 'true' } else { 'false' }) + '}')
}
[void]$out.Append("]};`n")
$outPath = Join-Path $Root 'data.js'
[IO.File]::WriteAllText($outPath, $out.ToString(), (New-Object Text.UTF8Encoding($false)))

# ---------------- move processed files to BaoCao\DaNap ----------------
foreach ($f in $readFiles) {
  if ($f.DirectoryName -eq $donePath) { continue }
  $dest = Join-Path $donePath $f.Name
  if (Test-Path $dest) { $dest = Join-Path $donePath ($f.BaseName + '_' + (Get-Date).ToString('HHmmss') + $f.Extension) }
  try { Move-Item -LiteralPath $f.FullName -Destination $dest -ErrorAction Stop }
  catch { Write-Warning ("    Khong chuyen duoc " + $f.Name + " vao DaNap (file dang mo?). Lan sau se doc lai, khong anh huong so lieu.") }
}

$g = @($list | Where-Object { $_.loai -eq 'G' }).Count
$h = @($list | Where-Object { $_.loai -eq 'H' }).Count
Write-Host ""
if ($skipProd -gt 0) { Write-Host ("  Bo qua " + $skipProd + " dong khong phai san pham Mot doi mot (chi lay Ma SP " + ($Products -join ',') + ")") }
if ($skipDate -gt 0) { Write-Host ("  Bo qua " + $skipDate + " dong truoc ngay " + $FromDate) }
Write-Host ("  Ghi nhan: " + $g + "  |  Huy (SDBS): " + $h + "  |  Thuc te: " + ($g - $h))
Write-Host ("  Da ghi: " + $outPath)
