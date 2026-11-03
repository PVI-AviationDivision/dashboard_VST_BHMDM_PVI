# Shared helpers: read .xlsx without Excel (used by build-vst.ps1). ASCII-only.
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
