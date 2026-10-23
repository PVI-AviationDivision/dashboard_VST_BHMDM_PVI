# ------------------------------------------------------------------
#  Tạo tin nhắn báo cáo ngắn từ data.js và copy vào clipboard (dán Zalo)
#  File lưu UTF-8 có BOM để Windows PowerShell 5.1 đọc đúng tiếng Việt.
# ------------------------------------------------------------------
param([string]$Root = (Split-Path -Parent $PSScriptRoot))
$ErrorActionPreference = 'Stop'

$Link = 'https://pvi-aviationdivision.github.io/dashboard_VST_BHMDM_PVI/'

$store = Join-Path $Root 'BaoCao\_tonghop.csv'
if (Test-Path $store) {
  # đọc file tổng hợp CSV (nhanh, không giới hạn dung lượng như ConvertFrom-Json của PS 5.1)
  $rows = @(Import-Csv -Path $store -Encoding UTF8)
} else {
  $raw = [IO.File]::ReadAllText((Join-Path $Root 'data.js'), [Text.Encoding]::UTF8)
  $json = $raw.Substring($raw.IndexOf('{'), $raw.LastIndexOf('}') - $raw.IndexOf('{') + 1)
  $rows = @(($json | ConvertFrom-Json).rows)
}

$vi = [Globalization.CultureInfo]::GetCultureInfo('vi-VN')
function N($x) { return ([double]$x).ToString('N0', $vi) }
function P($a, $b) { if ($b -eq 0) { return '0' } ; return ([double]$a * 100 / $b).ToString('0.0', $vi) }
function DMY([string]$s) { $p = $s.Split('-'); return $p[2] + '/' + $p[1] + '/' + $p[0] }

$GList = @($rows | Where-Object { $_.loai -eq 'G' })
# một đơn hủy cho mỗi đơn gốc (lấy SĐBS sớm nhất)
$origBy = @{}; foreach ($og in $GList) { $origBy[$og.goc] = $og }
$cancelBy = @{}
foreach ($h in @($rows | Where-Object { $_.loai -eq 'H' } | Sort-Object ngay)) { if (-not $cancelBy.ContainsKey($h.goc)) { $cancelBy[$h.goc] = $h } }
$CList = @($cancelBy.Values)

$last = ($rows | ForEach-Object { $_.ngay } | Where-Object { $_ } | Sort-Object | Select-Object -Last 1)
$month = $last.Substring(0, 7)
$first = ($rows | ForEach-Object { $_.ngay } | Where-Object { $_ } | Sort-Object | Select-Object -First 1)

$dG = @($GList | Where-Object { $_.ngay -eq $last }).Count
$dH = @($CList | Where-Object { $_.ngay -eq $last }).Count
$mG = @($GList | Where-Object { $_.ngay.StartsWith($month) }).Count
$mH = @($CList | Where-Object { $_.ngay.StartsWith($month) }).Count
$tG = $GList.Count; $tH = $CList.Count

$phiG = 0; foreach ($og in $GList) { $phiG += [double]$og.phi }
$phiH = 0; foreach ($cx in $CList) { if ($origBy.ContainsKey($cx.goc)) { $phiH += [double]$origBy[$cx.goc].phi } else { $phiH += [double]$cx.phi } }
$mm = $month.Split('-'); $monthLabel = $mm[1] + '/' + $mm[0]

$msg = @(
  "📊 BÁO CÁO BH MỘT ĐỔI MỘT – VIETTEL STORE",
  "Số liệu đến ngày $(DMY $last)",
  "",
  "▪️ Ngày $(DMY $last): +$(N $dG) đơn ghi nhận, $(N $dH) đơn hủy",
  "▪️ Tháng ${monthLabel}: $(N $mG) ghi nhận – $(N $mH) hủy – $(N ($mG - $mH)) đơn thực tế",
  "▪️ Lũy kế từ $(DMY $first): $(N $tG) ghi nhận – $(N $tH) hủy – $(N ($tG - $tH)) đơn thực tế (tỷ lệ hủy $(P $tH $tG)%)",
  "▪️ Phí BH thực tế lũy kế: $(N ($phiG - $phiH)) đ",
  "",
  "👉 Xem chi tiết: $Link"
) -join "`r`n"

Write-Host ""
Write-Host "  ---------------- TIN NHẮN ----------------"
Write-Host $msg
Write-Host "  ------------------------------------------"
try {
  Set-Clipboard -Value $msg
  Write-Host "  Đã copy vào clipboard - mở Zalo và bấm Ctrl+V để dán."
} catch {
  try { Add-Type -AssemblyName System.Windows.Forms; [Windows.Forms.Clipboard]::SetText($msg); Write-Host "  Đã copy vào clipboard - mở Zalo và bấm Ctrl+V để dán." }
  catch { Write-Host "  Không copy được vào clipboard, hãy bôi đen đoạn trên để copy." }
}
