# Cap nhat bao cao chuoi thep len GitHub (chay boi auto_push.bat / lich hang ngay)
# 1) Don file cu: moi thu muc 1_CUTTING, 2_TON_KHO chi giu 1 file moi nhat (theo ngay trong ten file),
#    file cu chuyen sang ..\LUU_TRU_FILE_CU (khong xoa)
# 2) Day len GitHub -> GitHub tu chay bao cao
# 3) Neu lich su git qua lon (> 600 MB) thi gom gon lich su de repo khong bi day
$ErrorActionPreference = 'Continue'
$root = $PSScriptRoot
$arch = Join-Path (Split-Path $root -Parent) 'LUU_TRU_FILE_CU'
$logf = Join-Path $root 'nhat_ky_day_len.txt'
function Log($m) { $l = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') + '  ' + $m; Add-Content -Path $logf -Value $l -Encoding UTF8; Write-Host $l }
function NgayTrongTen($name) {
  $m = [regex]::Matches($name, '(?<!\d)(\d{1,2})[._-](\d{1,2})(?:[._-](\d{2,4}))?(?!\d)')
  foreach ($x in $m) {
    $y = if ($x.Groups[3].Success) { [int]$x.Groups[3].Value } else { (Get-Date).Year }
    if ($y -lt 100) { $y += 2000 }
    try { return (Get-Date -Year $y -Month ([int]$x.Groups[2].Value) -Day ([int]$x.Groups[1].Value)).Date } catch { }
  }
  return [datetime]::MinValue
}
Set-Location $root
Log '=== Bat dau ==='
foreach ($f in '1_CUTTING', '2_TON_KHO') {
  $dir = Join-Path $root "du_lieu\$f"
  $files = @(Get-ChildItem $dir -File | Where-Object { $_.Extension -match '^\.xls' -and $_.Name -notlike '~$*' })
  if ($files.Count -gt 1) {
    $sorted = @($files | Sort-Object @{ Expression = { NgayTrongTen $_.Name } }, LastWriteTime)
    $keep = $sorted[-1]
    New-Item -ItemType Directory -Force (Join-Path $arch $f) | Out-Null
    foreach ($o in $sorted[0..($sorted.Count - 2)]) {
      try { Move-Item $o.FullName (Join-Path $arch $f) -Force -ErrorAction Stop; Log "Chuyen file cu: $f\$($o.Name)" }
      catch { Log "KHONG chuyen duoc (file dang mo?): $($o.Name)" }
    }
    Log "Giu file moi nhat: $f\$($keep.Name)"
  }
}
git add -A 2>&1 | Out-Null
git diff --cached --quiet
if ($LASTEXITCODE -ne 0) { git commit -m ("Cap nhat du lieu " + (Get-Date -Format 'dd/MM/yyyy HH:mm')) 2>&1 | ForEach-Object { Log $_ } }
else { Log 'Khong co file moi' }
git pull --rebase --autostash origin main 2>&1 | ForEach-Object { Log $_ }
git push origin main 2>&1 | ForEach-Object { Log $_ }
# Kiem tra dung luong lich su
$kb = 0
git count-objects -v | ForEach-Object { if ($_ -match '^size-pack:\s*(\d+)') { $kb += [int]$Matches[1] }; if ($_ -match '^size:\s*(\d+)') { $kb += [int]$Matches[1] } }
Log ("Dung luong lich su git: {0:N0} MB" -f ($kb / 1024))
if ($kb -gt 600 * 1024) {
  Log 'Lich su > 600 MB: gom gon (giu nguyen toan bo file hien tai, bao cao cu trong lich_su/ van con)'
  git checkout --orphan gom_gon 2>&1 | Out-Null
  git add -A 2>&1 | Out-Null
  git commit -m ("Gom gon lich su " + (Get-Date -Format 'dd/MM/yyyy')) 2>&1 | Out-Null
  git branch -D main 2>&1 | Out-Null
  git branch -m main 2>&1 | Out-Null
  git push -f origin main 2>&1 | ForEach-Object { Log $_ }
  git reflog expire --expire=now --all 2>&1 | Out-Null
  git gc --prune=now 2>&1 | Out-Null
  Log 'Da gom gon xong'
}
Log '=== Xong ==='
