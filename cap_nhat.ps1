# Cap nhat bao cao chuoi thep len GitHub
# Chay boi: luc mo may (Startup) + lich 11:45 / 17:15 (Task Scheduler) + chay tay: chay_ngay.bat
# 1) Lay ban moi nhat tu GitHub (git pull)
# 2) Quet o chung nesting (Z:\13-KTSX-BPGC\.NESTING-CNC-DDC): moi du an lay lenh nesting moi nhat
#    (so lenh lon nhat, ban REV moi nhat; bo DELETE, file tam ~$, file do dang nho hon 1/2 ban dang dung)
# 3) Don file cu: 1_CUTTING, 2_TON_KHO chi giu file moi nhat (file cu chuyen sang ..\LUU_TRU_FILE_CU, khong xoa)
# 4) Day len GitHub -> GitHub tu dung bao cao
# 5) Lich su git qua lon (> 600 MB) thi gom gon
$ErrorActionPreference = 'Continue'
$root = $PSScriptRoot
$arch = Join-Path (Split-Path $root -Parent) 'LUU_TRU_FILE_CU'
$logf = Join-Path $root 'nhat_ky_day_len.txt'
$NES_GOC = 'Z:\13-KTSX-BPGC\.NESTING-CNC-DDC'
# Thu muc trong o nesting KHONG quet (du an da xong / thu muc phu). Them ten vao day neu can.
$BO_QUA = @('0.KHO-REMAIN','00-DATA','DELETE','TOOL','O CHUNG','BPGC+C.TRINH NOI BO','GIA CONG CAT LOC-AH+LA+BC',
            '01- DU AN H1033','02-MOMBASA','03-APEC-S3','06-MOCKUP-HSC')
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

function Quet-Nesting {
  $dst = Join-Path $root 'du_lieu\3_NESTING'
  if (-not (Test-Path $NES_GOC)) { Log "Khong vao duoc o nesting $NES_GOC (chua ket noi mang?) - bo qua buoc quet nesting"; return }
  foreach ($da in Get-ChildItem $NES_GOC -Directory) {
    if ($BO_QUA -contains $da.Name) { continue }
    # cac thu muc lenh: ...VT-<MA>-TT-<so>...
    $lenh = @(Get-ChildItem $da.FullName -Directory | ForEach-Object {
      if ($_.Name -match 'VT-([A-Z0-9]+)-TT-?(\d{3})' -and $_.Name -notmatch '(?i)ĐÃ HỦY|DA HUY|HỦY LỆNH|HUY LENH|DELETE') {
        [pscustomobject]@{ Dir = $_; Ma = $Matches[1]; So = [int]$Matches[2] } } })
    if (-not $lenh.Count) { continue }
    # chi theo ma xuat hien nhieu nhat trong du an (bo thu muc go sai ma, vd VT-CBD thay vi VT-CDB)
    foreach ($g in @($lenh | Group-Object Ma | Sort-Object Count -Descending | Select-Object -First 1)) {
      $ma = $g.Name
      $hien = @(Get-ChildItem $dst -File | Where-Object { $_.Name -match "^VT-$ma-TT-" -and $_.Extension -match '^\.xls' })
      $coHien = ($hien | Measure-Object Length -Maximum).Maximum
      $chon = $null
      foreach ($l in ($g.Group | Sort-Object So -Descending)) {
        $so = '{0:D3}' -f $l.So
        $ung = @(Get-ChildItem $l.Dir.FullName -Recurse -File -ErrorAction SilentlyContinue | Where-Object {
            $_.Extension -match '^\.xls[xbm]?$' -and $_.Name -notlike '~$*' -and $_.FullName -notmatch '(?i)\\DELETE\\' -and
            $_.Name -match "^VT-[A-Z0-9]+-TT-?0*$($l.So)(?!\d)" } | Sort-Object LastWriteTime -Descending)
        foreach ($f in $ung) {
          if ($coHien -and $f.Length -lt 0.5 * $coHien) { Log "  bo qua $($f.Name) ($([int]($f.Length/1KB)) KB, nho hon 1/2 ban dang dung - co the dang lam do)"; continue }
          if (-not $coHien -and $f.Length -lt 15KB) { continue }
          $chon = $f; break
        }
        if ($chon) { break }
      }
      if (-not $chon) { continue }
      $cu = $hien | Where-Object { $_.Name -eq $chon.Name }
      if ($cu -and $cu.Length -eq $chon.Length -and $cu.LastWriteTime -ge $chon.LastWriteTime) { continue }   # da moi nhat
      $maFile = if ($chon.Name -match '^VT-([A-Z0-9]+)-TT-') { $Matches[1] } else { $ma }
      foreach ($o in (Get-ChildItem $dst -File | Where-Object { ($_.Name -match "^VT-$maFile-TT-" -or $_.Name -match "^VT-$ma-TT-") -and $_.Name -ne $chon.Name })) {
        Remove-Item $o.FullName -Force; Log "  bo file nesting cu: $($o.Name)" }
      Copy-Item $chon.FullName (Join-Path $dst $chon.Name) -Force
      Log ("Nesting {0}: {1} ({2} KB, sua {3:dd/MM HH:mm})" -f $da.Name, $chon.Name, [int]($chon.Length/1KB), $chon.LastWriteTime)
    }
  }
}

Set-Location $root
Log '=== Bat dau ==='
git pull --rebase --autostash origin main 2>&1 | ForEach-Object { Log $_ }
try { Quet-Nesting } catch { Log "Loi khi quet nesting: $_" }
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
