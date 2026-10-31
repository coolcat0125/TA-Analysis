# run_full.ps1 — DCD 全量数据库优化编排(浏览器直抓 + 自动回填)
#
# 四步流水线:
#   1) 建队列 + 导出抓取目标        python dcd_queue.py / export_targets.py
#   2) 浏览器抓取(人工)            打开 dongchedi.com → 控制台粘贴 extract_browser.js
#                                  → 选 targets.json → 自动下载 dcd_bundle_*.json
#   3) 导入 + 校验                 import_bundle.py(品牌校验 → mapping)
#   4) 批量回填 + 自检 + 看板       from-mapping 循环 apply,每批刷新看板
#
# 用法:
#   .\run_full.ps1                        # 全流程(第2步会停下来等你抓完)
#   .\run_full.ps1 -SkipExport            # 已导出过 targets,跳过第1步
#   .\run_full.ps1 -BundlesDir <目录>     # bundle 放在别处时指定
#
# 随时停:创建 dcd_auto\state\STOP 文件,循环跑完当前批即停
[CmdletBinding()]
param(
    [int]$BatchSize = 5,
    [int]$IntervalSec = 20,
    [switch]$SkipExport,
    [string]$BundlesDir = '',
    [string]$Python = 'python'
)
$ErrorActionPreference = 'Stop'
$Root = Split-Path -Parent $PSScriptRoot
$Auto = $PSScriptRoot
$State = Join-Path $Auto 'state'
$Bundles = if ($BundlesDir) { $BundlesDir } else { Join-Path $State 'bundles' }
$Log = Join-Path $State 'run_full_log.txt'
if (-not (Test-Path $State)) { New-Item -ItemType Directory -Path $State | Out-Null }
if (-not (Test-Path $Bundles)) { New-Item -ItemType Directory -Path $Bundles | Out-Null }

function Write-Log([string]$msg) {
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $msg"
    Add-Content -Path $Log -Value $line
    Write-Host $line
}

Set-Location $Root
Write-Log '======== DCD 全量优化开始 ========'

# ---------- 1) 队列 + 抓取目标 ----------
if (-not $SkipExport) {
    Write-Log '1/4 建缺口优先队列'
    & $Python "$Auto\dcd_queue.py"
    if ($LASTEXITCODE -ne 0) { Write-Log '队列生成失败'; exit 1 }
    Write-Log '1/4 导出抓取目标(targets.json / targets.csv)'
    & $Python "$Auto\export_targets.py"
    if ($LASTEXITCODE -ne 0) { Write-Log '目标导出失败'; exit 1 }
} else {
    Write-Log '1/4 跳过导出(沿用既有 targets.json)'
}
if (-not (Test-Path (Join-Path $State 'targets.json'))) { Write-Log '缺 targets.json'; exit 1 }

# ---------- 2) 浏览器抓取(人工) ----------
$existing = Get-ChildItem $Bundles -Filter 'dcd_bundle_*.json' -ErrorAction SilentlyContinue
if (-not $existing) {
    Write-Host ''
    Write-Host '  ┌──────────────────────────────────────────────────────────────┐'
    Write-Host '  │  第 2 步:浏览器抓取(必须人工做,约 1 小时)                     │'
    Write-Host '  │                                                              │'
    Write-Host '  │  1. 浏览器打开 https://www.dongchedi.com/ 并停在该页面         │'
    Write-Host '  │  2. F12 → Console → 粘贴 dcd_auto\extract_browser.js 全文 → 回车 │'
    Write-Host '  │  3. 首次运行弹框选择: dcd_auto\state\targets.json              │'
    Write-Host '  │  4. 每 15 个车系自动下载一个 bundle;中断后重跑会续抓            │'
    Write-Host '  │  5. 把下载目录里的 dcd_bundle_*.json 拷到:                     │'
    Write-Host "  │     $Bundles"
    Write-Host '  └──────────────────────────────────────────────────────────────┘'
    Write-Host ''
    Read-Host '抓取完成并把 bundle 拷到上面目录后,按回车继续'
} else {
    Write-Log "2/4 检测到已有 bundle $($existing.Count) 个,跳过人工抓取提示"
}

$bundles = Get-ChildItem $Bundles -Filter 'dcd_bundle_*.json' -ErrorAction SilentlyContinue
if (-not $bundles) { Write-Log '未检测到 bundle,无法继续'; exit 1 }

# ---------- 3) 导入 + 品牌校验 ----------
Write-Log "3/4 导入 bundle($($bundles.Count) 个)并做品牌校验"
& $Python "$Auto\import_bundle.py" --bundles $Bundles
if ($LASTEXITCODE -ne 0) { Write-Log '导入失败'; exit 1 }
Write-Log '3/4 刷新缺口队列(新数据已就位)'
& $Python "$Auto\dcd_queue.py" | Out-Null

# ---------- 4) 批量回填 ----------
Write-Log '4/4 批量回填(from-mapping 穷尽已抓取数据)'
& $Python "$Auto\dcd_pipeline.py" baseline
if ($LASTEXITCODE -ne 0) { Write-Log '基线捕获失败'; exit 1 }

$i = 0; $noProgress = 0; $lastGreen = -1
while ($true) {
    if (Test-Path (Join-Path $State 'STOP')) {
        Write-Log '检测到 STOP 文件,循环停止'
        Remove-Item (Join-Path $State 'STOP') -Force
        break
    }
    $i++
    & $Python "$Auto\dcd_pipeline.py" next-batch --n $BatchSize --from-mapping *> $Log.Append
    if ($LASTEXITCODE -ne 0) { Write-Log '已抓取数据全部处理完毕,回填结束'; break }
    $batch = (Get-ChildItem "$State\search\b*.json" | Where-Object { $_.Name -notlike '*.results.json' } |
              Sort-Object Name | Select-Object -Last 1).BaseName
    & $Python "$Auto\dcd_pipeline.py" decide --batch $batch *>> $Log
    & $Python "$Auto\dcd_pipeline.py" apply --batch $batch *>> $Log
    $code = $LASTEXITCODE
    & $Python "$Auto\render_dashboard.py" --no-coverage | Out-Null
    $green = & $Python -c "import json;print(json.load(open(r'$State\pipeline_state.json'))['totals'].get('green_fill',0))"
    if ("$green" -eq "$lastGreen") { $noProgress++ } else { $noProgress = 0 }
    $lastGreen = $green
    & $Python "$Auto\dcd_pipeline.py" eta 2>$null | ForEach-Object {
        try { $p = $_ | ConvertFrom-Json
              Write-Log "批次 $batch code=$code | 累计绿填 $green | 进度 $($p.series_done)/$($p.series_total)($($p.pct)%) 剩余 $($p.series_remaining) 车系 ETA $($p.eta_human)" }
        catch {}
    }
    if ($code -eq 3) { Write-Log '自检失败已自动回退,循环暂停——处理完后运行: python dcd_auto\dcd_pipeline.py resume'; break }
    if ($noProgress -ge 4) { Write-Log '连续 4 批零新增,防死循环退出'; break }
    Start-Sleep -Seconds $IntervalSec
}

& $Python "$Auto\render_dashboard.py"
Write-Log '======== 全量优化结束 ========'
& $Python "$Auto\dcd_pipeline.py" status
Write-Host ''
Write-Host "看板: $Root\DCD优化监控.html(15s 自动刷新)"
Write-Host "剩余待抓取车系与歧义清单见: $Auto\进度与ETA报告.md / $Auto\问题记录.md"
exit 0
