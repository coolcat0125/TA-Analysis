# run_loop.ps1 — DCD 自动优化主循环(护栏全自动)
#
# 用法:
#   .\run_loop.ps1                          # 默认:每批 5 车系,最多 10 批,批间隔 60s
#   .\run_loop.ps1 -MaxBatches 40 -IntervalSec 30
#   .\run_loop.ps1 -BatchSize 8
#
# 循环体(每批):
#   1) dcd_pipeline.py next-batch      → state/search/b<NNN>.json(缺口优先队列取车系)
#   2) dcd_extract.ps1  -Mode search   → Tabbit 搜索页取候选(批次请求 → candidates)
#   3) dcd_pipeline.py decide          → 品牌/包含校验定链(accept/reject/hold)
#   4) dcd_extract.ps1  -Mode fetch    → Tabbit 参数页取 rawData → raw/dcd_refill/
#   5) dcd_pipeline.py apply           → 备份→写入→三件套自检→失败自动回退并暂停
#   6) render_dashboard.py             → 刷新 DCD优化监控.html
#
# 停止条件(满足即停,不重试):
#   - state/STOP 文件存在(随时 touch 该文件即优雅停)
#   - apply 返回 3(自检失败已回退,循环暂停待人工)
#   - 队列无待处理车系
#   - 达到 -MaxBatches
#
# 时段纪律:22:40-06:00 不启动新批次的 apply(pipeline 自身也会拒写,--force 才越权)
[CmdletBinding()]
param(
    [int]$MaxBatches = 10,
    [int]$BatchSize = 5,
    [int]$IntervalSec = 60,
    [switch]$SkipBaseline,
    [string]$Python = 'python'
)
$ErrorActionPreference = 'Stop'
$Root = Split-Path -Parent $PSScriptRoot
$Auto = $PSScriptRoot
$State = Join-Path $Auto 'state'
$StopFile = Join-Path $State 'STOP'
$LoopLog = Join-Path $State 'loop_log.txt'
if (-not (Test-Path $State)) { New-Item -ItemType Directory -Path $State | Out-Null }

function Write-Log([string]$msg) {
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $msg"
    Add-Content -Path $LoopLog -Value $line
    Write-Host $line
}

function InWriteWindow {
    $now = Get-Date
    return -not (($now.Hour -eq 22 -and $now.Minute -ge 40) -or $now.Hour -ge 23 -or $now.Hour -lt 6)
}

Set-Location $Root
Write-Log "==== DCD 主循环启动:MaxBatches=$MaxBatches BatchSize=$BatchSize Interval=${IntervalSec}s ===="

if (-not $SkipBaseline) {
    Write-Log '捕获门禁基线…'
    & $Python "$Auto\dcd_pipeline.py" baseline
    if ($LASTEXITCODE -ne 0) { Write-Log '基线捕获失败,终止'; exit 1 }
}

$done = 0
while ($done -lt $MaxBatches) {
    if (Test-Path $StopFile) {
        Write-Log '检测到 STOP 文件,循环停止'
        Remove-Item $StopFile -Force
        & $Python "$Auto\render_dashboard.py" --no-coverage
        exit 0
    }
    if (-not (InWriteWindow)) {
        Write-Log '当前在禁写窗(22:40-06:00),等待到可写窗口…'
        Start-Sleep -Seconds 300
        continue
    }

    Write-Log "--- 第 $($done+1)/$MaxBatches 批:取车系 ---"
    & $Python "$Auto\dcd_pipeline.py" next-batch --n $BatchSize
    if ($LASTEXITCODE -ne 0) { Write-Log '队列已无待处理车系,循环结束'; break }

    # 从 next-batch 的 stdout 抓不到 batch id,改从 search 目录取最新
    $latest = Get-ChildItem "$State\search\b*.json" | Sort-Object LastWriteTime -Descending |
              Where-Object { $_.Name -notlike '*.results.json' } | Select-Object -First 1
    if (-not $latest) { Write-Log '未找到批次请求文件,终止'; exit 1 }
    $batch = [IO.Path]::GetFileNameWithoutExtension($latest.Name)
    Write-Log "批次 $batch :搜索阶段"
    & "$Auto\dcd_extract.ps1" -Mode search -Batch $batch
    if ($LASTEXITCODE -ne 0) { Write-Log "搜索阶段失败(退出码 $LASTEXITCODE),跳过本批"; $done++; continue }

    Write-Log "批次 $batch :定链"
    & $Python "$Auto\dcd_pipeline.py" decide --batch $batch
    if ($LASTEXITCODE -ne 0) { Write-Log '定链失败,跳过本批'; $done++; continue }

    Write-Log "批次 $batch :抓取参数页"
    & "$Auto\dcd_extract.ps1" -Mode fetch -Batch $batch
    if ($LASTEXITCODE -ne 0) { Write-Log "抓取阶段失败(退出码 $LASTEXITCODE),跳过本批"; $done++; continue }

    Write-Log "批次 $batch :应用+自检"
    & $Python "$Auto\dcd_pipeline.py" apply --batch $batch
    $applyCode = $LASTEXITCODE
    & $Python "$Auto\render_dashboard.py" --no-coverage
    if ($applyCode -eq 3) {
        Write-Log '自检失败已自动回退,循环暂停(人工处理后在 state 里清 paused_reason 再跑)'
        exit 3
    }
    if ($applyCode -eq 2) {
        Write-Log '撞上禁写窗,本批未写库;等窗口后重跑本批即可'
    }
    $done++
    if ($done -lt $MaxBatches) {
        Write-Log "批间隔 ${IntervalSec}s…"
        Start-Sleep -Seconds $IntervalSec
    }
}

Write-Log "==== 主循环结束:共处理 $done 批 ===="
& $Python "$Auto\render_dashboard.py"
exit 0
