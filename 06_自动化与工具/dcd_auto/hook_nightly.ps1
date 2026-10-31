# hook_nightly.ps1 — 挂接每日 23:00 夜间自动轮(可选)
#
# 定位:不替换、不修改现有夜间轮(check_new_batch → media_fill → fill_logic → fill_consensus
# → reconcile),只在其后串一段 DCD 批次处理。现有轮的收口逻辑(不改 CHANGELOG/不推送)不变。
#
# 用法一(手动验证):
#   .\hook_nightly.ps1 -MaxBatches 5
#
# 用法二(注册计划任务,每天 23:30 跑——给现有 23:00 轮 30 分钟收口):
#   schtasks /create /tn "DCD-夜间优化" /sc daily /st 23:30 /rl highest ^
#     /tr "powershell -ExecutionPolicy Bypass -File `"C:\...\TA-Analysis\dcd_auto\hook_nightly.ps1`""
#   (路径换成你机器上的实际路径;先手动跑通再注册)
#
# 与独立循环的分工:
#   - 想白天持续跑:用 run_loop.ps1(手动启动,STOP 文件可停)
#   - 想每晚自动跑:注册本脚本的计划任务
#   - 两者可并存,但不要同时跑(pipeline_state.json 是单写入者;apply 有备份+自检+回退兜底)
[CmdletBinding()]
param(
    [int]$MaxBatches = 5,
    [int]$BatchSize = 5,
    [string]$Python = 'python'
)
$ErrorActionPreference = 'Stop'
$Auto = $PSScriptRoot
$Root = Split-Path -Parent $Auto
$State = Join-Path $Auto 'state'
$Log = Join-Path $State 'hook_log.txt'
if (-not (Test-Path $State)) { New-Item -ItemType Directory -Path $State | Out-Null }

function Write-Log([string]$msg) {
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [nightly-hook] $msg"
    Add-Content -Path $Log -Value $line
    Write-Host $line
}

Set-Location $Root
Write-Log "夜间钩子启动:MaxBatches=$MaxBatches(接在每日 23:00 自动轮之后)"

# 禁写窗内直接退出(23:00 轮本身也不写库;本钩子同样遵守)
$now = Get-Date
if ($now.Hour -ge 23 -or $now.Hour -lt 6) {
    Write-Log '当前 23:00-06:00 禁写窗,钩子不动作(写库纪律)'
    exit 0
}

# 只跑循环,不重复捕获基线(基线由白天或上一次运行建立;没有则先建)
$stateFile = Join-Path $State 'pipeline_state.json'
if (-not (Test-Path $stateFile)) {
    Write-Log '无 pipeline_state.json,先捕获基线'
    & $Python "$Auto\dcd_pipeline.py" baseline
}

& "$Auto\run_loop.ps1" -MaxBatches $MaxBatches -BatchSize $BatchSize -IntervalSec 45 -SkipBaseline
$code = $LASTEXITCODE
Write-Log "循环结束,退出码 $code"

# 夜间钩子只做到"应用+自检+看板刷新";不推送、不动 CHANGELOG——
# 版本登记与推送按仓库 SYNC 协议由主线终端显式执行(见 Update/DCD自动优化工作流.md)
exit $code
