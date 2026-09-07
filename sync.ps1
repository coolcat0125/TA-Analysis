# sync.ps1 — TA-Analysis 多终端一键同步（只快进，不自动 push）
$OutputEncoding = [Text.Encoding]::UTF8
[Console]::OutputEncoding = [Text.Encoding]::UTF8
chcp 65001 | Out-Null
Set-Location -Path $PSScriptRoot
function Step($m) { Write-Host ("==> " + $m) -ForegroundColor Cyan }
function GitQ([string[]]$GitArgs) { & git @GitArgs 2>&1 | ForEach-Object { "$_" } ; return $LASTEXITCODE }

Step "fetch origin"
GitQ @("fetch","origin","--prune") | Out-Null

$local  = (GitQ @("rev-parse","HEAD") | Select-Object -First 1)
$remote = (GitQ @("rev-parse","origin/main") | Select-Object -First 1)
if ("$local" -eq "$remote") {
    Step ("本地已是最新（" + $local.Substring(0,7) + "）")
} else {
    $ahead  = [int](GitQ @("rev-list","--count","origin/main..HEAD") | Select-Object -First 1)
    $behind = [int](GitQ @("rev-list","--count","HEAD..origin/main") | Select-Object -First 1)
    if ($ahead -gt 0 -and $behind -gt 0) { Write-Host "!! 本地与远端分叉（ahead=$ahead behind=$behind），请人工处理" -ForegroundColor Red; exit 1 }
    if ($ahead -gt 0) { Write-Host "!! 本地领先 $ahead 个提交未推送；本脚本不自动 push，请人工确认后 git push" -ForegroundColor Yellow; exit 0 }
    Step ("快进 " + $behind + " 个提交 ...")
    GitQ @("merge","--ff-only","origin/main") | Out-Null
    if ($LASTEXITCODE -ne 0) { Write-Host "!! 快进失败" -ForegroundColor Red; exit 1 }
}

# 工作树检查：纯换行符幻影自动 renormalize；真实改动仅提示
$dirty = (GitQ @("status","--porcelain") | Where-Object { $_ -match '\S' -and $_ -notmatch '^\d+$' })
if ($dirty) {
    GitQ @("diff","--ignore-cr-at-eol","--quiet") | Out-Null
    if ($LASTEXITCODE -eq 0) {
        Step "检测到纯换行符幻影，执行 renormalize"
        GitQ @("add","--renormalize",".") | Out-Null
        GitQ @("commit","-m","chore: EOL renormalize（零内容变化）") | Out-Null
        Write-Host "   已本地提交；请随下次 git push 一并推送" -ForegroundColor Yellow
    } else {
        Write-Host "!! 工作树有真实改动，未做处理：" -ForegroundColor Yellow
        $dirty | ForEach-Object { Write-Host ("   " + $_) }
    }
}

Step "质量门禁 audit --strict（临时目录，不触碰跟踪的 audit-output）"
$tmpAudit = Join-Path $env:TEMP ("audit_check_" + [guid]::NewGuid().ToString("N").Substring(0,8))
python audit_data.py --strict --output-dir $tmpAudit 2>&1 | Select-Object -Last 1
try {
    $r = Get-Content (Join-Path $tmpAudit "data_audit_report.json") -Raw -Encoding UTF8 | ConvertFrom-Json
    $verdict = "FAIL"; if ($r.quality_gates.passed) { $verdict = "PASS" }
    Write-Host ("    记录数=" + $r.dataset.records + "  门禁=" + $verdict)
} catch { Write-Host "    (审计报告读取跳过)" }

$firstLine = (Get-Content "CHANGELOG.md" -Encoding UTF8 | Where-Object { $_ -match '^## \[(.+)\]' } | Select-Object -First 1)
$ver = $firstLine -replace '^## \[(.+?)\].*','$1'
Write-Host ("==> 同步完成：当前版本 v" + $ver + " · HEAD " + (GitQ @("rev-parse","--short","HEAD") | Select-Object -First 1)) -ForegroundColor Green
