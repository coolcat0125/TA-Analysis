# dcd_extract.ps1 — 懂车帝抽取驱动(Tabbit 用户会话通道)
#
# 用法:
#   .\dcd_extract.ps1 -Mode probe                    # 首次校准:探测运行时与页面结构
#   .\dcd_extract.ps1 -Mode search -Batch b001       # 搜索阶段:批次请求 → 候选车系
#   .\dcd_extract.ps1 -Mode fetch  -Batch b001       # 抓取阶段:定链 sid → raw/dcd_refill/
#   .\dcd_extract.ps1 -Mode search -Batch b001 -JsPath <你自己的抽取.js>
#
# 调用式依据:Update/DCD参数回填_进度跟踪.md「通道配方(已验证,复用)」——
#   %LOCALAPPDATA%\Tabbit\LocalAgent\bin\tabbit-cli.exe nodejs --task dcd-params
#     --request-id <ID> --timeout-ms 180000 < <js文件>
#   (Git Bash 调用需 MSYS_NO_PATHCONV=1 cmd.exe /d /c <包装cmd>;PowerShell 下用 cmd /d /c 包一层吃 < 重定向)
#
# 退出码:0=成功;2=环境未就绪(probe 未过/无页面对象);3=Tabbit 调用失败(含重试后)
[CmdletBinding()]
param(
    [ValidateSet('probe','search','fetch')][string]$Mode = 'probe',
    [string]$Batch = '',
    [string]$JsPath = '',
    [int]$Retry = 2
)
$ErrorActionPreference = 'Stop'
$Root = Split-Path -Parent $PSScriptRoot
$Auto = $PSScriptRoot
$State = Join-Path $Auto 'state'
$Raw = Join-Path $Root 'raw\dcd_refill'
$Log = Join-Path $State 'extract_log.txt'
if (-not (Test-Path $State)) { New-Item -ItemType Directory -Path $State | Out-Null }
if (-not (Test-Path $Raw)) { New-Item -ItemType Directory -Path $Raw | Out-Null }
if (-not $JsPath) { $JsPath = Join-Path $Auto 'dcd_extract.js' }
if (-not (Test-Path $JsPath)) { Write-Error "找不到抽取脚本: $JsPath"; exit 3 }

$TabbitCli = Join-Path $env:LOCALAPPDATA 'Tabbit\LocalAgent\bin\tabbit-cli.exe'
if (-not (Test-Path $TabbitCli)) {
    Write-Error "找不到 tabbit-cli.exe: $TabbitCli(确认 Tabbit LocalAgent 已安装)"
    exit 3
}

# 依据模式装配输入/输出
$Input = ''; $Out = ''
switch ($Mode) {
    'probe'  { $Out = Join-Path $State 'probe.json' }
    'search' {
        if (-not $Batch) { Write-Error 'search 模式需要 -Batch'; exit 3 }
        $Input = Join-Path $State "search\$Batch.json"
        $Out   = Join-Path $State "search\$Batch.results.json"
    }
    'fetch'  {
        if (-not $Batch) { Write-Error 'fetch 模式需要 -Batch'; exit 3 }
        $Input = Join-Path $State "decided\$Batch.json"
        $Out   = Join-Path $State "decided\$Batch.fetchlog.json"
    }
}
if ($Mode -ne 'probe' -and -not (Test-Path $Input)) { Write-Error "缺输入文件: $Input"; exit 3 }

$env:DCD_MODE = $Mode
$env:DCD_INPUT = $Input
$env:DCD_OUT = $Out
$env:DCD_RAW = $Raw
$RequestId = "dcd-$Mode-$(Get-Date -Format 'HHmmss')"

function Write-Log([string]$msg) {
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $msg"
    Add-Content -Path $Log -Value $line
    Write-Host $line
}

$attempt = 0
while ($attempt -le $Retry) {
    $attempt++
    # PowerShell 无 "<" 重定向,经 cmd /d /c 包一层;JS 路径带引号防空格
    $cmd = "`"$TabbitCli`" nodejs --task dcd-params --request-id $RequestId --timeout-ms 180000 < `"$JsPath`""
    Write-Log "调用 Tabbit(第 $attempt 次): $cmd"
    $p = Start-Process -FilePath 'cmd.exe' -ArgumentList '/d','/c',$cmd -NoNewWindow -Wait -PassThru
    $code = $p.ExitCode
    Write-Log "Tabbit 退出码: $code"
    if ($code -eq 0 -and (Test-Path $Out)) {
        # 结果自检:probe/search 的输出必须是合法 JSON 且无 error 字段
        try {
            $j = Get-Content $Out -Raw | ConvertFrom-Json
            if ($j.PSObject.Properties.Name -contains 'error' -or $j.PSObject.Properties.Name -contains 'fatal') {
                Write-Log "输出含 error 字段: $($j.error)$($j.fatal)"
                if ($attempt -le $Retry) { Start-Sleep -Seconds 10; continue }
                exit 2
            }
        } catch {
            Write-Log "输出不是合法 JSON: $_"
            if ($attempt -le $Retry) { Start-Sleep -Seconds 10; continue }
            exit 2
        }
        Write-Log "完成 → $Out"
        exit 0
    }
    if ($attempt -le $Retry) { Start-Sleep -Seconds 15 }
}
Write-Log "Tabbit 调用重试 $Retry 次仍失败"
exit 3
