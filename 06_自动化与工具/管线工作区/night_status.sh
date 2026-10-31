#!/bin/bash
# night_status.sh — 后台运行状态采集（供10分钟播报调用）
cd "/d/#AI/00 Project/17 TA analysis/01-主仓库/TA-Analysis" || exit 1
echo "== $(date '+%F %T') =="
# 1) 扫描进度
python - << 'PYEOF' 2>/dev/null
import json, os
p=r'D:\#AI\00 Project\17 TA analysis\02-数据底表\data\nev-announcements\raw\official_nev_passenger_scan.json'
if os.path.exists(p):
    d=json.load(open(p,encoding='utf-8'))
    done=sorted(d,key=int)
    total=sum(len(v) for v in d.values())
    latest=done[-1] if done else '-'
    recent={k:len(d[k]) for k in done[-3:]}
    print(f"[扫描] {len(done)}/70 批 | 最新{latest}批 | 累计{total}款 | 近3批:{recent}")
else:
    print("[扫描] 未开始")
PYEOF
# 2) 进程健康
n=$(powershell -Command "(Get-CimInstance Win32_Process | Where-Object { \$_.Name -eq 'python.exe' -and \$_.CommandLine -match 'eidc_scan' } | Measure-Object).Count" 2>/dev/null | tr -d '\r\n ')
echo "[进程] eidc_scan存活=$n"
# 3) git 状态
echo "[git] $(git log --oneline -1 | head -c 60) | $(git status -sb | head -1 | head -c 50)"
