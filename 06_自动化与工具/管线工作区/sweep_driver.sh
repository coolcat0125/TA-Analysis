#!/bin/bash
# sweep_driver.sh v2 — 串行驱动（等待逻辑用独立 ps1，无转义风险）
cd "/d/#AI/00 Project/17 TA analysis/01-主仓库/TA-Analysis" || exit 1
export PYTHONIOENCODING=utf-8
STOP_AT=$(date -d "05:20" +%s); [ $STOP_AT -le $(date +%s) ] && STOP_AT=$((STOP_AT+86400))
SEGS=("404,403,402,401,400" "399,398,397,396,395" "394,393,392,391,390" "389,388,387,386,385" "384,383,382,381,380" "379,378,377,376,375" "374,373,372,371,370" "369,368,367,366,365" "364,363,362,361,360" "359,358,357,356,355" "354,353,352,351,350" "349,348,347,346,345" "344,343,342,341")
# 等待任何在跑的 media_fill 结束（独立 ps1，退出码 0=在跑）
while powershell -ExecutionPolicy Bypass -File check_media.ps1; do
  echo "[$(date '+%T')] 检测到 media_fill 在跑，等待..."
  sleep 60
done
n=0
for SEG in "${SEGS[@]}"; do
  now=$(date +%s)
  if [ $now -ge $STOP_AT ]; then echo "[$(date '+%T')] 05:20 停机线，剩余段收尾后处理"; break; fi
  # 开段前再查一次并发
  if powershell -ExecutionPolicy Bypass -File check_media.ps1; then
    echo "[$(date '+%T')] 检测到并发 media_fill，跳过本段防覆盖"; sleep 120; continue
  fi
  n=$((n+1))
  echo "[$(date '+%T')] === 段$n: 批次 $SEG ==="
  python media_fill.py --scope "$SEG" --delay 0.4 2>&1 | tail -4
  git add -A >/dev/null 2>&1
  git commit -m "data: 由近及远扫描段$n(批次$SEG)媒体补空" >/dev/null 2>&1
  git pull --rebase origin main >/dev/null 2>&1
  git push origin main >/dev/null 2>&1 && echo "  已提交推送"
  sleep 45
done
echo "[$(date '+%T')] 媒体全量扫描完成"
if [ $(date +%s) -lt $(date -d "04:30" +%s) ]; then
  echo "进入共识轮（第二轮扫描）"
  python enrich_v5.py "NEV公告参数汇总表_合并版（341~410批）.xlsx" "D:/#AI/00 Project/17 TA analysis/02-数据底表/data/nev-announcements/db/master_export_fixed.csv" 2>&1 | tail -2
  python fill_consensus_v5.py 2>&1 | tail -3
  python generate_dashboard.py >/dev/null 2>&1 && python generate_dashboard.py --release >/dev/null 2>&1
  python verify_consistency.py 2>&1 | tail -2
  git add -A >/dev/null 2>&1; git commit -m "data: 共识轮(第二轮扫描)enrich+consensus回填" >/dev/null 2>&1
  git pull --rebase origin main >/dev/null 2>&1; git push origin main >/dev/null 2>&1
  echo "共识轮完成"
fi
echo "DRIVER DONE $(date '+%F %T')"
