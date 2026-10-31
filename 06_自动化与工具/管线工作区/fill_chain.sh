#!/bin/bash
cd "/d/#AI/00 Project/17 TA analysis/01-主仓库/TA-Analysis" || exit 1
export PYTHONIOENCODING=utf-8
SEGS=("365,366,367,368,369,370,371,372,373,374" "355,356,357,358,359,360,361,362,363,364" "341,342,343,344,345,346,347,348,349,350,351,352,353,354")
n=0
for SEG in "${SEGS[@]}"; do
  n=$((n+1))
  echo "[$(date '+%T')] === 补录段$n: $SEG ==="
  python media_fill.py --scope "$SEG" --delay 0.45 2>&1 | tail -6
  PYTHONIOENCODING=utf-8 python verify_consistency.py 2>&1 | tail -2
  python generate_dashboard.py >/dev/null 2>&1 && python generate_dashboard.py --release >/dev/null 2>&1
  git add -A >/dev/null 2>&1
  git commit -m "data: 补录段$n(批次$SEG)媒体补空" >/dev/null 2>&1
  git pull --rebase origin main >/dev/null 2>&1
  git push origin main >/dev/null 2>&1 && echo "  已提交推送"
  sleep 60
done
echo "[$(date '+%T')] === 末轮: 共识+公式级联 ==="
python enrich_v5.py "NEV公告参数汇总表_合并版（341~410批）.xlsx" "D:/#AI/00 Project/17 TA analysis/02-数据底表/data/nev-announcements/db/master_export_fixed.csv" 2>&1 | tail -2
python fill_consensus_v5.py 2>&1 | tail -4
PYTHONIOENCODING=utf-8 python verify_consistency.py 2>&1 | tail -2
python audit_data.py --strict 2>&1 | tail -1
python generate_dashboard.py >/dev/null 2>&1 && python generate_dashboard.py --release >/dev/null 2>&1
git add -A >/dev/null 2>&1
git commit -m "data: 末轮共识+公式级联(补录链收口)" >/dev/null 2>&1
git pull --rebase origin main >/dev/null 2>&1
git push origin main >/dev/null 2>&1 && echo "  已提交推送"
echo "CHAIN DONE $(date '+%F %T')"
