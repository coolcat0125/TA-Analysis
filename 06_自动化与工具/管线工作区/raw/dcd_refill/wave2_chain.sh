#!/bin/bash
# wave2_chain.sh — 映射完成后自动接力：alias→抓取→jobgen→ingest（Wave 2）
# 用法：bash wave2_chain.sh   （在 raw/dcd_refill/ 下）
set -u
cd "$(dirname "$0")"
export PYTHONIOENCODING=utf-8
echo "[chain] $(date '+%H:%M:%S') alias 消解开始"
python dcd_map_direct.py 2>&1 | tail -8
python dcd_map_direct.py --alias 2>&1 | tail -25
echo "[chain] $(date '+%H:%M:%S') 批量抓取开始"
python dcd_fetch_batch.py 2>&1 | tail -40
echo "[chain] $(date '+%H:%M:%S') jobgen+ingest"
cd "../.."
for b in b01 b02 b03 b04 b05 b06 b07 b08 r09 r10 r11 r12 r13 r14; do
  python dcd_batch_run.py jobgen --batch $b >/dev/null 2>&1
  python dcd_batch_run.py ingest --batch $b 2>&1 | head -1
done
echo "[chain] $(date '+%H:%M:%S') 全部完成，等待 apply"
