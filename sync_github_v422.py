# -*- coding: utf-8 -*-
"""v4.2.2 云端同步：把本地最新核心文件推送到 GitHub coolcat0125/TA-analysis（私有库）。

用法（先设置 token）：
  PowerShell:  $env:GITHUB_TOKEN="ghp_xxxx"; python sync_github_v422.py
  CMD:         set GITHUB_TOKEN=ghp_xxxx && python sync_github_v422.py

约定（协同规则）：任何 PUT 前先 GET 远程 SHA（竞态保护 + 跳过无变化文件）；
token 需对 coolcat0125/TA-analysis 有 repo 写权限。
Copyright © 2026 David YE
"""
import base64
import hashlib
import json
import os
import sys
import urllib.parse
import urllib.request
from pathlib import Path

HERE = Path(__file__).resolve().parent
REPO = "coolcat0125/TA-analysis"
BRANCH = "main"
TOKEN = os.environ.get("GITHUB_TOKEN", "").strip()
API = f"https://api.github.com/repos/{REPO}"

# 待同步文件（相对本目录；目录文件自动展开一层）
TARGETS = [
    "NEV公告参数汇总表_合并版（341~410批）.xlsx",
    "check_new_batch.py",
    "media_fill.py",
    "migrate_motor_v430.py",
    "clean_motor_v440.py",
    "NEV公告车型行业分析_数据底表.xlsx",
    "NEV公告车型行业分析_演示文稿.pptx",
    "NEV公告车型行业分析报告.docx",
    "NEV公告数据看板.html",
    "NEV公告数据看板_发布版.html",
    "generate_dashboard.py",
    "audit_data.py",
    "clean_outliers_v421.py",
    "fill_batch_dates_v421.py",
    "fill_logic_v422.py",
    "sync_github_v422.py",
    "CHANGELOG.md",
    "PROJECT_STATUS.md",
    "quality_gates.json",
    "audit-output/data_audit_report.json",
    "audit-output/data_audit_report.md",
    "audit-output/column_coverage.csv",
    "audit-output/data_inventory.json",
    "audit-output/battery_80kwh_priority.csv",
    "audit-output/motor_torque_missing.csv",
]
COMMIT_MSG = "v4.2.2 表内逻辑推理补全+1,883格；发布版取消导出/导入；v4.2.1 数据清洗+批次时间表官方日期"

for k in ("HTTPS_PROXY", "HTTP_PROXY", "https_proxy", "http_proxy"):
    os.environ.pop(k, None)
OPENER = urllib.request.build_opener(urllib.request.ProxyHandler({}))


def git_blob_sha(data: bytes) -> str:
    return hashlib.sha1(b"blob " + str(len(data)).encode() + b"\x00" + data).hexdigest()


def api(method: str, path: str, payload: dict | None = None):
    url = f"{API}{path}"
    req = urllib.request.Request(url, method=method, headers={
        "Authorization": f"token {TOKEN}",
        "Accept": "application/vnd.github.v3+json",
        "User-Agent": "ta-analysis-sync",
    })
    body = None
    if payload is not None:
        body = json.dumps(payload).encode()
        req.add_header("Content-Type", "application/json")
    with OPENER.open(req, body, timeout=60) as r:
        return json.loads(r.read().decode())


def main() -> int:
    if not TOKEN:
        print("ERROR: 请先设置 GITHUB_TOKEN（需对 coolcat0125/TA-analysis 有 repo 写权限）")
        print('  PowerShell:  $env:GITHUB_TOKEN="ghp_xxxx"; python sync_github_v422.py')
        return 2
    updated, skipped, failed = [], [], []
    # 一次性拉取远程文件树（仅 SHA，不含内容；contents API 对较大文件会断流）
    tree_sha = {}
    try:
        req = urllib.request.Request(
            f"{API}/git/trees/{BRANCH}?recursive=1",
            headers={"Authorization": f"token {TOKEN}",
                     "Accept": "application/vnd.github.v3+json",
                     "User-Agent": "ta-analysis-sync"})
        with OPENER.open(req, timeout=60) as r:
            tree = json.loads(r.read().decode())
        tree_sha = {b["path"]: b["sha"] for b in tree.get("tree", []) if b.get("type") == "blob"}
    except Exception as e:
        print(f"WARN: 树接口失败（将逐文件 GET 兜底）: {e}")
    for rel in TARGETS:
        p = HERE / rel
        if not p.exists():
            failed.append((rel, "本地文件不存在"))
            continue
        data = p.read_bytes()
        local_sha = git_blob_sha(data)
        remote_sha = tree_sha.get(rel)
        if remote_sha is None:
            try:
                meta = api("GET", f"/contents/{urllib.parse.quote(rel)}?ref={BRANCH}")
                remote_sha = meta.get("sha")
            except Exception as e:
                if "404" not in str(e):
                    failed.append((rel, f"GET 失败: {e}"))
                    continue
        if remote_sha and remote_sha == local_sha:
            skipped.append(rel)
            print(f"  = 无变化，跳过: {rel}")
            continue
        try:
            api("PUT", f"/contents/{urllib.parse.quote(rel)}", {
                "message": f"{COMMIT_MSG} · {Path(rel).name}",
                "content": base64.b64encode(data).decode(),
                "branch": BRANCH,
                **({"sha": remote_sha} if remote_sha else {}),
            })
            updated.append(rel)
            print(f"  ↑ 已更新: {rel}（{len(data)/1024:.0f} KB）")
        except Exception as e:
            failed.append((rel, str(e)[:200]))
            print(f"  ! 失败: {rel}: {str(e)[:200]}")
    print(f"\n完成：更新 {len(updated)}，跳过 {len(skipped)}，失败 {len(failed)}")
    for rel, err in failed:
        print(f"  失败清单: {rel}: {err}")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
