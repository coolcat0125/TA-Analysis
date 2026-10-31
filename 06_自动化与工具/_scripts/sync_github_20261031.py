# -*- coding: utf-8 -*-
"""
本地 -> 云端 结构对齐同步（staging 组装）

策略：以云端 clone 为底座，按七层映射归位；再用本地文件覆盖同名、补入新增。
保证云端独有文件不丢失。07_公告归档不入仓（851MB）。

用法：
    python sync_github_20261031.py stage     # 只组装 + 统计
    python sync_github_20261031.py apply     # 组装后同步进 git 工作区（不 commit）
"""
import os
import shutil
import sys
from pathlib import Path

LOCAL = Path(r"C:\00 AI\01 Project\14 TA-Analysis")
CLOUD = Path(r"C:\Users\scjc8y\AppData\Local\Temp\ta_probe")
STAGE = Path(r"C:\Users\scjc8y\AppData\Local\Temp\ta_stage")

# 根目录文档 -> 05_文档与流程/仓库文档
ROOT_DOCS = {"README.md", "CHANGELOG.md", "PROJECT_STATUS.md", "SYNC.md", "AUDIT.md", "LICENSE"}
# 根目录交付物 -> 02_交付物
ROOT_DELIVER = {
    "NEV公告车型行业分析报告.docx",
    "NEV公告车型行业分析_演示文稿.pptx",
    "NEV公告车型行业分析_数据底表.xlsx",
    "NEV公告数据看板.html",
    "NEV公告数据看板_发布版.html",
    "NEV公告车型行业分析_移动版.html",
}
WORK_DIRS = {"raw", "archive", "audit-output", "Update", "公告批次扫描_411-412_20261017"}
DOC_DIRS = {"updates", "candidate", "docs"}

# 本地排除（不入仓）
LOCAL_SKIP_DIRS = {
    "07_公告与减免购置税",
    "_tools",
    "_cloud_scan",
    "__pycache__",
    "node_modules",
    "_备份_脚本_文件名切换前",
    ".workbuddy",
    ".git",
}
LOCAL_SKIP_FILES_SUFFIX = (".pyc",)
LOCAL_SKIP_FILES_EXACT = {
    "NEV公告参数汇总表_合并版（341~412批）_备份_多批次合并前_20261028.xlsx",
    "NEV公告参数汇总表_合并版（341~412批）_备份_电池类型归一前_20261028.xlsx",
    "NEV公告参数汇总表_合并版（341~410批）_备份_合并前_20261028.xlsx",
    "NEV公告核定底表（341~412批）_v1_备份_多批次合并前_20261028.xlsx",
    "NEV公告核定底表（341~412批）_v1_备份_电池类型归一前_20261028.xlsx",
    "NEV公告核定底表（341~411批+412公示）_v1_备份_合并前_20261028.xlsx",
    # 中间版本演示文稿，仅保留定稿
    "NEV公告车型行业分析_2026Q4_v3.pptx",
    "NEV公告车型行业分析_2026Q4_v4936.pptx",
    "NEV公告车型行业分析_2026Q4_v4936b.pptx",
    "NEV公告车型行业分析_RESS_v4936.pptx",
    "NEV公告车型行业分析_演示文稿_v4934.pptx",
}


def map_cloud(rel: str) -> str:
    """云端相对路径 -> 新结构相对路径"""
    p = rel.replace("\\", "/")
    top = p.split("/")[0]
    if top == "核定底表":
        return "01_当前数据/" + p
    if top == "dcd_auto":
        return "06_自动化与工具/" + p
    if top in WORK_DIRS:
        return "06_自动化与工具/管线工作区/" + p
    if top in DOC_DIRS:
        return "05_文档与流程/" + p
    if top == ".github":
        return p
    if "/" not in p:
        if p in ROOT_DOCS:
            return "05_文档与流程/仓库文档/" + p
        if p in ROOT_DELIVER:
            return "02_交付物/" + p
        if p == "工信部新能源车型能耗数据20260915.xlsx":
            return "01_当前数据/能耗数据/" + p
    return "06_自动化与工具/管线工作区/" + p


# _scripts 白名单：只保留可复用工具链，一次性调试脚本不入仓
SCRIPT_KEEP_FILES = {
    "README_云端维护手册.md",
    "_sync_from_github.py",
    "sync_github_20261031.py",
}
SCRIPT_KEEP_DIRS = {"announce_archive"}


def skipped_local(rel: str) -> bool:
    parts = rel.replace("\\", "/").split("/")
    if any(x in LOCAL_SKIP_DIRS for x in parts[:-1]):
        return True
    # 06_自动化与工具/_scripts 只放行白名单
    if len(parts) >= 3 and parts[0] == "06_自动化与工具" and parts[1] == "_scripts":
        if len(parts) >= 4 and parts[2] in SCRIPT_KEEP_DIRS:
            pass
        elif len(parts) == 3 and parts[2] in SCRIPT_KEEP_FILES:
            pass
        else:
            return True
    fn = parts[-1]
    if fn.endswith(LOCAL_SKIP_FILES_SUFFIX):
        return True
    if fn in LOCAL_SKIP_FILES_EXACT:
        return True
    if fn.startswith("~$"):
        return True
    return False


def copy_file(src: Path, dst: Path) -> int:
    dst.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src, dst)
    return dst.stat().st_size


def main():
    mode = sys.argv[1] if len(sys.argv) > 1 else "stage"
    if STAGE.exists():
        shutil.rmtree(STAGE)
    STAGE.mkdir(parents=True)

    # 1) 云端底座归位
    cloud_map = {}
    n_cloud = 0
    for dirpath, dirnames, files in os.walk(CLOUD):
        if ".git" in Path(dirpath).parts:
            continue
        for fn in files:
            src = Path(dirpath) / fn
            rel = str(src.relative_to(CLOUD))
            cloud_map[map_cloud(rel)] = src
            n_cloud += 1
    for newrel, src in cloud_map.items():
        copy_file(src, STAGE / newrel)
    print(f"[云端底座] {n_cloud} 个文件已归位到七层结构")

    # 2) 本地覆盖 + 新增
    n_over, n_new, n_skip, bytes_new = 0, 0, 0, 0
    local_only = []
    for dirpath, dirnames, files in os.walk(LOCAL):
        rel_dir = str(Path(dirpath).relative_to(LOCAL))
        dirnames[:] = [d for d in dirnames if d not in LOCAL_SKIP_DIRS]
        for fn in files:
            src = Path(dirpath) / fn
            rel = fn if rel_dir == "." else str(Path(rel_dir) / fn)
            if skipped_local(rel):
                n_skip += 1
                continue
            dst = STAGE / rel.replace("\\", "/")
            if dst.exists():
                n_over += 1
            else:
                n_new += 1
                local_only.append(rel.replace("\\", "/"))
                bytes_new += src.stat().st_size
            copy_file(src, dst)
    print(f"[本地覆盖] 同名覆盖 {n_over} 个 | 本地新增 {n_new} 个 ({bytes_new/1048576:.1f} MB) | 排除 {n_skip} 个")

    # 3) 统计
    total, total_bytes = 0, 0
    for dirpath, _, files in os.walk(STAGE):
        for fn in files:
            total += 1
            total_bytes += (Path(dirpath) / fn).stat().st_size
    print(f"[结果] {total} 个文件 / {total_bytes/1048576:.1f} MB")

    # 4) 顶层结构
    print("\n顶层结构：")
    for d in sorted(p for p in STAGE.iterdir()):
        if d.is_dir():
            cnt = sum(1 for _ in d.rglob("*") if _.is_file())
            print(f"  {d.name:<28} {cnt:>5} 文件")
        else:
            print(f"  {d.name:<28} (文件)")

    with open(STAGE.parent / "stage_report.txt", "w", encoding="utf-8") as f:
        f.write(f"总文件 {total} / {total_bytes/1048576:.1f} MB\n")
        f.write(f"覆盖 {n_over} / 新增 {n_new} / 排除 {n_skip}\n\n本地新增清单:\n")
        for x in sorted(local_only):
            f.write(x + "\n")
    print(f"\n清单已写: {STAGE.parent / 'stage_report.txt'}")


if __name__ == "__main__":
    main()
