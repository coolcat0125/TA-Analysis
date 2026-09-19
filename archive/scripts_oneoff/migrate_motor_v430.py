# -*- coding: utf-8 -*-
"""v4.3.0 底表电机字段结构迁移（v2，正确顺序：采值→调列→写值）
================================================================
旧：电机峰值功率(kW)[数值] · 电机总功率(kW)[数值] · 电机功率/扭矩[混合文本]
新：前电机功率/扭矩[文本 P/T] · 后电机功率/扭矩[文本 P/T，两驱为空] · 电机总功率/扭矩[文本 P/T = 前+后]

规则：
  R1 旧「电机功率/扭矩」可解析 → 前/后按标注(F/前/R/后/F&R)或数字对拆分；总=前+后
  R2 旧列缺失/不可解析 → 数值列兜底：峰值≈总(±2%)视为单电机(前=峰值)；峰值≠总为双电机不可拆
     (前/后留空、原峰值台账留档)；仅峰值→前=峰值=总；总=总功率
  R3 解析总功率与数值总功率冲突 >2% → 总以数值列为准（台账记冲突）
  规范格式："P/T"（扭矩未知仅 "P"）；逐格台账（类型=结构迁移），迁移格浅灰 D9D9D9 底纹
"""
import re
import sys

import openpyxl

sys.stdout.reconfigure(encoding="utf-8")
SRC = "NEV公告参数汇总表_合并版（341~410批）.xlsx"
GRAY = openpyxl.styles.PatternFill("solid", fgColor="D9D9D9")
NUM = r"(\d+(?:\.\d+)?)"


def f2s(v):
    if v is None:
        return ""
    return str(int(v)) if float(v) == int(v) else str(v)


def parse_old_pt(raw):
    """解析旧「电机功率/扭矩」。返回 ('single', p, t, None, None) / ('dual', fp, ft, rp, rt) / None。"""
    if raw is None:
        return None
    s = str(raw).strip()
    if not s or not re.search(r"\d", s):
        return None
    s = s.replace("／", "/").replace("？", "?").replace("，", " ").replace(",", " ")
    kw_n = len(re.findall(r"kW", s, re.I))
    if re.match(r"^\s*F\s*&\s*R", s, re.I):          # F&R 同规格双电机
        nums = re.findall(NUM, s)
        if len(nums) >= 2:
            return ("dual", float(nums[0]), float(nums[1]), float(nums[0]), float(nums[1]))
        if len(nums) == 1:
            return ("dual", float(nums[0]), None, float(nums[0]), None)
        return None
    mf = re.search(r"[Ff前]\s*[:：]?\s*" + NUM + r"\s*(?:/\s*" + NUM + r")?", s)
    mr = re.search(r"[Rr后]\s*[:：]?\s*" + NUM + r"\s*(?:/\s*" + NUM + r")?", s)
    if mf and mr:
        return ("dual",
                float(mf.group(1)), float(mf.group(2)) if mf.group(2) else None,
                float(mr.group(1)), float(mr.group(2)) if mr.group(2) else None)
    nums = [float(x) for x in re.findall(NUM, s)]
    if kw_n >= 2 and len(nums) == 2:                  # "130kW/201kW"、"/33kW/100kW"
        return ("dual", nums[0], None, nums[1], None)
    if len(nums) == 1:
        return ("single", nums[0], None)
    if len(nums) == 2:
        if "?" in s:                                  # '#/?' 扭矩未知
            return ("single", nums[0], None)
        return ("single", nums[0], nums[1])
    if len(nums) == 4:                                # 无标注双对 "30/120/36/95"
        return ("dual", nums[0], nums[1], nums[2], nums[3])
    return None                                       # 3 数字等歧义 → 不可解析


def pt_cell(p, t):
    if p is None:
        return ""
    return f2s(p) if t is None else f"{f2s(p)}/{f2s(t)}"


def main():
    wb = openpyxl.load_workbook(SRC)
    ws = wb["NEV公告参数汇总"]
    hdr = [str(c.value).strip() if c.value else "" for c in ws[1]]
    hi = {h: i for i, h in enumerate(hdr)}
    for need in ("电机峰值功率(kW)", "电机总功率(kW)", "电机功率/扭矩", "批次", "产品型号", "动力类型"):
        assert need in hi, f"缺列: {need}"
    i_peak, i_tot, i_pt = hi["电机峰值功率(kW)"], hi["电机总功率(kW)"], hi["电机功率/扭矩"]

    # ---- 1) 采值（仅内存，不动表）----
    rows = []
    for r in range(2, ws.max_row + 1):
        if ws.cell(r, 1).value is None:
            continue
        rows.append({
            "r": r,
            "b": ws.cell(r, hi["批次"] + 1).value,
            "m": ws.cell(r, hi["产品型号"] + 1).value,
            "d": ws.cell(r, hi["动力类型"] + 1).value,
            "peak": ws.cell(r, i_peak + 1).value,
            "tot": ws.cell(r, i_tot + 1).value,
            "pt": ws.cell(r, i_pt + 1).value,
        })
    print(f"数据行 {len(rows)}")

    # ---- 2) 逐行解析（纯计算）----
    cr = wb["变更记录"]
    seq = 0
    for rr in range(cr.max_row, 1, -1):
        v = cr.cell(rr, 1).value
        if isinstance(v, int):
            seq = v
            break
    led = []
    stats = {"single": 0, "dual": 0, "fb_eq": 0, "fb_unsplit": 0, "unparse": 0,
             "conflict": 0, "empty": 0}
    for rec in rows:
        peak = rec["peak"]
        tot = rec["tot"]
        pt = rec["pt"]
        peak_f = float(peak) if isinstance(peak, (int, float)) else None
        tot_f = float(tot) if isinstance(tot, (int, float)) else None
        parsed = parse_old_pt(pt)
        front = rear = ""
        tp = tt = None
        note = ""
        if parsed:
            if parsed[0] == "single":
                _, p, t = parsed
                front, rear, tp, tt = pt_cell(p, t), "", p, t
                note = "R1旧功率扭矩列·单电机解析"
                stats["single"] += 1
            else:
                _, fp, ft, rp, rt = parsed
                front, rear = pt_cell(fp, ft), pt_cell(rp, rt)
                tp = fp + rp
                tt = ft + rt if (ft is not None and rt is not None) else None
                note = "R1旧功率扭矩列·前后拆分"
                stats["dual"] += 1
        else:
            if pt is not None and str(pt).strip():
                stats["unparse"] += 1
                led.append([rec["b"], rec["m"], rec["d"], "电机功率/扭矩(退役)", "结构迁移",
                            str(pt)[:60], "(脏值不迁移，原值留档)", "R2旧列脏值"])
            if peak_f is not None and tot_f is not None:
                if abs(peak_f - tot_f) / max(tot_f, 1) <= 0.02:
                    front, tp = f2s(peak_f), peak_f
                    note = "R2峰值≈总·单电机兜底"
                    stats["fb_eq"] += 1
                else:
                    tp = tot_f
                    note = "R2峰值≠总·双电机不可拆"
                    stats["fb_unsplit"] += 1
                    led.append([rec["b"], rec["m"], rec["d"], "电机峰值功率(kW)(退役)", "结构迁移",
                                f2s(peak_f), "(双电机不可拆，峰值留档)", note])
            elif peak_f is not None:
                front, tp = f2s(peak_f), peak_f
                note = "R2仅有峰值"
                stats["fb_eq"] += 1
            elif tot_f is not None:
                tp = tot_f
                note = "R2仅有总功率"
                stats["fb_unsplit"] += 1
            else:
                note = "无源"
                stats["empty"] += 1
        if tp is not None and tot_f is not None and abs(tp - tot_f) / max(tot_f, 1) > 0.02:
            stats["conflict"] += 1
            led.append([rec["b"], rec["m"], rec["d"], "电机总功率/扭矩", "结构迁移",
                        f"解析{pt_cell(tp, tt)}/数值{f2s(tot_f)}", pt_cell(tp, tt),
                        "R3备注：解析总与旧数值列冲突>2%；按用户规则总=前+后，旧数值留档于此"])
        rec.update(front=front, rear=rear, total=pt_cell(tp, tt), note=note)

    # ---- 3) 调列结构（先于写值）----
    ws.cell(1, i_peak + 1).value = "前电机功率/扭矩"       # 峰值 → 前
    ws.insert_cols(i_tot + 2)                              # 总功率列后插"后电机功率/扭矩"
    ws.cell(1, i_tot + 2).value = "后电机功率/扭矩"
    ws.cell(1, i_tot + 1).value = "电机总功率/扭矩"
    i_pt_del = hi["电机功率/扭矩"] + 1 + 1                 # 旧混合列（insert 后右移一位）
    assert str(ws.cell(1, i_pt_del).value).strip() == "电机功率/扭矩"
    ws.delete_cols(i_pt_del)
    new_hdr = [str(c.value).strip() if c.value else "" for c in ws[1]]
    iF = new_hdr.index("前电机功率/扭矩") + 1
    iR = new_hdr.index("后电机功率/扭矩") + 1
    iT = new_hdr.index("电机总功率/扭矩") + 1
    print(f"列结构完成: 前@{iF} 后@{iR} 总@{iT}，共 {len(new_hdr)} 列")

    # ---- 4) 写值 + 台账 ----
    for rec in rows:
        r = rec["r"]
        for col, val, old, field in (
                (iF, rec["front"] or None, rec["peak"], "前电机功率/扭矩"),
                (iR, rec["rear"] or None, None, "后电机功率/扭矩"),
                (iT, rec["total"] or None, rec["tot"], "电机总功率/扭矩")):
            cell = ws.cell(r, col)
            if val is not None:
                cell.value = val
                cell.fill = GRAY
                led.append([rec["b"], rec["m"], rec["d"], field, "结构迁移",
                          "" if old is None else str(old)[:60], val, rec["note"]])
            else:
                cell.value = None

    for row in led:
        seq += 1
        cr.append([seq] + row)
    # ---- 5) 颜色说明 ----
    cs = wb["颜色说明"]
    have = {str(row[1] or "") for row in cs.iter_rows(values_only=True) if len(row) > 1}
    if "结构迁移" not in have:
        cs.append(["浅灰底纹", "结构迁移",
                   "v4.3.0 电机字段重构：旧「电机峰值功率/电机总功率/电机功率/扭矩」按解析规则迁移为"
                   "「前电机功率/扭矩、后电机功率/扭矩、电机总功率/扭矩」（前+后=总），原值见变更记录。"])

    wb.save(SRC)
    print(f"台账追加 {len(led)} 条（末序号 {seq}）")
    print("统计:", stats)
    print("已保存")


if __name__ == "__main__":
    main()
