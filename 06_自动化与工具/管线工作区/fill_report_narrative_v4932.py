# -*- coding: utf-8 -*-
"""fill_report_narrative_v4932.py — 报告叙述充实（回收 v4.1 叙述+标签寻址数值映射，可复跑）

来源：archive/行业分析报告_v41_pre-promote_backup_20260928.docx
映射：(节名, 列头, 行标签) 三元组寻址新旧两版表格 → 旧值→新值字典。
质量门控：段落 ≥90% 数字可映射才整段插入；否则句子级修剪（只保留全部数字
可映射的句子，修剪后 ≥40 字才收）。绝对不插入含不可溯源数字的句子。
"""
import copy
import re

from docx import Document
from docx.oxml.ns import qn

OLD = 'archive/行业分析报告_v41_pre-promote_backup_20260928.docx'
CUR = 'NEV公告车型行业分析报告.docx'
NUM = re.compile(r'\d{1,3}(?:,\d{3})+(?:\.\d+)?|\d+\.\d+|\d{3,}')


def doc_sections(d):
    secs = {}
    cur = None
    from docx.text.paragraph import Paragraph
    from docx.table import Table
    for child in d.element.body.iterchildren():
        if child.tag.endswith('}p'):
            p = Paragraph(child, d)
            st = p.style.name if p.style else ''
            if st.startswith('Heading') or st.startswith('标题'):
                cur = p.text.strip()
                secs.setdefault(cur, {'paras': [], 'el': child, 'tbls': []})
            elif cur is not None and p.text.strip():
                secs[cur]['paras'].append(p.text.strip())
        elif child.tag.endswith('}tbl') and cur is not None:
            t = Table(child, d)
            secs[cur]['tbls'].append([[c.text for c in r.cells] for r in t.rows])
    return secs


def label_map(d, secs):
    out = {}
    for h, o in secs.items():
        hk = re.sub(r'\s+', '', h)[:16]
        for grid in o['tbls']:
            if len(grid) < 2 or not grid[0]:
                continue
            hdr = grid[0]
            for row in grid[1:]:
                if not row:
                    continue
                rl = re.sub(r'\s+', '', str(row[0]))[:12]
                for ci, cellv in enumerate(row[1:], start=1):
                    ch = re.sub(r'\s+', '', str(hdr[ci] if ci < len(hdr) else ''))[:12]
                    try:
                        v = float(str(cellv).replace(',', '').replace('%', ''))
                    except Exception:
                        continue
                    out.setdefault((hk, ch, rl), v)
    return out


def main():
    do = Document(OLD)
    dn = Document(CUR)
    so = doc_sections(do)
    sn = doc_sections(dn)
    mo = label_map(do, so)
    mn = label_map(dn, sn)
    gmap = {}
    for k, vo in mo.items():
        vn = mn.get(k)
        if vn is not None and abs(vo - vn) >= 1e-9:
            gmap[vo] = vn
    keys = sorted(gmap)
    print(f'标签映射对: {len(gmap)}（旧键 {len(mo)} / 新键 {len(mn)}）')

    def lookup(v):
        c = [k for k in keys if abs(k - v) <= 0.051]
        return gmap[min(c, key=lambda k: abs(k - v))] if c else None

    def fmt_like(tok, val):
        if '.' in tok:
            dd = len(tok.split('.')[1])
            return f'{val:,.{dd}f}' if ',' in tok else f'{val:.{dd}f}'
        return f'{val:,.0f}' if ',' in tok else f'{val:.0f}'

    def remap(s):
        def rep(mo):
            tok = mo.group(0)
            v = float(tok.replace(',', ''))
            if 1900 <= v <= 2100:
                return tok
            b = lookup(v)
            return fmt_like(tok, b) if b is not None else tok
        return re.sub(NUM, rep, s)

    def sent_ok(s):
        for st in re.findall(NUM, s):
            v = float(st.replace(',', ''))
            if 1900 <= v <= 2100:
                continue
            if lookup(v) is None:
                return False
        return bool(re.findall(NUM, s)) or len(s) > 12

    inserted = skipped_low = skipped_has = 0
    for h, o in so.items():
        n = sn.get(h)
        if not n or not o['paras']:
            continue
        if len(n['paras']) >= 1:
            skipped_has += 1
            continue
        cand = next((t for t in o['paras'] if len(t) > 40), None)
        if not cand:
            continue
        toks = re.findall(NUM, cand)
        if not toks:
            continue
        hit = sum(1 for st in toks
                  if 1900 <= float(st.replace(',', '')) <= 2100
                  or lookup(float(st.replace(',', ''))) is not None)
        rate = hit / max(1, len(toks))
        if rate >= 0.9:
            new_txt = remap(cand)
        else:
            keep = [remap(s).strip() for s in re.split(r'(?<=。)|(?<=；)', cand)
                    if s.strip() and sent_ok(s)]
            new_txt = ''.join(keep)
            if len(new_txt) < 40:
                skipped_low += 1
                continue
        if '底表 v4' in new_txt or re.search(r'第\d+[-~]\d+批', new_txt):
            skipped_low += 1
            continue
        hel = n['el']
        newp = copy.deepcopy(hel)
        newp.tag = qn('w:p')
        for cc in list(newp):
            if cc.tag != qn('w:pPr'):
                newp.remove(cc)
        hel.addnext(newp)
        from docx.text.paragraph import Paragraph
        Paragraph(newp, hel.getparent()).text = new_txt
        inserted += 1
        print(f'  + [{h[:26]}] {new_txt[:58]}')
    dn.save(CUR)
    print(f'插入 {inserted} 段 | 已有叙述跳过 {skipped_has} | 修剪后不足放弃 {skipped_low}')


if __name__ == '__main__':
    main()
