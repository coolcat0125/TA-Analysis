# -*- coding: utf-8 -*-
"""eidc_full_sweep_v1.py — EIDC 官方批次核证+核心参数初次状态记录（可复跑，支持断点续跑）

用户指令：结合 EIDC 新查询系统，对全库做官方批次核证，核心参数记录初次状态，由近及远。
产出：audit-output/eidc_full_sweep_results.json（逐型号核证记录）
纪律：只读扫描不改底表；1.5s 限速；每 200 条落盘断点。
用法：
  python eidc_full_sweep_v1.py --run --limit 200    # 从头跑 200 条
  python eidc_full_sweep_v1.py --run --limit 200 --offset 200  # 断点续跑
  python eidc_full_sweep_v1.py --summary              # 汇总已有结果
"""
import json, os, sys, time, urllib.parse, urllib.request, ssl
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
WB = os.path.join(HERE, 'NEV公告参数汇总表_合并版（341~410批）.xlsx')
OUT = os.path.join(HERE, 'audit-output', 'eidc_full_sweep_results.json')
BASE_API = 'https://service.miit-eidc.org.cn/miitxxgk/gonggao/xxgk/doCpQueryByeOne'
UA = {'User-Agent': 'Mozilla/5.0', 'Referer': 'https://service.miit-eidc.org.cn/miitxxgk/gonggao/xxgk/index?querylb=qy'}


def load_models_by_batch():
    """按批次分组返回全部唯一型号，由近及远排序"""
    wb = openpyxl.load_workbook(WB, read_only=True)
    ws = wb['NEV公告参数汇总']
    hdr = [str(c.value or '').strip() for c in next(ws.iter_rows(min_row=1, max_row=1))]
    hi = {h: i for i, h in enumerate(hdr)}
    by_batch = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        model = str(row[hi['产品型号']] or '').strip()
        batch = str(row[hi['批次']] or '').strip()
        if model and len(model) >= 4:
            by_batch.setdefault(batch, set()).add(model)
    wb.close()
    return by_batch  # {batch_str: set(model_codes)}


def query_eidc(model_code):
    """查 EIDC 单型号 → 官方记录列表"""
    url = (f'https://service.miit-eidc.org.cn/miitxxgk/gonggao/xxgk/doCpQueryByeOne'
           f'?querylb=cp&querydata={urllib.parse.quote(model_code)}&pageSize=20&pageNum=1')
    req = urllib.request.Request(url, headers=UA)
    ctx = ssl.create_default_context()
    d = json.loads(urllib.request.urlopen(req, timeout=20, context=ctx).read().decode('utf-8', 'ignore'))
    rows = d.get('cpList') or []
    total = d.get('countResult', {}).get('total', 0)
    results = []
    for r in rows:
        results.append({
            'eidc_pc': r.get('pc', ''),
            'clxh': r.get('clxh', ''),
            'clmc': r.get('clmc', ''),
            'qymc': r.get('qymc', ''),
            'cpsb': r.get('cpsb', ''),
        })
    return results, total


def main():
    by_batch = load_models_by_batch()
    # 由近及远排序
    sorted_batches = sorted(by_batch.keys(), key=lambda x: -int(x) if x.isdigit() else 0)
    all_models = []
    for b in sorted_batches:
        for m in sorted(by_batch[b]):
            all_models.append({'batch': b, 'model': m})
    total = len(all_models)

    results = []
    if os.path.exists(OUT):
        results = json.load(open(OUT, encoding='utf-8'))
    done = {r['model'] for r in results}
    todo = [m for m in all_models if m['model'] not in done]

    limit = int(sys.argv[sys.argv.index('--limit') + 1]) if '--limit' in sys.argv else 200
    offset = int(sys.argv[sys.argv.index('--offset') + 1]) if '--offset' in sys.argv else 0
    todo = todo[offset:offset + limit]

    print(f'全库型号: {total} | 已完成: {len(done)} | 本轮: {len(todo)}')
    match = mismatch = not_found = 0
    for i, item in enumerate(todo):
        model, wb_batch = item['model'], item['batch']
        try:
            eidc_rows, total_found = query_eidc(model)
        except Exception as e:
            results.append({'model': model, 'wb_batch': wb_batch, 'status': f'err:{type(e).__name__}'})
            time.sleep(3)
            continue
        time.sleep(1.5)

        if not eidc_rows:
            results.append({'model': model, 'wb_batch': wb_batch, 'status': 'not_found',
                            'note': 'EIDC 系统未收录该型号'})
            not_found += 1
        else:
            eidc_pcs = sorted({r['eidc_pc'] for r in eidc_rows if r['eidc_pc']})
            wb_ok = wb_batch in eidc_pcs
            latest_pc = max(eidc_pcs, key=lambda x: int(x) if x.isdigit() else 0)
            entry = {
                'model': model, 'wb_batch': wb_batch,
                'eidc_batches': eidc_pcs, 'eidc_latest_pc': latest_pc,
                'batch_match': wb_batch in eidc_pcs,
                'is_change_extension': len(eidc_pcs) > 1,
                'status': 'match' if wb_ok else 'batch_diff',
                'detail': eidc_rows[0],
            }
            if wb_ok:
                match += 1
            else:
                mismatch += 1
            results.append(entry)

        if (i + 1) % 50 == 0:
            json.dump(results, open(OUT, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
            print(f'  进度 {i+1}/{len(todo)} | match {match} | diff {mismatch} | not_found {not_found}')

    json.dump(results, open(OUT, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    print(f'本轮完成: {len(todo)} 条 | match {match} | diff {mismatch} | not_found {not_found}')
    print(f'累计结果: {len(results)} 条 → {OUT}')


if __name__ == '__main__':
    main()
