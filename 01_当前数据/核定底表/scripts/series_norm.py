# -*- coding: utf-8 -*-
"""车系名归一化（12 家全量）。自 15 _ress_oem_bubble_data.norm_series 扩展：

规则链：取首别名 → 去空白/零宽字符 → 剥品牌前缀 → 拉丁转大写 → 极氪补零 →
短代号补板块前缀 → 别名台账 canon 合并。

别名台账 = data/series_alias_master.json（单一事实源；缺失时用内置种子）。
"""
import json
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import ALIAS_JSON

# 品牌前缀（长前缀优先）
PREFIX = {
    'byd': ['比亚迪', '腾势', '方程豹', '仰望', '领汇'],
    'geely': ['沃尔沃', '极星', 'smart', '精灵', '领克', '极氪', '吉利', '银河',
              '几何', '睿蓝', '雷达', '远程', '帝豪'],
    'nio': ['蔚来', '乐道', '萤火虫'],
    'xiaomi': ['小米'],
    'tesla': ['特斯拉'],
    'liauto': ['理想'],
    'xpeng': ['小鹏'],
    'gwm': ['哈弗', '魏牌', '坦克', '欧拉', '长城'],
    'changan': ['长安', '深蓝', '启源', '阿维塔'],
    'gac': ['广汽', '埃安', '传祺', '昊铂', '合创'],
    'chery': ['奇瑞', '星途', '捷途', 'iCAR', '智界', '开瑞'],
    'saic': ['上汽', '荣威', '名爵', '五菱', '宝骏', '大通', '飞凡', '智己', 'MILA'],
}

# 短代号车系需补板块前缀（「7X」→「极氪7X」），否则跨品牌无法区分
SHORT_CODE = re.compile(r'^[0-9A-Z]{1,5}(GT|X|PLUS|MAX)?$')

# 内置种子（与 data/series_alias_master.json 初始内容一致；文件存在则以文件为准）
SEED_ALIAS = {
    # 比亚迪
    '宋PLUSEV': '宋PLUS', '比亚迪SONGPLUS': '宋PLUS', 'SONGPLUS': '宋PLUS',
    '比亚迪SONG': '宋PLUS', '宋PROEV': '宋PRO', '宋PRODM-i': '宋PRO',
    '海豹05EV': '海豹06', '海豹05': '海豹06',
    '元PLUSEV': '元PLUS', '元PRO': '元',
    '秦PLUSEV': '秦PLUS', '秦Plus': '秦PLUS',
    # 吉利系
    '吉利E5EV': '吉利E5', '银河E5': '吉利E5',
    '吉利E8EV': '吉利E8', '银河E8': '吉利E8',
    '银河星耀8': '星耀8', '星耀8EV': '星耀8',
    '银河A7': '吉利A7', '吉利A7EV': '吉利A7',
    '极氪001FR': '极氪001',
    '009光辉': '极氪009',
    '领克10': '领克Z10',
    '星愿UP': '星愿', '星愿PRO': '星愿', '星愿STAR': '星愿',
    '元EV': '元', '摩卡PHEV': '摩卡',
    # 特斯拉（空格已去、大写后归并到公开名）
    'MODEL3': 'Model 3', 'MODEL3高性能版': 'Model 3', 'MODEL3P': 'Model 3',
    'MODELY': 'Model Y', 'MODELY高性能版': 'Model Y', 'MODELYL': 'Model Y L',
    # 奇瑞
    '奇瑞EQ1': '小蚂蚁', 'EQ1': '小蚂蚁', 'EQ5': '奇瑞eQ5',
    # 理想
    'ONE': '理想ONE',
    # 广汽
    '埃安AIONS': 'AION S', '埃安AIONY': 'AION Y', '埃安AIONV': 'AION V',
    '埃安S': 'AION S', '埃安Y': 'AION Y', '埃安V': 'AION V',
    '埃安UT': 'AION UT',
    # 第二批（2026-10-25 取证 miss 分析）
    '欧拉R1': '黑猫', '欧拉R2': '白猫', '欧拉ORA5': '欧拉5',
    '深蓝S07EV': '深蓝S07', '启源A06E': '启源A06',
    'L9LIVIS': '理想L9', 'YU7GT': 'YU7',
    '睿行EM60Ⅱ': '睿行EM60', '睿行EM80Ⅱ': '睿行EM80',
    '埃安AIONYPLUS': 'AION Y', '埃安AIONVPLUS': 'AION V',
    # 第三批（2026-10-25 原始行反查身份）
    '奇瑞R7': '智界R7', '长城500': '坦克500', '长城400': '坦克400',
    '#1': '精灵#1', 'smart1': '精灵#1', '广汽AIONS': 'AION S',
    '传祺AIONV': 'AION V', '广汽AIONLX': 'AION LX', '埃安SPLUS': 'AION S',
    '如歌': '枭龙', '哈弗HX': '哈弗H10', 'XSPACE': '翼真L380',
}

_ALIAS = None


def load_alias():
    global _ALIAS
    if _ALIAS is not None:
        return _ALIAS
    if os.path.exists(ALIAS_JSON):
        _ALIAS = json.load(open(ALIAS_JSON, encoding='utf-8'))
    else:
        _ALIAS = dict(SEED_ALIAS)
    return _ALIAS


def save_alias(mapping):
    """把别名台账写回 data/series_alias_master.json（按主名排序，便于 diff 审阅）。"""
    os.makedirs(os.path.dirname(ALIAS_JSON), exist_ok=True)
    ordered = dict(sorted(mapping.items(), key=lambda kv: (kv[1], kv[0])))
    with open(ALIAS_JSON, 'w', encoding='utf-8') as f:
        json.dump(ordered, f, ensure_ascii=False, indent=1)
    global _ALIAS
    _ALIAS = dict(ordered)


def canon(s):
    """别名合并后的主名。"""
    return load_alias().get(s, s)


def norm_series(s, key, board):
    raw = str(s or '').strip()
    raw = re.split(r'[,，、/;；]', raw)[0]                       # 多别名拼接取首个
    t = re.sub(r'[\s\u200b\u200c\u200d\ufeff.·]+', '', raw)     # 去空白/零宽/点号
    for p in PREFIX.get(key, []):
        if len(t) > len(p) and t.upper().startswith(p.upper()):
            t = t[len(p):]
            break
    t = t.upper()
    if board == '极氪' and re.fullmatch(r'\d{1,2}', t):         # 「9」→「极氪009」
        return canon(f'极氪{t.zfill(3)}')
    if SHORT_CODE.match(t) and board and board != '其他':
        t = f'{board}{t}'
    return canon(t) if t else '未命名'


if __name__ == '__main__':
    if not os.path.exists(ALIAS_JSON):
        save_alias(SEED_ALIAS)
        print('初始化别名台账 →', ALIAS_JSON, len(SEED_ALIAS), '条')
    else:
        a = load_alias()
        print('别名台账已存在', len(a), '条')
    # 快速自检
    tests = [('比亚迪海豚,比亚迪EA1', 'byd', '比亚迪'), ('001', 'geely', '极氪'),
             ('Model 3', 'tesla', '特斯拉'), ('One', 'liauto', '理想'),
             ('银河E5', 'geely', '吉利'), ('AION S', 'gac', '埃安')]
    for s, k, b in tests:
        print(f'  {s!r:28s} [{k}/{b}] -> {norm_series(s, k, b)}')
