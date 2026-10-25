# -*- coding: utf-8 -*-
"""Build structured scan data for MIIT announcement batches 411 and 412.

411 is a formal announcement published on 2026-10-16 (document date 2026-10-15).
412 is a public consultation published on 2026-10-16 and is not yet a formal
announcement. The script keeps that distinction explicit in every output record.
"""

from __future__ import annotations

import json
import re
import time
import urllib.parse
import urllib.request
from pathlib import Path
from typing import Any

from lxml import html


ROOT = Path(__file__).resolve().parent
RAW = ROOT / "raw"
DATA = ROOT / "data"
REPO = ROOT.parent / "01-主仓库" / "TA-Analysis"
B411_ARCHIVE = ROOT.parent / "02-数据底表" / "data" / "nev-announcements" / "raw" / "b411"

LIST_API = "https://www.miit.gov.cn/api-gateway/jpaas-publish-server/front/page/build/unit"
LIST_QUERY = {
    "parseType": "buildstatic",
    "webId": "b3eba6883f9240e2b51025f690afbae8",
    "tplSetId": "9a9a7b87a4444169bdef99ff1f84e1aa",
    "pageType": "column",
    "tagId": "信息列表",
    "editType": "null",
    "pageId": "16020d9000724867b6348813f902e0ef",
}
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
OPENER = urllib.request.build_opener(urllib.request.ProxyHandler({}))

PASSENGER_MARKERS = (
    "轿车",
    "乘用车",
    "越野车",
)
EXCLUDE_MARKERS = (
    "客车",
    "货车",
    "载货",
    "专用车",
    "垃圾车",
    "牵引车",
    "摩托车",
    "三轮",
    "底盘",
    "消防车",
    "旅居车",
    "救护车",
    "巡逻车",
    "宿营车",
    "工程车",
    "作业车",
    "教练车",
    "检测车",
    "殡仪车",
    "运钞车",
)
NEV_MARKERS = (
    "纯电动",
    "换电式纯电动",
    "插电式混合动力",
    "插电式增程混合动力",
    "增程式",
    "燃料电池",
)


def clean(value: Any) -> str | None:
    if value is None:
        return None
    text = " ".join(str(value).replace("\xa0", " ").split()).strip()
    return text or None


def normalize_model(value: Any) -> str:
    return re.sub(r"\s+", "", clean(value) or "").upper()


def fetch_bytes(url: str, params: dict[str, str] | None = None, timeout: int = 60) -> bytes:
    if params:
        url = url + "?" + urllib.parse.urlencode(params)
    req = urllib.request.Request(url, headers={"User-Agent": UA})
    with OPENER.open(req, timeout=timeout) as response:
        return response.read()


def fetch_json(url: str, params: dict[str, str] | None = None, timeout: int = 60) -> Any:
    return json.loads(fetch_bytes(url, params, timeout).decode("utf-8", "replace"))


def parse_list(batch: int, page_id: str) -> list[dict[str, Any]]:
    records: list[dict[str, Any]] = []
    seen: set[str] = set()
    page = 1
    while True:
        search = json.dumps(
            {
                "PICI": str(batch),
                "title": "",
                "QYMC": "",
                "CPSB": "",
                "CPMC": "",
                "CPXH": "",
            },
            ensure_ascii=False,
            separators=(",", ":"),
        )
        param_json = json.dumps(
            {"pageNo": page, "pageSize": 100, "loadEnabled": True, "search": search},
            ensure_ascii=False,
            separators=(",", ":"),
        )
        params = dict(LIST_QUERY)
        params["pageId"] = page_id
        params["paramJson"] = param_json
        payload = fetch_json(LIST_API, params)
        content = ((payload.get("data") or {}).get("html") or "")
        doc = html.fromstring(content)
        rows = doc.xpath("//table/tr[td]")
        page_records: list[dict[str, Any]] = []
        for row in rows:
            cells = row.xpath("./td")
            if len(cells) < 6:
                continue
            values = [clean("".join(cell.itertext())) for cell in cells]
            link = row.xpath(".//a/@href")
            detail_url = urllib.parse.urljoin("https://www.miit.gov.cn/", link[0]) if link else None
            record = {
                "批次": batch,
                "批次状态": "正式公告" if batch == 411 else "公示",
                "标题": values[0],
                "企业名称": values[2],
                "产品商标": values[3],
                "产品名称": values[4],
                "产品型号": values[5],
                "详情页URL": detail_url,
                "来源URL": detail_url,
            }
            key = f'{record["企业名称"]}|{record["产品型号"]}|{record["产品名称"]}'
            if key not in seen:
                seen.add(key)
                page_records.append(record)
        if not page_records:
            break
        records.extend(page_records)
        if len(rows) < 100:
            break
        page += 1
        time.sleep(0.3)
    return records


def is_nev_passenger(record: dict[str, Any]) -> bool:
    product_name = clean(record.get("产品名称")) or ""
    if any(marker in product_name for marker in EXCLUDE_MARKERS):
        return False
    if not any(marker in product_name for marker in PASSENGER_MARKERS):
        return False
    return any(marker in product_name for marker in NEV_MARKERS)


def kv_from_table(table: Any) -> dict[str, str]:
    pairs: dict[str, str] = {}
    for row in table.xpath(".//tr"):
        cells = row.xpath("./th|./td")
        i = 0
        while i < len(cells):
            if cells[i].tag.lower() == "th":
                key = clean("".join(cells[i].itertext()))
                span = int(cells[i].get("colspan") or 1)
                value_cells = []
                j = i + 1
                while j < len(cells) and len(value_cells) < span:
                    if cells[j].tag.lower() == "td":
                        value_cells.append(clean("".join(cells[j].itertext())))
                    j += 1
                if key and key not in pairs and value_cells:
                    pairs[key] = value_cells[0] or ""
                i = j
            else:
                i += 1
    return pairs


def parse_dimension(value: str | None) -> tuple[str | None, str | None, str | None]:
    if not value:
        return None, None, None
    match = re.search(
        r"长[:：]\s*([\d.,]+)\s*宽[:：]\s*([\d.,]+)\s*高[:：]\s*([\d.,]+)",
        value,
    )
    if not match:
        return None, None, None
    return match.group(1), match.group(2), match.group(3)


def parse_first_number(value: str | None) -> str | None:
    if not value:
        return None
    match = re.search(r"\d+(?:\.\d+)?", value)
    return match.group(0) if match else None


def parse_power_notes(other: str | None) -> dict[str, str | None]:
    text = clean(other) or ""
    result: dict[str, str | None] = {
        "电池类型_公示其它": None,
        "电芯供应商_公示其它": None,
        "电池包供应商_公示其它": None,
        "电机额定功率_kW_公示其它": None,
        "电机峰值功率_kW_公示其它": None,
        "电机额定功率数列_kW_公示其它": None,
        "电机峰值功率数列_kW_公示其它": None,
        "续航_km_公示其它": None,
        "电池容量_kWh_公示其它": None,
    }
    material = re.search(r"储能装置种类[:：]\s*([^,，;；。]+)", text)
    if material:
        result["电池类型_公示其它"] = clean(material.group(1))
    suppliers = re.search(
        r"储能装置生产企业[:：]\s*([^,，;；。]+?)(?:\(总成\)|（总成）)?\s*/?\s*([^,，;；。]+?)(?:\(单体\)|（单体）)?(?=[,，;；。]|$)",
        text,
    )
    if suppliers:
        result["电池包供应商_公示其它"] = clean(suppliers.group(1))
        result["电芯供应商_公示其它"] = clean(suppliers.group(2))
    power_tail = r"((?:前[:：]?)?\d+(?:\.\d+)?\s*kW(?:\s*/\s*(?:后(?:左|右)?[:：]?)?\d+(?:\.\d+)?\s*kW)*)"
    rated = re.search(
        r"(?:驱动电机)?额定功率(?:\([^)]*\)|（[^）]*）)?\s*[:：为]?\s*" + power_tail,
        text,
        re.I,
    )
    peak = re.search(
        r"(?:驱动电机)?峰值功率(?:\([^)]*\)|（[^）]*）)?\s*[:：为]?\s*" + power_tail,
        text,
        re.I,
    )
    if rated:
        values = re.findall(r"\d+(?:\.\d+)?", rated.group(1))
        if values:
            result["电机额定功率数列_kW_公示其它"] = "/".join(values)
            result["电机额定功率_kW_公示其它"] = "/".join(values)
    if peak:
        values = re.findall(r"\d+(?:\.\d+)?", peak.group(1))
        if values:
            result["电机峰值功率数列_kW_公示其它"] = "/".join(values)
            result["电机峰值功率_kW_公示其它"] = "/".join(values)
    range_match = re.search(r"(?:纯电|CLTC|WLTC)?(?:续驶里程|续航里程)[:：]?\s*([\d/]+)\s*(?:km|公里)", text, re.I)
    if range_match:
        result["续航_km_公示其它"] = range_match.group(1)
    capacity_match = re.search(r"(?:动力蓄电池组总能量|电池容量|储能装置总能量)[:：]?\s*([\d.]+)\s*kWh", text, re.I)
    if capacity_match:
        result["电池容量_kWh_公示其它"] = capacity_match.group(1)
    return result


def component_rows(table: Any) -> list[dict[str, str]]:
    rows = table.xpath(".//tr")
    if not rows:
        return []
    header_cells = rows[0].xpath("./th|./td")
    headers = [clean("".join(cell.itertext())) or "" for cell in header_cells]
    out: list[dict[str, str]] = []
    for row in rows[1:]:
        cells = row.xpath("./td")
        values = [clean("".join(cell.itertext())) or "" for cell in cells]
        if not any(values):
            continue
        rec = {headers[i] if i < len(headers) else f"列{i+1}": values[i] for i in range(len(values))}
        out.append(rec)
    return out


def parse_detail(record: dict[str, Any], html_text: str) -> dict[str, Any]:
    doc = html.fromstring(html_text)
    tables = doc.xpath("//table[contains(@class,'w_table')]")
    tables = tables[:4]
    pairs: dict[str, str] = {}
    for table in tables:
        for key, value in kv_from_table(table).items():
            pairs.setdefault(key, value)

    components: list[dict[str, str]] = []
    for table in tables:
        header = " ".join(table.xpath(".//tr[1]//th/text()"))
        if "发动机型号" in header or "驱动电机型号" in header:
            components.extend(component_rows(table))

    dimensions = pairs.get("外形尺寸(mm)：") or pairs.get("外形尺寸(mm):")
    length, width, height = parse_dimension(dimensions)
    other = pairs.get("其它：") or pairs.get("其它:")
    note_fields = parse_power_notes(other)

    motor_rows = []
    engine_rows = []
    for row in components:
        model = next((v for k, v in row.items() if "发动机型号" in k or "电机型号" in k), "")
        power = next((v for k, v in row.items() if "功率" in k), "")
        supplier = next((v for k, v in row.items() if "企业" in k or "生产企业" in k), "")
        displacement = next((v for k, v in row.items() if "排量" in k), "")
        row_clean = {
            "型号": model,
            "生产企业": supplier,
            "排量": displacement,
            "功率": power,
        }
        is_motor = model.upper().startswith("TZ") or "电机" in model or (displacement in ("-", "") and power)
        is_engine = (not is_motor) and (displacement not in ("", "-") or (model and not model.upper().startswith("TZ")))
        if is_motor and power:
            motor_rows.append(row_clean)
        if is_engine:
            engine_rows.append(row_clean)

    motor_power = note_fields.get("电机峰值功率数列_kW_公示其它")
    rated_power = note_fields.get("电机额定功率数列_kW_公示其它")
    if not motor_power and motor_rows:
        powers = [row["功率"] for row in motor_rows if row["功率"]]
        if powers:
            motor_power = "/".join(powers)
    power_values = [float(value) for value in (motor_power or "").split("/") if value]
    rated_values = [float(value) for value in (rated_power or "").split("/") if value]
    total_power = sum(power_values) if power_values else (sum(rated_values) if rated_values else None)
    front_power = power_values[0] if power_values else None
    rear_power = "/".join(str(int(v)) if v.is_integer() else str(v) for v in power_values[1:]) if len(power_values) > 1 else None
    motor_count = len(power_values) or len(motor_rows) or None
    if motor_count == 1:
        drive_layout = "单电机"
    elif motor_count and motor_count >= 4:
        drive_layout = "四电机"
    elif motor_count == 3:
        drive_layout = "三电机"
    elif motor_count == 2:
        drive_layout = "双电机"
    else:
        drive_layout = None

    result = {
        **record,
        "车长(mm)": length,
        "车宽(mm)": width,
        "车高(mm)": height,
        "轴距(mm)": parse_first_number(pairs.get("轴距(mm)：") or pairs.get("轴距(mm):")),
        "整备质量(kg)": pairs.get("整备质量(kg)：") or pairs.get("整备质量(kg):"),
        "总质量(kg)": pairs.get("总质量(kg)：") or pairs.get("总质量(kg):"),
        "最高车速(km/h)": pairs.get("最高车速(km/h)：") or pairs.get("最高车速(km/h):"),
        "额定载客(人)": pairs.get("额定载客（含驾驶员）（座位数）：") or pairs.get("额定载客(含驾驶员)(座位数):"),
        "燃料种类": pairs.get("燃料种类：") or pairs.get("燃料种类:"),
        "油耗申报值(L/100km)": pairs.get("油耗申报值(L/100km)：") or pairs.get("油耗申报值(L/100km):"),
        "电机功率(kW)_公示详情": motor_power,
        "电机额定功率数列_kW": rated_power,
        "电机峰值功率数列_kW": motor_power,
        "电机总功率_kW": total_power,
        "前电机峰值功率_kW": front_power,
        "后电机峰值功率_kW": rear_power,
        "电机数量": motor_count,
        "驱动形式_电机数量推断": drive_layout,
        "电机组件_公示详情": motor_rows,
        "发动机组件_公示详情": engine_rows,
        **note_fields,
        "公示详情原始其它": other,
    }
    if result.get("产品商标") is None:
        result["产品商标"] = pairs.get("产品商标：")
    if result.get("产品型号") is None:
        result["产品型号"] = pairs.get("产品型号：")
    if result.get("产品名称") is None:
        result["产品名称"] = pairs.get("产品名称：")
    if result.get("企业名称") is None:
        result["企业名称"] = pairs.get("企业名称：")
    return result


def enrich_412_details(records: list[dict[str, Any]]) -> list[dict[str, Any]]:
    enriched: list[dict[str, Any]] = []
    for index, record in enumerate(records, 1):
        url = record.get("详情页URL")
        if not url:
            enriched.append(record)
            continue
        try:
            text = fetch_bytes(url).decode("utf-8", "replace")
            enriched.append(parse_detail(record, text))
            print(f"  412 detail {index}/{len(records)} {record.get('产品型号')}")
        except Exception as exc:
            fallback = dict(record)
            fallback["详情抓取错误"] = f"{type(exc).__name__}: {exc}"
            enriched.append(fallback)
        time.sleep(0.25)
    return enriched


def parse_411_tax_dirs() -> list[dict[str, Any]]:
    vessel_path = RAW / "411_doc2_vessel90.txt"
    purchase_path = RAW / "411_doc3_purchase35.txt"
    vessel_lines = vessel_path.read_text(encoding="utf-8").splitlines()
    purchase_lines = purchase_path.read_text(encoding="utf-8").splitlines()

    def rows_between(lines: list[str], start_marker: str, end_prefixes: tuple[str, ...]) -> list[list[str]]:
        out: list[list[str]] = []
        active = False
        for line in lines:
            stripped = line.strip()
            if start_marker in stripped:
                active = True
                continue
            if active and any(stripped.startswith(prefix) for prefix in end_prefixes):
                if not stripped.endswith("乘用车"):
                    break
            if active and stripped and stripped[0].isdigit() and "\t" in stripped:
                out.append([cell.strip() for cell in stripped.split("\t")])
        return out

    def rows_in_section(
        lines: list[str],
        section_marker: str,
        passenger_marker: str,
        end_markers: tuple[str, ...],
    ) -> list[list[str]]:
        in_section = False
        in_passenger = False
        out: list[list[str]] = []
        for line in lines:
            stripped = line.strip()
            if section_marker in stripped:
                in_section = True
                continue
            if not in_section:
                continue
            if any(stripped.startswith(marker) for marker in end_markers):
                break
            if stripped.startswith(passenger_marker):
                in_passenger = True
                continue
            if in_passenger and stripped.startswith(("（二）", "（三）", "（四）", "（五）")):
                break
            if in_passenger and stripped and stripped[0].isdigit() and "\t" in stripped:
                out.append([cell.strip() for cell in stripped.split("\t")])
        return out

    bev_rows = rows_in_section(purchase_lines, "一、纯电动汽车", "（一）乘用车", ("二、",))
    phev_rows = rows_in_section(purchase_lines, "二、插电式混合动力汽车", "（一）乘用车", ("三、",))
    vessel_phev = rows_in_section(vessel_lines, "二、新能源汽车", "（一）插电式混合动力乘用车", ("（二）",))

    by_model: dict[str, dict[str, Any]] = {}

    def add(record: dict[str, Any], source: str, priority: int) -> None:
        model = normalize_model(record.get("产品型号"))
        if not model:
            return
        current = by_model.get(model)
        if current is None or priority >= current["_priority"]:
            merged = dict(current or {})
            merged.update({k: v for k, v in record.items() if v not in (None, "")})
            merged["_priority"] = priority
            merged["来源清单"] = source
            by_model[model] = merged

    for row in bev_rows:
        if len(row) < 9:
            continue
        add(
            {
                "批次": 411,
                "批次状态": "正式公告",
                "企业名称": row[1],
                "产品型号": row[2],
                "通用名称": row[3],
                "产品名称": row[4],
                "纯电续航里程(km)": row[5],
                "整备质量(kg)": row[6],
                "动力蓄电池组总质量(kg)": row[7],
                "电池容量(kWh)": row[8],
                "来源清单": "减免车辆购置税目录第三十五批（411公告附件3）",
            },
            "PURCHASE_TAX_BEV",
            90,
        )

    for row in phev_rows:
        if len(row) < 11:
            continue
        add(
            {
                "批次": 411,
                "批次状态": "正式公告",
                "企业名称": row[1],
                "产品型号": row[2],
                "通用名称": row[3],
                "产品名称": row[4],
                "纯电续航里程(km)": row[5],
                "综合油耗(L/100km)": row[6],
                "发动机排量(mL)": row[7],
                "整备质量(kg)": row[8],
                "动力蓄电池组总质量(kg)": row[9],
                "电池容量(kWh)": row[10],
            },
            "PURCHASE_TAX_PHEV",
            90,
        )

    for row in vessel_phev:
        if len(row) < 12:
            continue
        add(
            {
                "批次": 411,
                "批次状态": "正式公告",
                "企业名称": row[1],
                "产品商标": row[2],
                "产品型号": row[3],
                "通用名称": row[4],
                "产品名称": "插电式混合动力乘用车",
                "纯电续航里程(km)": row[5],
                "综合油耗(L/100km)": row[6],
                "发动机排量(mL)": row[7],
                "整备质量(kg)": row[8],
                "动力蓄电池组总质量(kg)": row[9],
                "电池容量(kWh)": row[10],
            },
            "VESSEL_TAX_PHEV",
            95,
        )

    return list(by_model.values())


def enrich_411_details(records: list[dict[str, Any]]) -> list[dict[str, Any]]:
    urls = json.loads((B411_ARCHIVE / "detail_urls.json").read_text(encoding="utf-8"))
    extracted = json.loads((B411_ARCHIVE / "details_extracted.json").read_text(encoding="utf-8"))
    cell_fill = json.loads((B411_ARCHIVE / "cell_fill_411_20260918.json").read_text(encoding="utf-8"))
    enriched = []
    for rec in records:
        model = normalize_model(rec.get("产品型号"))
        url = rec.get("详情页URL") or urls.get(model)
        if not url:
            fallback = dict(rec)
            for key, value in extracted.get(model, {}).items():
                if value not in (None, "") and not fallback.get(key):
                    fallback[key] = value
            enriched.append(fallback)
            continue
        try:
            detail = parse_detail(
                {
                    **rec,
                    "详情页URL": url,
                    "来源URL": url,
                },
                fetch_bytes(url).decode("utf-8", "replace"),
            )
            archive = extracted.get(model, {})
            for key, value in {
                "车长(mm)": archive.get("车长(mm)"),
                "整备质量(kg)": archive.get("整备质量(kg)"),
                "轴距(mm)": archive.get("轴距(mm)"),
                "电池类型_正式": archive.get("电池类型"),
                "电芯供应商_正式": archive.get("电芯供应商"),
                "电池包供应商_正式": archive.get("电池包供应商"),
                "发动机型号": archive.get("发动机型号"),
                "发动机生产企业": archive.get("发动机生产企业"),
                "发动机排量(mL)": archive.get("发动机排量(mL)"),
                "发动机功率(kW)": archive.get("发动机功率(kW)"),
            }.items():
                if value not in (None, "") and not detail.get(key):
                    detail[key] = value
            fill = cell_fill.get(model, {})
            if fill.get("cell"):
                detail["电芯供应商_正式"] = fill["cell"]
            if fill.get("pack"):
                detail["电池包供应商_正式"] = fill["pack"]
            # Tax-directory values are authoritative for energy/range fields.
            for key, value in rec.items():
                if value not in (None, ""):
                    detail[key] = value
            enriched.append(detail)
        except Exception as exc:
            fallback = dict(rec)
            fallback["详情页URL"] = url
            fallback["详情抓取错误"] = f"{type(exc).__name__}: {exc}"
            enriched.append(fallback)
        time.sleep(0.2)
    return enriched


def parse_412_tax_draft() -> list[dict[str, Any]]:
    path = RAW / "412_doc4_vessel91_draft.txt"
    lines = path.read_text(encoding="utf-8").splitlines()
    records: list[dict[str, Any]] = []
    active = False
    for line in lines:
        stripped = line.strip()
        if "（一）插电式混合动力乘用车" in stripped:
            active = True
            continue
        if active and stripped.startswith(("（二）", "（三）", "（四）", "三、")):
            break
        if active and stripped and stripped[0].isdigit() and "\t" in stripped:
            row = [cell.strip() for cell in stripped.split("\t")]
            if len(row) < 11:
                continue
            records.append(
                {
                    "批次": 412,
                    "批次状态": "公示",
                    "企业名称": row[1],
                    "产品商标": row[2],
                    "产品型号": row[3],
                    "通用名称": row[4],
                    "产品名称": "插电式混合动力乘用车",
                    "纯电续航里程(km)": row[5],
                    "综合油耗(L/100km)": row[6],
                    "发动机排量(mL)": row[7],
                    "整备质量(kg)": row[8],
                    "动力蓄电池组总质量(kg)": row[9],
                    "电池容量(kWh)": row[10],
                    "来源清单": "车船税目录第九十一批拟发布清单（412公示附件4）",
                }
            )
    return records


def merge_records(base: list[dict[str, Any]], supplement: list[dict[str, Any]]) -> list[dict[str, Any]]:
    by_model: dict[str, dict[str, Any]] = {}
    for rec in base + supplement:
        model = normalize_model(rec.get("产品型号"))
        if not model:
            continue
        current = by_model.setdefault(model, {})
        for key, value in rec.items():
            if value not in (None, "", []):
                if key == "产品名称" and current.get(key) and len(current[key]) > len(str(value)):
                    continue
                current[key] = value
    return list(by_model.values())


def main() -> None:
    DATA.mkdir(parents=True, exist_ok=True)
    print("Fetching 411 and 412 lists...")
    list_411 = parse_list(411, "6e2409267b5f42c798f1d591764caa1a")
    list_412 = parse_list(412, "16020d9000724867b6348813f902e0ef")
    nev_passenger_411 = [record for record in list_411 if is_nev_passenger(record)]
    nev_passenger_412 = [record for record in list_412 if is_nev_passenger(record)]
    print(f"  411 all={len(list_411)} NEV passenger={len(nev_passenger_411)}")
    print(f"  412 all={len(list_412)} NEV passenger={len(nev_passenger_412)}")
    details_412 = enrich_412_details(nev_passenger_412)
    tax_412 = parse_412_tax_draft()
    details_412 = merge_records(details_412, tax_412)

    print("Building 411 official passenger data...")
    tax_411 = parse_411_tax_dirs()
    details_411 = enrich_411_details(merge_records(tax_411, nev_passenger_411))
    print(f"  411 passenger tax rows={len(tax_411)} combined={len(details_411)}")

    outputs = {
        "list_411_all.json": list_411,
        "list_411_nev_passenger.json": nev_passenger_411,
        "list_412_all.json": list_412,
        "list_412_nev_passenger.json": nev_passenger_412,
        "details_412.json": details_412,
        "tax_412_draft.json": tax_412,
        "details_411.json": details_411,
    }
    for name, payload in outputs.items():
        target = DATA / name
        target.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2, default=str),
            encoding="utf-8",
        )
        print(f"Wrote {target}")


if __name__ == "__main__":
    main()
