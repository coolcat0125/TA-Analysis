/**
 * Build the auditable TA-Analysis parameter-enrichment workbook.
 * Copyright © 2026 David YE
 * SPDX-License-Identifier: MIT
 */
import fs from "node:fs/promises";
import crypto from "node:crypto";
import { FileBlob, SpreadsheetFile } from "./vendor/artifact-tool-shim.mjs";

const COPYRIGHT = "Copyright © 2026 David YE";
const VERSION = "v3.4.6-candidate";
const SOURCE = process.argv[2] || "source_341_409.xlsx";
const RECORDS_PATH = process.argv[3] || "analysis-output/analysis_records.json";
const EXTERNAL_EVIDENCE_PATH = process.argv[4] || "external_evidence.json";
const OUTPUT_DIR = process.argv[5] || "deliverables";
const OUTPUT_FILE = `${OUTPUT_DIR}/NEV公告参数汇总表_证据补全候选（341~409批）.xlsx`;
const PREVIEW_FILE = `${OUTPUT_DIR}/enrichment_summary_preview.png`;
const INVALID_TEXT = /^(?:请提供|抱歉|无法|未提供|未找到|未提取|无有效信息|未直接提供|没有提供|正文中未|原文中未)/;
const MISSING = new Set(["", "/", "-", "--", "N/A", "NA", "null", "None"]);
const COLORS = {
  navy: "#17324D",
  teal: "#0F766E",
  blue: "#2563EB",
  paleBlue: "#EAF2F8",
  paleTeal: "#E8F5F2",
  paleAmber: "#FFF4D6",
  paleRed: "#FDECEC",
  gray: "#F3F4F6",
  white: "#FFFFFF",
  text: "#1F2937",
  border: "#D1D5DB",
};

function isMissing(value) {
  if (value === null || value === undefined) return true;
  const text = String(value).trim();
  return MISSING.has(text) || INVALID_TEXT.test(text);
}

function parseTorque(value) {
  if (isMissing(value)) return { state: "blank", torqueNm: [] };
  const text = String(value).trim().replace(/[：]/g, ":").replace(/[；，]/g, " ");
  const values = [];
  for (const match of text.matchAll(/\/\s*(-?\d+(?:\.\d+)?)/g)) values.push(Number(match[1]));
  for (const match of text.matchAll(/(-?\d+(?:\.\d+)?)\s*N\s*[·.\-]?\s*m\b/gi)) values.push(Number(match[1]));
  if (!text.includes("/")) {
    for (const match of text.matchAll(/(?:^|\s)(?:F|R|前|后)\s*:\s*(-?\d+(?:\.\d+)?)/gi)) values.push(Number(match[1]));
  }
  const unique = [...new Set(values.filter((number) => Number.isFinite(number) && number > 0 && number < 3000))];
  if (unique.length) return { state: "parseable", torqueNm: unique };
  if (text.includes("/")) return { state: "power_only", torqueNm: [] };
  return { state: "ambiguous", torqueNm: [] };
}

function excelColumn(number) {
  let value = number;
  let output = "";
  while (value > 0) {
    value -= 1;
    output = String.fromCharCode(65 + (value % 26)) + output;
    value = Math.floor(value / 26);
  }
  return output;
}

function firstNumericR1C1(offset) {
  const ref = `RC[${offset}]`;
  return `IF(ISNUMBER(${ref}),${ref},IFERROR(VALUE(${ref}),VALUE(LEFT(TRIM(${ref}),FIND(\"|\",SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(TRIM(${ref}),\"/\",\"|\"),\"(\",\"|\"),\"（\",\"|\"),\";\",\"|\"),\"；\",\"|\")&\"|\")-1))))`;
}

function toNumber(value) {
  if (isMissing(value)) return null;
  if (typeof value === "number" && Number.isFinite(value)) return value;
  const match = String(value).replace(/,/g, "").match(/-?\d+(?:\.\d+)?/);
  return match ? Number(match[0]) : null;
}

function modelTokens(value) {
  const matches = String(value ?? "").toUpperCase().match(/[A-Z0-9]{5,}/g) || [];
  return [...new Set(matches.filter((token) => /[A-Z]/.test(token) && /\d/.test(token)))];
}

function numericTokens(...values) {
  const numbers = [];
  for (const value of values) {
    for (const match of String(value ?? "").replace(/,/g, " ").matchAll(/-?\d+(?:\.\d+)?/g)) {
      const number = Number(match[0]);
      if (Number.isFinite(number)) numbers.push(number);
    }
  }
  return numbers;
}

function powerCompatible(target, evidence) {
  const applicableBrands = (evidence.applicable_brands || []).map((value) => String(value).trim());
  if (applicableBrands.length && !applicableBrands.includes(String(target.brand ?? "").trim())) return false;
  const applicableModels = (evidence.applicable_product_models || []).map((value) => String(value).trim().toUpperCase());
  if (applicableModels.length && !applicableModels.includes(String(target.product_model ?? "").trim().toUpperCase())) return false;
  const expected = (evidence.expected_peak_kw || []).map(Number).filter(Number.isFinite);
  if (!expected.length) return true;
  const observed = numericTokens(target.motor_peak_power_kw, target.motor_total_power_kw, target.current_value);
  if (!observed.length) return true;
  return expected.some((expectedValue) => observed.some((observedValue) => Math.abs(expectedValue - observedValue) < 0.01));
}

function cellValue(value) {
  if (value === undefined) return null;
  if (typeof value === "string" && value.length > 30000) return `${value.slice(0, 29970)}…[truncated]`;
  return value;
}

function recordMatrix(records, fields) {
  return records.map((record) => fields.map((field) => cellValue(record[field])));
}

function styleHeader(range, fill = COLORS.navy) {
  range.format = {
    fill,
    font: { bold: true, color: COLORS.white },
    wrapText: true,
    verticalAlignment: "center",
    horizontalAlignment: "center",
    borders: { preset: "all", style: "thin", color: COLORS.border },
  };
}

function addDataSheet(workbook, name, title, note, fields, labels, records, tableName, widths = {}) {
  const sheet = workbook.worksheets.add(name);
  const lastColumn = excelColumn(fields.length);
  sheet.showGridLines = false;
  sheet.getRange(`A1:${lastColumn}1`).merge();
  sheet.getRange("A1").values = [[title]];
  sheet.getRange(`A1:${lastColumn}1`).format = {
    fill: COLORS.navy,
    font: { bold: true, color: COLORS.white, size: 16 },
    verticalAlignment: "center",
  };
  sheet.getRange(`A2:${lastColumn}2`).merge();
  sheet.getRange("A2").values = [[COPYRIGHT]];
  sheet.getRange(`A2:${lastColumn}2`).format = { fill: COLORS.paleBlue, font: { italic: true, color: COLORS.navy } };
  sheet.getRange(`A3:${lastColumn}3`).merge();
  sheet.getRange("A3").values = [[note]];
  sheet.getRange(`A3:${lastColumn}3`).format = { fill: COLORS.gray, font: { color: COLORS.text }, wrapText: true };
  sheet.getRange(`A5:${lastColumn}${records.length + 5}`).values = [labels, ...recordMatrix(records, fields)];
  styleHeader(sheet.getRange(`A5:${lastColumn}5`), COLORS.teal);
  if (records.length) {
    const dataRange = sheet.getRange(`A6:${lastColumn}${records.length + 5}`);
    dataRange.format.borders = { preset: "all", style: "thin", color: COLORS.border };
    dataRange.format.verticalAlignment = "top";
    dataRange.format.wrapText = true;
    const table = sheet.tables.add(`A5:${lastColumn}${records.length + 5}`, true, tableName);
    table.style = "TableStyleMedium2";
    table.showFilterButton = true;
  }
  sheet.freezePanes.freezeRows(5);
  for (let index = 0; index < fields.length; index += 1) {
    sheet.getRange(`${excelColumn(index + 1)}:${excelColumn(index + 1)}`).format.columnWidth = widths[fields[index]] || 16;
  }
  sheet.getRange("1:1").format.rowHeight = 28;
  sheet.getRange("3:3").format.rowHeight = 38;
  return sheet;
}

const sourceBytes = await fs.readFile(SOURCE);
const sourceSha256 = crypto.createHash("sha256").update(sourceBytes).digest("hex");
const records = JSON.parse(await fs.readFile(RECORDS_PATH, "utf8"));
if (records.copyright !== COPYRIGHT) throw new Error("Copyright marker mismatch in analysis records");
const externalEvidence = JSON.parse(await fs.readFile(EXTERNAL_EVIDENCE_PATH, "utf8"));
if (externalEvidence.copyright !== COPYRIGHT) throw new Error("Copyright marker mismatch in external evidence");
const motorEvidenceByModel = new Map((externalEvidence.motor_evidence || []).map((item) => [item.motor_model.toUpperCase(), item]));
const batteryEvidenceByModel = new Map((externalEvidence.battery_evidence || []).map((item) => [item.product_model.toUpperCase(), item]));

const blob = await FileBlob.load(SOURCE);
const workbook = await SpreadsheetFile.importXlsx(blob);
const main = workbook.worksheets.getItem("NEV公告参数汇总");
const mainValues = main.getUsedRange(true).values.map((row) => [...row]);
const originalHeaders = mainValues[0].map((value) => String(value ?? "").trim());
const headerIndex = new Map(originalHeaders.map((header, index) => [header, index]));
const newHeaders = [
  "电机扭矩解析值(N·m)",
  "电机扭矩补全候选(N·m)",
  "电机扭矩证据状态",
  "审计后百公里电耗(kWh/100km)",
  "电池组估算质量(kg，仅80±5kWh)",
  "电池组质量状态",
  "数据质量/补全状态",
  "补全版本",
];
mainValues[0].push(...newHeaders);
for (let rowIndex = 1; rowIndex < mainValues.length; rowIndex += 1) mainValues[rowIndex].push(...Array(newHeaders.length).fill(null));
const allHeaders = mainValues[0].map((value) => String(value ?? "").trim());
const allHeaderIndex = new Map(allHeaders.map((header, index) => [header, index]));
const auditRows = [];
const rowFlags = new Map();
const flag = (excelRow, value) => {
  if (!rowFlags.has(excelRow)) rowFlags.set(excelRow, new Set());
  rowFlags.get(excelRow).add(value);
};

for (const item of records.contamination_records) {
  const rowIndex = item.excel_row - 1;
  const columnIndex = headerIndex.get(item.field);
  if (rowIndex < 1 || columnIndex === undefined) continue;
  const oldValue = mainValues[rowIndex][columnIndex];
  if (INVALID_TEXT.test(String(oldValue ?? "").trim())) {
    mainValues[rowIndex][columnIndex] = null;
    flag(item.excel_row, "污染值已清理");
    auditRows.push({
      excel_row: item.excel_row, batch: item.batch, product_model: item.product_model, field: item.field,
      old_value: oldValue, new_value: "", action: "清理污染占位符", evidence_tier: "DQ",
      method: "规则识别非数据型模型回复文本", source_reference: `原始单元格 ${excelColumn(columnIndex + 1)}${item.excel_row}`,
      confidence: "high", status: "applied", copyright: COPYRIGHT,
    });
  }
}

for (const item of records.same_model_candidates) {
  const rowIndex = item.excel_row - 1;
  const columnIndex = headerIndex.get(item.field);
  if (rowIndex < 1 || columnIndex === undefined || !isMissing(mainValues[rowIndex][columnIndex])) continue;
  const oldValue = mainValues[rowIndex][columnIndex];
  mainValues[rowIndex][columnIndex] = cellValue(item.candidate_value);
  flag(item.excel_row, "同型号候选已写入");
  auditRows.push({
    excel_row: item.excel_row, batch: item.batch, product_model: item.product_model, field: item.field,
    old_value: oldValue, new_value: item.candidate_value, action: "补全缺失值", evidence_tier: item.evidence_tier,
    method: item.method, source_reference: `Excel行 ${item.source_rows}; 批次 ${item.source_batches}`,
    confidence: item.confidence, status: "applied_candidate", copyright: COPYRIGHT,
  });
}

const externalMotorByRow = new Map();
for (const target of records.torque_records) {
  const enrichedRow = mainValues[target.excel_row - 1] || [];
  const enrichedTarget = {
    ...target,
    brand: enrichedRow[headerIndex.get("产品商标")] || target.brand,
    motor_model: enrichedRow[headerIndex.get("电机型号")] || target.motor_model,
    motor_peak_power_kw: enrichedRow[headerIndex.get("电机峰值功率(kW)")] || target.motor_peak_power_kw,
    motor_total_power_kw: enrichedRow[headerIndex.get("电机总功率(kW)")] || target.motor_total_power_kw,
    current_value: enrichedRow[headerIndex.get("电机功率/扭矩")] || target.current_value,
  };
  const models = modelTokens(enrichedTarget.motor_model);
  const matched = models
    .map((model) => motorEvidenceByModel.get(model))
    .filter((evidence) => evidence && powerCompatible(enrichedTarget, evidence));
  if (!matched.length) continue;
  const matchedModels = new Set(matched.map((item) => item.motor_model.toUpperCase()));
  const complete = models.length === 1 || (models.length > 1 && models.every((model) => matchedModels.has(model)));
  const value = complete && matched.length === 1
    ? matched[0].value_nm
    : matched.map((item) => `${item.motor_model}:${item.value_nm}`).join("；");
  externalMotorByRow.set(target.excel_row, { models, matched, complete, value, enrichedTarget });
}

const torqueRecordByRow = new Map(records.torque_records.map((item) => [item.excel_row, item]));
const torqueOriginalIndex = headerIndex.get("电机功率/扭矩");
for (let excelRow = 2; excelRow <= mainValues.length; excelRow += 1) {
  const rowIndex = excelRow - 1;
  const parsed = parseTorque(mainValues[rowIndex][torqueOriginalIndex]);
  const target = torqueRecordByRow.get(excelRow);
  let parsedValue = "";
  let candidateValue = "";
  let status = "";
  if (parsed.state === "parseable") {
    parsedValue = parsed.torqueNm.join("|");
    status = "源字段可解析";
  } else if (externalMotorByRow.has(excelRow)) {
    const external = externalMotorByRow.get(excelRow);
    const tiers = [...new Set(external.matched.map((item) => item.evidence_tier))];
    candidateValue = external.value;
    status = `${tiers.join("/")}外部${external.complete ? "候选" : "局部候选"}`;
    flag(excelRow, external.complete ? "扭矩外部候选" : "扭矩外部局部候选");
    auditRows.push({
      excel_row: excelRow, batch: target.batch, product_model: target.product_model, field: "电机扭矩(N·m)",
      old_value: target.current_value, new_value: candidateValue, action: external.complete ? "生成外部候选，不覆盖原字段" : "生成多电机局部候选，不覆盖原字段",
      evidence_tier: tiers.join("/"),
      method: external.matched.map((item) => item.method).join("；"),
      source_reference: external.matched.map((item) => item.source_url).join("；"),
      confidence: external.matched.every((item) => item.confidence === "high") ? "high" : "medium",
      status: external.complete ? "external_candidate" : "external_partial_candidate", copyright: COPYRIGHT,
    });
  } else if (target?.preferred_candidate_nm) {
    candidateValue = target.preferred_candidate_nm;
    status = `C级候选：${target.preferred_method}`;
    flag(excelRow, "扭矩内部候选");
    auditRows.push({
      excel_row: excelRow, batch: target.batch, product_model: target.product_model, field: "电机扭矩(N·m)",
      old_value: target.current_value, new_value: candidateValue, action: "生成候选，不覆盖原字段", evidence_tier: "C-internal",
      method: target.preferred_method,
      source_reference: target.motor_model_source_rows || target.same_model_source_rows,
      confidence: target.confidence || "medium", status: "candidate", copyright: COPYRIGHT,
    });
  } else if (parsed.state === "blank") status = "待OCR/外部证据：原字段为空";
  else if (parsed.state === "power_only") status = "待OCR/外部证据：仅有功率";
  else status = "待OCR/外部证据：数值语义歧义";
  mainValues[rowIndex][allHeaderIndex.get("电机扭矩解析值(N·m)")] = parsedValue;
  mainValues[rowIndex][allHeaderIndex.get("电机扭矩补全候选(N·m)")] = candidateValue;
  mainValues[rowIndex][allHeaderIndex.get("电机扭矩证据状态")] = status;
}

const consumptionFormulaCandidates = [];
for (let excelRow = 2; excelRow <= mainValues.length; excelRow += 1) {
  const row = mainValues[excelRow - 1];
  const current = row[headerIndex.get("百公里电耗(kWh/100km)")];
  const capacity = toNumber(row[headerIndex.get("电池容量(kWh)")]);
  const rangeKm = toNumber(row[headerIndex.get("纯电续航里程(km)")]);
  if (!isMissing(current) || !capacity || !rangeKm) continue;
  const candidate = capacity * 100 / rangeKm;
  if (candidate < 3 || candidate > 40) continue;
  const item = {
    excel_row: excelRow,
    batch: row[headerIndex.get("批次")],
    product_model: row[headerIndex.get("产品型号")],
    candidate_kwh_100km: Number(candidate.toFixed(3)),
    evidence_tier: "C-formula",
    method: "同一行电池容量×100÷纯电续航",
    confidence: "medium",
  };
  consumptionFormulaCandidates.push(item);
  auditRows.push({
    excel_row: item.excel_row, batch: item.batch, product_model: item.product_model, field: "审计后百公里电耗(kWh/100km)",
    old_value: "", new_value: item.candidate_kwh_100km, action: "公式补全", evidence_tier: item.evidence_tier,
    method: item.method, source_reference: `同一行 O${item.excel_row}/N${item.excel_row}`,
    confidence: item.confidence, status: "formula", copyright: COPYRIGHT,
  });
  flag(item.excel_row, "电耗公式可补");
}

for (const item of records.battery_80kwh_records.filter((record) => record.derived_pack_mass_kg !== "")) {
  auditRows.push({
    excel_row: item.excel_row, batch: item.batch, product_model: item.product_model, field: "电池组估算质量(kg，仅80±5kWh)",
    old_value: "", new_value: item.derived_pack_mass_kg, action: "公式推导", evidence_tier: "C-formula",
    method: "同一行电池容量×1000÷系统能量密度", source_reference: `同一行 O${item.excel_row}/Q${item.excel_row}`,
    confidence: "medium", status: "非称重实测", copyright: COPYRIGHT,
  });
  flag(item.excel_row, "80kWh质量可推导");
}

const batterySpecialRecords = records.battery_80kwh_records.map((item) => {
  const evidence = batteryEvidenceByModel.get(String(item.product_model ?? "").toUpperCase());
  if (!evidence) {
    return {
      ...item, cell_supplier: "", battery_evidence_tier: "", battery_source_title: "", battery_source_url: "",
      battery_announcement_date: "", battery_published_date: "", battery_review_status: "pending_direct_evidence",
    };
  }
  flag(item.excel_row, "电池供应商外部候选");
  for (const [field, newValue] of [["电芯生产企业", evidence.cell_supplier], ["电池总成生产企业", evidence.pack_supplier]]) {
    if (!newValue) continue;
    auditRows.push({
      excel_row: item.excel_row, batch: item.batch, product_model: item.product_model, field,
      old_value: "", new_value: newValue, action: "生成外部候选，不传播到近似型号", evidence_tier: evidence.evidence_tier,
      method: evidence.method, source_reference: evidence.source_url,
      confidence: evidence.confidence, status: "external_candidate", copyright: COPYRIGHT,
    });
  }
  return {
    ...item,
    cell_supplier: evidence.cell_supplier,
    pack_supplier: evidence.pack_supplier,
    battery_evidence_tier: evidence.evidence_tier,
    battery_source_title: evidence.source_title,
    battery_source_url: evidence.source_url,
    battery_announcement_date: evidence.announcement_effective_date,
    battery_published_date: evidence.published_date,
    battery_review_status: "exact_product_model_candidate",
  };
});

for (let excelRow = 2; excelRow <= mainValues.length; excelRow += 1) {
  const rowIndex = excelRow - 1;
  mainValues[rowIndex][allHeaderIndex.get("数据质量/补全状态")] = [...(rowFlags.get(excelRow) || [])].join("；");
  mainValues[rowIndex][allHeaderIndex.get("补全版本")] = VERSION;
}

const mainLastColumn = excelColumn(allHeaders.length);
main.getRange(`A1:${mainLastColumn}${mainValues.length}`).values = mainValues;
styleHeader(main.getRange(`A1:${mainLastColumn}1`), COLORS.navy);
main.freezePanes.freezeRows(1);
main.freezePanes.freezeColumns(2);
main.getRange(`AE1:${mainLastColumn}${mainValues.length}`).format.borders = { preset: "all", style: "thin", color: COLORS.border };
main.getRange(`AE2:${mainLastColumn}${mainValues.length}`).format.wrapText = true;
main.getRange("AE:AG").format.columnWidth = 24;
main.getRange("AH:AJ").format.columnWidth = 18;
main.getRange("AK:AK").format.columnWidth = 30;
main.getRange("AL:AL").format.columnWidth = 18;
const capacityR1C1 = firstNumericR1C1(-19);
const rangeR1C1 = firstNumericR1C1(-20);
main.getRange("AH2").formulasR1C1 = [[`=IF(RC[-16]<>\"\",RC[-16],IFERROR(IF(AND(${capacityR1C1}*100/${rangeR1C1}>=3,${capacityR1C1}*100/${rangeR1C1}<=40),${capacityR1C1}*100/${rangeR1C1},\"\"),\"\"))`]];
main.getRange(`AH2:AH${mainValues.length}`).fillDown();
const capacityMassR1C1 = firstNumericR1C1(-20);
const densityR1C1 = firstNumericR1C1(-18);
main.getRange("AI2").formulasR1C1 = [[`=IFERROR(IF(AND(${capacityMassR1C1}>=75,${capacityMassR1C1}<=85,${densityR1C1}>0),${capacityMassR1C1}*1000/${densityR1C1},\"\"),\"\")`]];
main.getRange(`AI2:AI${mainValues.length}`).fillDown();
main.getRange("AJ2").formulasR1C1 = [["=IF(RC[-1]=\"\",\"\",\"非称重实测\")"]];
main.getRange(`AJ2:AJ${mainValues.length}`).fillDown();
main.getRange(`AH2:AI${mainValues.length}`).format.numberFormat = "0.0";

const auditFields = ["excel_row", "batch", "product_model", "field", "old_value", "new_value", "action", "evidence_tier", "method", "source_reference", "confidence", "status", "copyright"];
addDataSheet(
  workbook, "【参数补全审计】", "参数补全审计", "事实、公式推导和候选值分层记录；原始非空有效值不被覆盖。",
  auditFields,
  ["Excel行", "批次", "产品型号", "字段", "原值", "新值/候选", "动作", "证据等级", "方法", "来源行/来源ID", "置信度", "状态", "版权"],
  auditRows, "EnrichmentAuditTable",
  { product_model: 22, field: 25, old_value: 20, new_value: 20, method: 34, source_reference: 32, copyright: 28 },
);

const torqueQueue = records.torque_records.map((item) => {
  const external = externalMotorByRow.get(item.excel_row);
  const queueItem = external?.enrichedTarget || item;
  return {
    priority: external ? (external.complete ? "P0-外部候选复核" : "P0-外部局部候选复核") : Number(item.batch) >= 405 ? "P1-最新批次" : item.preferred_candidate_nm ? "P2-内部候选复核" : item.motor_model ? "P3-电机型号检索" : "P4-车型OCR",
    excel_row: item.excel_row,
    batch: item.batch,
    product_model: item.product_model,
    brand: item.brand,
    enterprise: item.enterprise,
    motor_supplier: item.motor_supplier,
    motor_model: queueItem.motor_model,
    motor_peak_power_kw: queueItem.motor_peak_power_kw,
    motor_total_power_kw: queueItem.motor_total_power_kw,
    current_value: item.current_value,
    source_state: item.state,
    internal_candidate_nm: item.preferred_candidate_nm,
    internal_method: item.preferred_method,
    external_value_nm: external?.value || "",
    external_source_level: external ? [...new Set(external.matched.map((evidence) => evidence.evidence_tier))].join("/") : "",
    source_title: external ? external.matched.map((evidence) => evidence.source_title).join("；") : "",
    source_url: external ? external.matched.map((evidence) => evidence.source_url).join("；") : "",
    published_date: external ? external.matched.map((evidence) => evidence.published_date).filter(Boolean).join("；") : "",
    evidence_location: external ? external.matched.map((evidence) => evidence.evidence_location).join("；") : "",
    review_status: external ? (external.complete ? "evidence_collected_pending_final" : "partial_component_evidence_pending") : "pending",
    copyright: COPYRIGHT,
  };
}).sort((a, b) => a.priority.localeCompare(b.priority) || Number(b.batch) - Number(a.batch) || a.excel_row - b.excel_row);
const torqueFields = Object.keys(torqueQueue[0]);
addDataSheet(
  workbook, "【电机扭矩OCR队列】", "电机扭矩 OCR / 外部证据队列", "仅有“功率/扭矩”明确复合值时直接解析；纯数值、kW 值和空白均进入证据队列。OCR 写入前必须二次复核。",
  torqueFields,
  ["优先级", "Excel行", "批次", "产品型号", "品牌", "企业", "电机企业", "电机型号", "峰值功率", "总功率", "原字段", "原字段状态", "内部候选扭矩", "内部方法", "外部扭矩", "来源等级", "来源标题", "来源URL", "发布日期", "证据位置", "复核状态", "版权"],
  torqueQueue, "MotorTorqueQueueTable",
  { product_model: 22, enterprise: 28, motor_supplier: 30, motor_model: 24, current_value: 20, internal_method: 32, source_title: 30, source_url: 44, evidence_location: 24, copyright: 28 },
);

const batteryFields = [
  "excel_row", "batch", "product_model", "brand", "enterprise", "vehicle_name", "generic_name", "capacity_kwh", "chemistry", "system_density_wh_kg",
  "derived_pack_mass_kg", "derived_mass_status", "cell_supplier", "pack_supplier", "exact_system_voltage_v", "pack_length_mm", "pack_width_mm", "pack_height_mm",
  "battery_evidence_tier", "battery_source_title", "battery_source_url", "battery_announcement_date", "battery_published_date", "battery_review_status",
  "direct_dimension_evidence_required", "copyright",
];
addDataSheet(
  workbook, "【80kWh电池专项】", "80±5 kWh 电池包参数专项", "质量仅由同一行容量与系统能量密度推导并标为“非称重实测”；供应商、精确电压和包体 L×W×H 需直接证据。",
  batteryFields,
  ["Excel行", "批次", "产品型号", "品牌", "企业", "车型名称", "通用名称", "容量(kWh)", "化学体系", "系统能量密度(Wh/kg)", "估算包质量(kg)", "质量状态", "电芯生产企业", "电池总成生产企业", "精确系统电压(V)", "包长(mm)", "包宽(mm)", "包高(mm)", "证据等级", "来源标题", "来源URL", "公告生效日期", "页面发布日期", "复核状态", "尺寸直接证据要求", "版权"],
  batterySpecialRecords, "Battery80kWhTable",
  { product_model: 22, enterprise: 28, vehicle_name: 24, generic_name: 24, chemistry: 22, cell_supplier: 36, pack_supplier: 30, battery_source_title: 38, battery_source_url: 48, battery_review_status: 28, direct_dimension_evidence_required: 24, copyright: 28 },
);

const externalEvidenceLedger = [
  ...(externalEvidence.motor_evidence || []).map((item) => ({
    evidence_type: "motor_torque", key: item.motor_model, parameter: "峰值扭矩", value: item.value_nm, unit: "N·m",
    expected_power_kw: (item.expected_peak_kw || []).join("|"), applicable_scope: [
      (item.applicable_brands || []).length ? `品牌=${item.applicable_brands.join("|")}` : "",
      (item.applicable_product_models || []).length ? `产品型号=${item.applicable_product_models.join("|")}` : "",
    ].filter(Boolean).join("；") || "电机型号＋功率", evidence_tier: item.evidence_tier, source_title: item.source_title,
    source_url: item.source_url, announcement_date: "", published_date: item.published_date, retrieved_date: externalEvidence.retrieved_date,
    evidence_location: item.evidence_location, method: item.method, confidence: item.confidence, status: item.status, notes: item.notes, copyright: COPYRIGHT,
  })),
  ...(externalEvidence.battery_evidence || []).map((item) => ({
    evidence_type: "battery_supplier", key: item.product_model, parameter: "电芯/电池总成生产企业",
    value: [item.cell_supplier, item.pack_supplier].filter(Boolean).join(" / "), unit: "", expected_power_kw: "", applicable_scope: `产品型号=${item.product_model}`, evidence_tier: item.evidence_tier,
    source_title: item.source_title, source_url: item.source_url, announcement_date: item.announcement_effective_date,
    published_date: item.published_date, retrieved_date: externalEvidence.retrieved_date, evidence_location: item.evidence_location,
    method: item.method, confidence: item.confidence, status: item.status, notes: item.notes, copyright: COPYRIGHT,
  })),
];
const externalEvidenceFields = ["evidence_type", "key", "parameter", "value", "unit", "expected_power_kw", "applicable_scope", "evidence_tier", "source_title", "source_url", "announcement_date", "published_date", "retrieved_date", "evidence_location", "method", "confidence", "status", "notes", "copyright"];
addDataSheet(
  workbook, "【外部证据台账】", "外部证据台账", "外部资料按证据等级登记；候选值不覆盖原始字段。发布日期、公告生效日期与检索日期分列。",
  externalEvidenceFields,
  ["证据类型", "匹配键", "参数", "值", "单位", "适用峰值功率(kW)", "适用范围", "证据等级", "来源标题", "来源URL", "公告生效日期", "页面发布日期", "检索日期", "证据位置", "匹配方法", "置信度", "状态", "说明", "版权"],
  externalEvidenceLedger, "ExternalEvidenceLedgerTable",
  { key: 24, parameter: 26, expected_power_kw: 22, applicable_scope: 34, source_title: 42, source_url: 52, evidence_location: 48, method: 36, notes: 48, copyright: 28 },
);

const motorGroupFields = ["motor_model", "target_rows", "batches", "brands", "motor_suppliers", "peak_power_kw_values", "total_power_kw_values", "sample_vehicle_models", "research_priority", "copyright"];
addDataSheet(
  workbook, "【扭矩研究分组】", "电机扭矩研究分组", "按电机型号聚合队列，优先处理覆盖行数最多的型号；“电机型号缺失”组必须按车型型号 OCR。",
  motorGroupFields,
  ["电机型号", "目标行数", "批次", "品牌", "电机企业", "峰值功率值", "总功率值", "样例车型型号", "研究路径", "版权"],
  records.motor_research_groups, "MotorResearchGroupsTable",
  { motor_model: 22, batches: 26, brands: 32, motor_suppliers: 34, peak_power_kw_values: 30, total_power_kw_values: 24, sample_vehicle_models: 46, research_priority: 24, copyright: 28 },
);

const batteryGroupFields = ["brand", "capacity_kwh", "chemistry", "system_density_wh_kg", "target_rows", "derived_pack_mass_kg", "derived_mass_status", "batches", "enterprises", "sample_vehicle_models", "direct_evidence_needed", "copyright"];
addDataSheet(
  workbook, "【80kWh研究分组】", "80±5 kWh 电池研究分组", "按品牌、容量、化学体系和系统能量密度聚合；外部参数仅在同配置直接证据充分时传播。",
  batteryGroupFields,
  ["品牌", "容量(kWh)", "化学体系", "系统能量密度", "覆盖行数", "估算包质量", "质量状态", "批次", "企业", "样例车型型号", "待查直接证据", "版权"],
  records.battery_research_groups, "BatteryResearchGroupsTable",
  { brand: 22, chemistry: 22, batches: 24, enterprises: 30, sample_vehicle_models: 46, direct_evidence_needed: 30, copyright: 28 },
);

const explanation = workbook.worksheets.add("【补全说明】");
explanation.showGridLines = false;
explanation.getRange("A1:F1").merge();
explanation.getRange("A1").values = [["TA-Analysis 参数证据补全 · 执行摘要"]];
explanation.getRange("A1:F1").format = { fill: COLORS.navy, font: { bold: true, color: COLORS.white, size: 18 } };
explanation.getRange("A2:F2").merge();
explanation.getRange("A2").values = [[COPYRIGHT]];
explanation.getRange("A2:F2").format = { fill: COLORS.paleBlue, font: { italic: true, color: COLORS.navy } };
explanation.getRange("A4:B14").values = [
  ["指标", "结果"],
  ["源表记录", 4459],
  ["污染单元格已识别/清理", records.contamination_records.length],
  ["同型号唯一一致补全", records.same_model_candidates.length],
  ["电耗公式候选", consumptionFormulaCandidates.length],
  ["电机扭矩待处理", records.torque_records.length],
  ["扭矩内部证据候选", records.torque_records.filter((row) => row.preferred_candidate_nm).length],
  ["扭矩外部证据候选（含局部）", externalMotorByRow.size],
  ["外部证据覆盖电机型号", externalEvidence.motor_evidence.length],
  ["电池供应商精确型号候选", batterySpecialRecords.filter((row) => row.battery_review_status === "exact_product_model_candidate").length],
  ["80±5 kWh记录 / 可推导质量", `${records.battery_80kwh_records.length} / ${records.battery_80kwh_records.filter((row) => row.derived_pack_mass_kg !== "").length}`],
];
styleHeader(explanation.getRange("A4:B4"), COLORS.teal);
explanation.getRange("D4:F4").merge();
explanation.getRange("D4").values = [["证据与推断边界"]];
styleHeader(explanation.getRange("D4:F4"), COLORS.teal);
explanation.getRange("D5:F11").merge(true);
explanation.getRange("D5:D11").values = [
  ["事实：原始公告/OEM/供应商直接资料。"],
  ["C级：同一车型或电机型号的唯一一致值，只作为候选或填充空值。"],
  ["外部证据：按电机型号＋功率匹配；多电机车型只记录已证明的单电机分量。"],
  ["电耗：同一行容量×100÷纯电续航，范围 3–40。"],
  ["80kWh质量：同一行容量×1000÷系统能量密度，统一标注“非称重实测”。"],
  ["电池包尺寸：必须有直接 L×W×H 证据，车辆尺寸不得替代。"],
  ["电压：营销“800V”不得当作精确系统电压。"],
];
explanation.getRange("D5:F11").format = { fill: COLORS.paleAmber, wrapText: true, verticalAlignment: "top", borders: { preset: "all", style: "thin", color: COLORS.border } };
explanation.getRange("A16:F16").merge();
explanation.getRange("A16").values = [["已核验的最新工信部来源登记（成文日期与发布日期分列）"]];
styleHeader(explanation.getRange("A16:F16"), COLORS.teal);
explanation.getRange("A17:F22").values = [
  ["批次", "公告号", "成文日期", "发布日期", "来源标题", "官方链接"],
  [405, "2026年第8号", "2026-04-13", "2026-04-14", "道路机动车辆生产企业及产品（第405批）等目录", "https://www.miit.gov.cn/jgsj/zbys/wjfb/art/2026/art_9fa19a953e2b485e8c8e8da67b207d65.html"],
  [406, "2026年第10号", "", "2026-05-09", "道路机动车辆生产企业及产品（第406批）等目录", "https://www.miit.gov.cn/jgsj/zbys/wjfb/art/2026/art_421a049da25f4b088bdb18a20fa8a690.html"],
  [407, "2026年第14号", "2026-06-11", "2026-06-12", "道路机动车辆生产企业及产品（第407批）等目录", "https://www.miit-eidc.org.cn/art/2026/6/12/art_1691_12457.html"],
  [408, "2026年第19号", "2026-07-16", "2026-07-17", "道路机动车辆生产企业及产品（第408批）等目录", "https://www.miit.gov.cn/zwgk/zcwj/wjfb/gg/art/2026/art_5d1f38038f8847509a6f3e8b700bef3d.html"],
  [409, "2026年第21号", "2026-08-12", "2026-08-13", "道路机动车辆生产企业及产品（第409批）等目录", "https://www.miit.gov.cn/zwgk/zcwj/wjfb/gg/art/2026/art_00d965bfa9ea4bc89bc7613cf911c896.html"],
];
styleHeader(explanation.getRange("A17:F17"), COLORS.navy);
explanation.getRange("A18:F22").format = { wrapText: true, verticalAlignment: "top", borders: { preset: "all", style: "thin", color: COLORS.border } };
explanation.getRange("A24:F24").merge();
explanation.getRange("A24").values = [[`源文件 SHA-256: ${sourceSha256} · 版本: ${VERSION}`]];
explanation.getRange("A24:F24").format = { fill: COLORS.gray, font: { color: COLORS.text, italic: true } };
explanation.getRange("A:A").format.columnWidth = 24;
explanation.getRange("B:B").format.columnWidth = 24;
explanation.getRange("C:D").format.columnWidth = 16;
explanation.getRange("E:E").format.columnWidth = 42;
explanation.getRange("F:F").format.columnWidth = 62;
explanation.freezePanes.freezeRows(2);

const copyrightSheet = workbook.worksheets.getItem("版权声明");
copyrightSheet.getRange("A1:B6").values = [
  [COPYRIGHT, null],
  ["文件", "NEV公告参数汇总表_证据补全候选（341~409批）.xlsx"],
  ["内容", "原始4459条数据 + 数据污染清理审计 + 同型号候选 + 电耗公式 + 外部扭矩证据 + 电机扭矩OCR队列 + 80±5kWh电池专项"],
  ["证据边界", "事实、同型号推断、公式推导和外部待核实字段严格分列；不将推断伪装成事实"],
  ["授权", "MIT License，详见仓库 LICENSE；引用数据请注明来源"],
  ["生成", `TA-Analysis ${VERSION} · 2026-08-23`],
];
copyrightSheet.getRange("A1:B1").format = { fill: COLORS.navy, font: { bold: true, color: COLORS.white } };
copyrightSheet.getRange("A1:B6").format.wrapText = true;
copyrightSheet.getRange("A:A").format.columnWidth = 18;
copyrightSheet.getRange("B:B").format.columnWidth = 84;

await fs.mkdir(OUTPUT_DIR, { recursive: true });
const inspection = await workbook.inspect({ kind: "workbook,sheet,table", maxChars: 12000, tableMaxRows: 3, tableMaxCols: 5, tableMaxCellChars: 80 });
await fs.writeFile(`${OUTPUT_DIR}/workbook_inspection.ndjson`, `${inspection.ndjson}\n`, "utf8");
try {
  const preview = await workbook.render({ sheetName: "【补全说明】", autoCrop: "all", scale: 1, format: "png" });
  await fs.writeFile(PREVIEW_FILE, new Uint8Array(await preview.arrayBuffer()));
} catch (previewError) {
  console.warn(`[warn] preview skipped: ${previewError.message}`);
}
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(OUTPUT_FILE);
console.log(JSON.stringify({ output: OUTPUT_FILE, preview: PREVIEW_FILE, source_sha256: sourceSha256, audit_rows: auditRows.length, copyright: COPYRIGHT }, null, 2));
