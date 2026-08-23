/**
 * Verify the TA-Analysis enrichment workbook.
 * Copyright © 2026 David YE
 * SPDX-License-Identifier: MIT
 */
import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "/opt/codex/runtimes/codex-primary-runtime/dependencies/node/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs";

const COPYRIGHT = "Copyright © 2026 David YE";
const file = process.argv[2];
const evidenceFile = process.argv[3];
if (!file) throw new Error("Workbook path required");
const invalidText = /^(?:请提供|抱歉|无法|未提供|未找到|未提取|无有效信息|未直接提供|没有提供|正文中未|原文中未)/;

function modelTokens(value) {
  const matches = String(value ?? "").toUpperCase().match(/[A-Z0-9]{5,}/g) || [];
  return [...new Set(matches.filter((token) => /[A-Z]/.test(token) && /\d/.test(token)))];
}

function numericTokens(...values) {
  return values.flatMap((value) => [...String(value ?? "").replace(/,/g, " ").matchAll(/-?\d+(?:\.\d+)?/g)].map((match) => Number(match[0]))).filter(Number.isFinite);
}

function powerCompatible(row, evidenceItem) {
  const applicableBrands = (evidenceItem.applicable_brands || []).map((value) => String(value).trim());
  if (applicableBrands.length && !applicableBrands.includes(String(row.brand ?? "").trim())) return false;
  const applicableModels = (evidenceItem.applicable_product_models || []).map((value) => String(value).trim().toUpperCase());
  if (applicableModels.length && !applicableModels.includes(String(row.productModel ?? "").trim().toUpperCase())) return false;
  const expected = (evidenceItem.expected_peak_kw || []).map(Number).filter(Number.isFinite);
  if (!expected.length) return true;
  const observed = numericTokens(row.peak, row.total, row.current);
  if (!observed.length) return true;
  return expected.some((expectedValue) => observed.some((observedValue) => Math.abs(expectedValue - observedValue) < 0.01));
}
const blob = await FileBlob.load(file);
const workbook = await SpreadsheetFile.importXlsx(blob);
const expectedSheets = [
  "NEV公告参数汇总", "版权声明", "【参数补全审计】", "【电机扭矩OCR队列】", "【80kWh电池专项】",
  "【外部证据台账】", "【扭矩研究分组】", "【80kWh研究分组】", "【补全说明】",
];
for (const name of expectedSheets) workbook.worksheets.getItem(name);
const main = workbook.worksheets.getItem("NEV公告参数汇总");
const values = main.getUsedRange(true).values;
const headers = values[0].map((value) => String(value ?? ""));
const column = Object.fromEntries(headers.map((header, index) => [header, index]));
const requiredHeaders = [
  "电机扭矩解析值(N·m)", "电机扭矩补全候选(N·m)", "电机扭矩证据状态", "审计后百公里电耗(kWh/100km)",
  "电池组估算质量(kg，仅80±5kWh)", "电池组质量状态", "数据质量/补全状态", "补全版本",
];
for (const header of requiredHeaders) if (!(header in column)) throw new Error(`Missing header: ${header}`);
const data = values.slice(1);
const nonBlank = (header) => data.filter((row) => row[column[header]] !== null && row[column[header]] !== undefined && String(row[column[header]]).trim() !== "").length;
const errorValues = [];
for (let rowIndex = 0; rowIndex < data.length; rowIndex += 1) {
  for (const header of ["审计后百公里电耗(kWh/100km)", "电池组估算质量(kg，仅80±5kWh)", "电池组质量状态"]) {
    const value = data[rowIndex][column[header]];
    if (typeof value === "string" && /^#(?:REF!|DIV\/0!|VALUE!|NAME\?|N\/A|NUM!|NULL!)/.test(value)) errorValues.push({ excel_row: rowIndex + 2, header, value });
  }
}
const contamination = [];
for (let rowIndex = 0; rowIndex < data.length; rowIndex += 1) {
  for (let columnIndex = 0; columnIndex < 30; columnIndex += 1) {
    const value = data[rowIndex][columnIndex];
    if (typeof value === "string" && invalidText.test(value.trim())) contamination.push({ excel_row: rowIndex + 2, field: headers[columnIndex], value });
  }
}
const audit = workbook.worksheets.getItem("【参数补全审计】").getUsedRange(true).values;
const torque = workbook.worksheets.getItem("【电机扭矩OCR队列】").getUsedRange(true).values;
const battery = workbook.worksheets.getItem("【80kWh电池专项】").getUsedRange(true).values;
const evidence = workbook.worksheets.getItem("【外部证据台账】").getUsedRange(true).values;
const externalEvidence = evidenceFile ? JSON.parse(await fs.readFile(evidenceFile, "utf8")) : null;
if (externalEvidence && externalEvidence.copyright !== COPYRIGHT) throw new Error("External evidence copyright mismatch");
const motorEvidenceByModel = new Map((externalEvidence?.motor_evidence || []).map((item) => [item.motor_model.toUpperCase(), item]));
const torqueHeaders = torque[4].map((value) => String(value ?? ""));
const torqueColumn = Object.fromEntries(torqueHeaders.map((header, index) => [header, index]));
const externalTorqueRows = torque.slice(5).filter((row) => String(row[torqueColumn["外部扭矩"]] ?? "").trim() !== "");
const externalEvidenceMismatches = [];
const externalMotorCounts = {};
if (externalEvidence) {
  for (const row of externalTorqueRows) {
    const models = modelTokens(row[torqueColumn["电机型号"]]);
    const rowForPower = {
      peak: row[torqueColumn["峰值功率"]], total: row[torqueColumn["总功率"]], current: row[torqueColumn["原字段"]],
      brand: row[torqueColumn["品牌"]], productModel: row[torqueColumn["产品型号"]],
    };
    const matched = models.map((model) => motorEvidenceByModel.get(model)).filter((item) => item && powerCompatible(rowForPower, item));
    const matchedModels = new Set(matched.map((item) => item.motor_model.toUpperCase()));
    const complete = models.length === 1 || (models.length > 1 && models.every((model) => matchedModels.has(model)));
    const expectedValue = complete && matched.length === 1 ? String(matched[0].value_nm) : matched.map((item) => `${item.motor_model}:${item.value_nm}`).join("；");
    const actualValue = String(row[torqueColumn["外部扭矩"]] ?? "");
    if (!matched.length || actualValue !== expectedValue) externalEvidenceMismatches.push({ excel_row: row[torqueColumn["Excel行"]], actualValue, expectedValue, models });
    for (const item of matched) externalMotorCounts[item.motor_model] = (externalMotorCounts[item.motor_model] || 0) + 1;
  }
}
const torqueStatusIndex = column["电机扭矩证据状态"];
const torqueExternalCandidates = data.filter((row) => /外部(?:局部)?候选$/.test(String(row[torqueStatusIndex] ?? ""))).length;
const torqueExternalPartialCandidates = data.filter((row) => String(row[torqueStatusIndex] ?? "").includes("外部局部候选")).length;
const torqueInternalCandidates = data.filter((row) => String(row[torqueStatusIndex] ?? "").startsWith("C级候选：")).length;
const batteryHeaders = battery[4].map((value) => String(value ?? ""));
const batteryReviewIndex = batteryHeaders.indexOf("复核状态");
const batterySupplierCandidates = battery.slice(5).filter((row) => row[batteryReviewIndex] === "exact_product_model_candidate").length;
const copyrightChecks = {
  copyright_sheet: workbook.worksheets.getItem("版权声明").getRange("A1").values[0][0],
  audit_sheet: workbook.worksheets.getItem("【参数补全审计】").getRange("A2").values[0][0],
  torque_sheet: workbook.worksheets.getItem("【电机扭矩OCR队列】").getRange("A2").values[0][0],
  battery_sheet: workbook.worksheets.getItem("【80kWh电池专项】").getRange("A2").values[0][0],
  evidence_sheet: workbook.worksheets.getItem("【外部证据台账】").getRange("A2").values[0][0],
  summary_sheet: workbook.worksheets.getItem("【补全说明】").getRange("A2").values[0][0],
};
const result = {
  file,
  bytes: (await fs.stat(file)).size,
  sheets: expectedSheets.length,
  records: data.length,
  columns: headers.length,
  torque_parsed: nonBlank("电机扭矩解析值(N·m)"),
  torque_candidate_total: nonBlank("电机扭矩补全候选(N·m)"),
  torque_internal_candidates: torqueInternalCandidates,
  torque_external_candidates: torqueExternalCandidates,
  torque_external_partial_candidates: torqueExternalPartialCandidates,
  audited_consumption_nonblank: nonBlank("审计后百公里电耗(kWh/100km)"),
  battery_mass_derived: nonBlank("电池组估算质量(kg，仅80±5kWh)"),
  audit_rows: audit.length - 5,
  torque_queue_rows: torque.length - 5,
  battery_80kwh_rows: battery.length - 5,
  battery_supplier_candidates: batterySupplierCandidates,
  external_evidence_rows: evidence.length - 5,
  external_motor_counts: externalMotorCounts,
  external_evidence_mismatches: externalEvidenceMismatches.length,
  formula_errors: errorValues.length,
  contamination_remaining: contamination.length,
  copyright_checks: copyrightChecks,
  copyright_ok: Object.values(copyrightChecks).every((value) => value === COPYRIGHT),
};
const failures = [];
if (result.records !== 4459) failures.push(`records=${result.records}`);
if (result.columns !== 38) failures.push(`columns=${result.columns}`);
if (result.formula_errors !== 0) failures.push(`formula_errors=${result.formula_errors}`);
if (result.contamination_remaining !== 0) failures.push(`contamination=${result.contamination_remaining}`);
if (!result.copyright_ok) failures.push("copyright");
if (result.torque_queue_rows !== 3180) failures.push(`torque_queue=${result.torque_queue_rows}`);
if (result.battery_80kwh_rows !== 453) failures.push(`battery_rows=${result.battery_80kwh_rows}`);
if (result.torque_parsed !== 1279) failures.push(`torque_parsed=${result.torque_parsed}`);
if (result.torque_candidate_total !== 364) failures.push(`torque_candidate_total=${result.torque_candidate_total}`);
if (result.torque_internal_candidates !== 38) failures.push(`torque_internal_candidates=${result.torque_internal_candidates}`);
if (result.torque_external_candidates !== 326) failures.push(`torque_external_candidates=${result.torque_external_candidates}`);
if (result.torque_external_partial_candidates !== 32) failures.push(`torque_external_partial_candidates=${result.torque_external_partial_candidates}`);
if (result.battery_supplier_candidates !== 5) failures.push(`battery_supplier_candidates=${result.battery_supplier_candidates}`);
if (result.external_evidence_rows !== 36) failures.push(`external_evidence_rows=${result.external_evidence_rows}`);
if (externalEvidence && result.external_evidence_mismatches !== 0) failures.push(`external_evidence_mismatches=${result.external_evidence_mismatches}`);
if (result.audited_consumption_nonblank !== 4265) failures.push(`audited_consumption=${result.audited_consumption_nonblank}`);
if (result.battery_mass_derived !== 424) failures.push(`battery_mass=${result.battery_mass_derived}`);
if (result.audit_rows !== 2820) failures.push(`audit_rows=${result.audit_rows}`);
result.status = failures.length ? "FAIL" : "PASS";
result.failures = failures;
console.log(JSON.stringify(result, null, 2));
if (failures.length) process.exitCode = 2;
