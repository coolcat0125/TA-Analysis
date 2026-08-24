/**
 * Build an evidence-safe inventory for TA-Analysis.
 * Copyright © 2026 David YE
 * SPDX-License-Identifier: MIT
 */
import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile } from "/opt/codex/runtimes/codex-primary-runtime/dependencies/node/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs";

const COPYRIGHT = "Copyright © 2026 David YE";
const SOURCE = process.argv[2] || "source_341_409.xlsx";
const OUTPUT_DIR = process.argv[3] || "analysis-output";
const MISSING = new Set(["", "/", "-", "--", "N/A", "NA", "null", "None"]);
const INVALID_TEXT = /^(?:请提供|抱歉|无法|未提供|未找到|未提取|无有效信息|未直接提供|没有提供|正文中未|原文中未)/;

function isMissing(value) {
  if (value === null || value === undefined) return true;
  const text = String(value).trim();
  return MISSING.has(text) || INVALID_TEXT.test(text);
}

function isContaminated(value) {
  return typeof value === "string" && INVALID_TEXT.test(value.trim());
}

function scalarKey(value) {
  if (typeof value === "number") return `n:${Number(value.toPrecision(12))}`;
  return `s:${String(value).trim().replace(/\s+/g, " ")}`;
}

function toNumber(value) {
  if (isMissing(value)) return null;
  if (typeof value === "number" && Number.isFinite(value)) return value;
  const match = String(value).replace(/,/g, "").match(/-?\d+(?:\.\d+)?/);
  return match ? Number(match[0]) : null;
}

function parseTorque(value) {
  if (isMissing(value)) return { state: "blank", torqueNm: [], normalized: "" };
  const text = String(value).trim().replace(/[：]/g, ":").replace(/[；，]/g, " ");
  const values = [];
  for (const match of text.matchAll(/\/\s*(-?\d+(?:\.\d+)?)/g)) values.push(Number(match[1]));
  for (const match of text.matchAll(/(-?\d+(?:\.\d+)?)\s*N\s*[·.\-]?\s*m\b/gi)) values.push(Number(match[1]));
  if (!text.includes("/")) {
    for (const match of text.matchAll(/(?:^|\s)(?:F|R|前|后)\s*:\s*(-?\d+(?:\.\d+)?)/gi)) values.push(Number(match[1]));
  }
  const unique = [...new Set(values.filter((v) => Number.isFinite(v) && v > 0 && v < 3000))];
  if (unique.length) return { state: "parseable", torqueNm: unique, normalized: text };
  if (text.includes("/")) return { state: "power_only", torqueNm: [], normalized: text };
  return { state: "ambiguous", torqueNm: [], normalized: text };
}

function splitMulti(value) {
  if (isMissing(value)) return [];
  return String(value).split(/[\/；;\n]+/).map((part) => part.trim()).filter(Boolean);
}

function csvEscape(value) {
  const text = value === null || value === undefined ? "" : String(value);
  return /[",\n\r]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
}

async function writeCsv(fileName, records, fields) {
  const lines = [fields.map(csvEscape).join(",")];
  for (const record of records) lines.push(fields.map((field) => csvEscape(record[field])).join(","));
  await fs.writeFile(path.join(OUTPUT_DIR, fileName), `\uFEFF${lines.join("\n")}\n`, "utf8");
}

const blob = await FileBlob.load(SOURCE);
const workbook = await SpreadsheetFile.importXlsx(blob);
const sheet = workbook.worksheets.getItem("NEV公告参数汇总");
const values = sheet.getUsedRange(true).values;
const headers = values[0].map((value) => String(value ?? "").trim());
const rows = values.slice(1).filter((row) => row.some((value) => !isMissing(value)));
const col = Object.fromEntries(headers.map((header, index) => [header, index]));
const required = ["批次", "产品型号", "电池容量(kWh)", "电池能量密度(Wh/kg)", "纯电续航里程(km)", "百公里电耗(kWh/100km)", "电机功率/扭矩"];
for (const header of required) if (!(header in col)) throw new Error(`Missing required column: ${header}`);

const coverage = headers.map((header, index) => {
  const filled = rows.filter((row) => !isMissing(row[index])).length;
  return { column_index: index + 1, header, filled, missing: rows.length - filled, coverage: Number((filled / rows.length).toFixed(6)), copyright: COPYRIGHT };
});

const contaminationRecords = [];
for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
  for (let columnIndex = 0; columnIndex < headers.length; columnIndex += 1) {
    if (isContaminated(rows[rowIndex][columnIndex])) contaminationRecords.push({
      excel_row: rowIndex + 2,
      batch: rows[rowIndex][col["批次"]],
      product_model: rows[rowIndex][col["产品型号"]],
      field: headers[columnIndex],
      contaminated_value: rows[rowIndex][columnIndex],
      recommended_action: "clear_and_research",
      copyright: COPYRIGHT,
    });
  }
}

const modelGroups = new Map();
for (let index = 0; index < rows.length; index += 1) {
  const model = String(rows[index][col["产品型号"]] ?? "").trim();
  if (!model) continue;
  if (!modelGroups.has(model)) modelGroups.set(model, []);
  modelGroups.get(model).push({ row: rows[index], excelRow: index + 2 });
}

const torqueRecords = [];
const torqueStateCounts = { blank: 0, power_only: 0, ambiguous: 0, parseable: 0 };
const torqueDonorsByModel = new Map();
const torqueDonorsByMotorModel = new Map();
for (let index = 0; index < rows.length; index += 1) {
  const row = rows[index];
  const parsed = parseTorque(row[col["电机功率/扭矩"]]);
  torqueStateCounts[parsed.state] += 1;
  const model = String(row[col["产品型号"]] ?? "").trim();
  if (parsed.state === "parseable" && model) {
    if (!torqueDonorsByModel.has(model)) torqueDonorsByModel.set(model, []);
    torqueDonorsByModel.get(model).push({ value: row[col["电机功率/扭矩"]], excelRow: index + 2, batch: row[col["批次"]], torqueNm: parsed.torqueNm });
  }
  if (parsed.state === "parseable") {
    const motorModels = splitMulti(row[col["电机型号"]]);
    if (motorModels.length === parsed.torqueNm.length) {
      for (let motorIndex = 0; motorIndex < motorModels.length; motorIndex += 1) {
        const motorModel = motorModels[motorIndex];
        if (!torqueDonorsByMotorModel.has(motorModel)) torqueDonorsByMotorModel.set(motorModel, []);
        torqueDonorsByMotorModel.get(motorModel).push({
          torqueNm: parsed.torqueNm[motorIndex],
          excelRow: index + 2,
          batch: row[col["批次"]],
          vehicleModel: model,
        });
      }
    }
  }
  if (parsed.state !== "parseable") torqueRecords.push({
    excel_row: index + 2,
    batch: row[col["批次"]],
    product_model: model,
    brand: row[col["产品商标"]],
    enterprise: row[col["企业名称"]],
    current_value: row[col["电机功率/扭矩"]],
    state: parsed.state,
    motor_peak_power_kw: row[col["电机峰值功率(kW)"]],
    motor_total_power_kw: row[col["电机总功率(kW)"]],
    motor_supplier: row[col["电机生产企业"]],
    motor_model: row[col["电机型号"]],
    same_model_candidate: "",
    same_model_source_rows: "",
    same_model_source_batches: "",
    inference_tier: "",
    confidence: "",
    motor_model_candidate_nm: "",
    motor_model_source_rows: "",
    motor_model_source_batches: "",
    preferred_candidate_nm: "",
    preferred_method: "",
    candidate_conflict: "",
    copyright: COPYRIGHT,
  });
}

let torqueSameModelCandidates = 0;
let torqueMotorModelCandidates = 0;
let torqueInternalResolved = 0;
for (const record of torqueRecords) {
  const donors = torqueDonorsByModel.get(record.product_model) || [];
  const unique = new Map();
  for (const donor of donors) unique.set(scalarKey(donor.value), donor.value);
  if (unique.size === 1) {
    record.same_model_candidate = [...unique.values()][0];
    record.same_model_source_rows = donors.map((d) => d.excelRow).join("|");
    record.same_model_source_batches = [...new Set(donors.map((d) => d.batch))].join("|");
    record.inference_tier = "C-same-model";
    record.confidence = donors.length >= 2 ? "high" : "medium";
    torqueSameModelCandidates += 1;
  }
  const motorModels = splitMulti(record.motor_model);
  const motorCandidates = [];
  const motorSourceRows = [];
  const motorSourceBatches = [];
  let motorModelsResolvable = motorModels.length > 0;
  for (const motorModel of motorModels) {
    const motorDonors = torqueDonorsByMotorModel.get(motorModel) || [];
    const uniqueTorque = [...new Set(motorDonors.map((donor) => donor.torqueNm))];
    if (uniqueTorque.length !== 1) {
      motorModelsResolvable = false;
      break;
    }
    motorCandidates.push(uniqueTorque[0]);
    motorSourceRows.push(...motorDonors.map((donor) => donor.excelRow));
    motorSourceBatches.push(...motorDonors.map((donor) => donor.batch));
  }
  if (motorModelsResolvable) {
    record.motor_model_candidate_nm = motorCandidates.join("|");
    record.motor_model_source_rows = [...new Set(motorSourceRows)].join("|");
    record.motor_model_source_batches = [...new Set(motorSourceBatches)].join("|");
    torqueMotorModelCandidates += 1;
  }
  const sameModelParsed = parseTorque(record.same_model_candidate).torqueNm;
  const sameModelValue = sameModelParsed.length ? sameModelParsed.join("|") : "";
  const motorModelValue = record.motor_model_candidate_nm;
  if (sameModelValue && motorModelValue && sameModelValue !== motorModelValue) {
    record.candidate_conflict = `${sameModelValue} != ${motorModelValue}`;
  } else if (motorModelValue) {
    record.preferred_candidate_nm = motorModelValue;
    record.preferred_method = "同一电机型号官方公告记录唯一一致";
    torqueInternalResolved += 1;
  } else if (sameModelValue) {
    record.preferred_candidate_nm = sameModelValue;
    record.preferred_method = "同一产品型号跨批次记录唯一一致";
    torqueInternalResolved += 1;
  }
}

const excludedConsensus = new Set(["批次", "是否减免购置税", "数据来源", "电机功率/扭矩"]);
const sameModelCandidates = [];
for (const [model, members] of modelGroups) {
  if (members.length < 2) continue;
  for (let columnIndex = 0; columnIndex < headers.length; columnIndex += 1) {
    const field = headers[columnIndex];
    if (excludedConsensus.has(field) || field === "产品型号") continue;
    const donors = members.filter((member) => !isMissing(member.row[columnIndex]));
    const unique = new Map(donors.map((member) => [scalarKey(member.row[columnIndex]), member.row[columnIndex]]));
    if (unique.size !== 1) continue;
    const candidateValue = [...unique.values()][0];
    for (const member of members.filter((item) => isMissing(item.row[columnIndex]))) {
      sameModelCandidates.push({
        excel_row: member.excelRow,
        batch: member.row[col["批次"]],
        product_model: model,
        field,
        current_value: member.row[columnIndex],
        candidate_value: candidateValue,
        source_rows: donors.map((donor) => donor.excelRow).join("|"),
        source_batches: [...new Set(donors.map((donor) => donor.row[col["批次"]]))].join("|"),
        evidence_tier: "C-same-model",
        method: "同一产品型号跨批次非空值唯一一致",
        confidence: donors.length >= 2 ? "high" : "medium",
        copyright: COPYRIGHT,
      });
    }
  }
}

const consumptionCandidates = [];
const packs80 = [];
for (let index = 0; index < rows.length; index += 1) {
  const row = rows[index];
  const capacity = toNumber(row[col["电池容量(kWh)"]]);
  const rangeKm = toNumber(row[col["纯电续航里程(km)"]]);
  const density = toNumber(row[col["电池能量密度(Wh/kg)"]]);
  const consumption = row[col["百公里电耗(kWh/100km)"]];
  if (isMissing(consumption) && capacity && rangeKm) {
    const candidate = capacity * 100 / rangeKm;
    if (candidate >= 3 && candidate <= 40) consumptionCandidates.push({
      excel_row: index + 2,
      batch: row[col["批次"]],
      product_model: row[col["产品型号"]],
      capacity_kwh: capacity,
      range_km: rangeKm,
      candidate_kwh_100km: Number(candidate.toFixed(3)),
      evidence_tier: "C-formula",
      method: "同一行电池容量×100÷纯电续航",
      confidence: "medium",
      copyright: COPYRIGHT,
    });
  }
  if (capacity !== null && capacity >= 75 && capacity <= 85) {
    const estimatedMass = density && density > 0 ? capacity * 1000 / density : null;
    packs80.push({
      excel_row: index + 2,
      batch: row[col["批次"]],
      product_model: row[col["产品型号"]],
      brand: row[col["产品商标"]],
      enterprise: row[col["企业名称"]],
      vehicle_name: row[col["车型名称"]],
      generic_name: row[col["通用名称"]],
      capacity_kwh: capacity,
      chemistry: row[col["电池类型"]],
      system_density_wh_kg: density,
      derived_pack_mass_kg: estimatedMass === null ? "" : Number(estimatedMass.toFixed(1)),
      derived_mass_status: estimatedMass === null ? "" : "非称重实测",
      pack_supplier: "",
      exact_system_voltage_v: "",
      pack_length_mm: "",
      pack_width_mm: "",
      pack_height_mm: "",
      direct_dimension_evidence_required: "yes",
      copyright: COPYRIGHT,
    });
  }
}

const unresolvedMotorGroupsMap = new Map();
for (const record of torqueRecords.filter((item) => !item.preferred_candidate_nm)) {
  const models = splitMulti(record.motor_model);
  for (const motorModel of models.length ? models : ["(电机型号缺失)"]) {
    if (!unresolvedMotorGroupsMap.has(motorModel)) unresolvedMotorGroupsMap.set(motorModel, []);
    unresolvedMotorGroupsMap.get(motorModel).push(record);
  }
}
const unresolvedMotorGroups = [...unresolvedMotorGroupsMap].map(([motorModel, members]) => ({
  motor_model: motorModel,
  target_rows: members.length,
  batches: [...new Set(members.map((item) => item.batch))].join("|"),
  brands: [...new Set(members.map((item) => item.brand).filter(Boolean))].join("|"),
  motor_suppliers: [...new Set(members.map((item) => item.motor_supplier).filter(Boolean))].join("|"),
  peak_power_kw_values: [...new Set(members.map((item) => item.motor_peak_power_kw).filter((value) => !isMissing(value)))].join("|"),
  total_power_kw_values: [...new Set(members.map((item) => item.motor_total_power_kw).filter((value) => !isMissing(value)))].join("|"),
  sample_vehicle_models: [...new Set(members.map((item) => item.product_model).filter(Boolean))].slice(0, 12).join("|"),
  research_priority: motorModel === "(电机型号缺失)" ? "vehicle-model-OCR" : "motor-model-source-search",
  copyright: COPYRIGHT,
})).sort((a, b) => b.target_rows - a.target_rows || a.motor_model.localeCompare(b.motor_model));

const batteryGroupsMap = new Map();
for (const record of packs80) {
  const key = [record.brand, record.capacity_kwh, record.chemistry, record.system_density_wh_kg].map((value) => String(value ?? "").trim()).join("||");
  if (!batteryGroupsMap.has(key)) batteryGroupsMap.set(key, []);
  batteryGroupsMap.get(key).push(record);
}
const batteryGroups = [...batteryGroupsMap.values()].map((members) => ({
  brand: members[0].brand,
  capacity_kwh: members[0].capacity_kwh,
  chemistry: members[0].chemistry,
  system_density_wh_kg: members[0].system_density_wh_kg,
  target_rows: members.length,
  derived_pack_mass_kg: members[0].derived_pack_mass_kg,
  derived_mass_status: members[0].derived_mass_status,
  batches: [...new Set(members.map((item) => item.batch))].join("|"),
  enterprises: [...new Set(members.map((item) => item.enterprise).filter(Boolean))].join("|"),
  sample_vehicle_models: [...new Set(members.map((item) => item.product_model).filter(Boolean))].slice(0, 12).join("|"),
  direct_evidence_needed: "supplier|exact_voltage|pack_LxWxH",
  copyright: COPYRIGHT,
})).sort((a, b) => b.target_rows - a.target_rows || String(a.brand).localeCompare(String(b.brand)));

await fs.mkdir(OUTPUT_DIR, { recursive: true });
await writeCsv("column_coverage.csv", coverage, Object.keys(coverage[0]));
await writeCsv("motor_torque_targets.csv", torqueRecords, Object.keys(torqueRecords[0]));
await writeCsv("same_model_fill_candidates.csv", sameModelCandidates, Object.keys(sameModelCandidates[0]));
await writeCsv("consumption_formula_candidates.csv", consumptionCandidates, Object.keys(consumptionCandidates[0]));
await writeCsv("battery_80kwh_priority.csv", packs80, Object.keys(packs80[0]));
await writeCsv("motor_torque_research_groups.csv", unresolvedMotorGroups, Object.keys(unresolvedMotorGroups[0]));
await writeCsv("battery_80kwh_research_groups.csv", batteryGroups, Object.keys(batteryGroups[0]));
await writeCsv("contaminated_cells.csv", contaminationRecords, Object.keys(contaminationRecords[0]));

const summary = {
  schema_version: "2.0.0",
  copyright: COPYRIGHT,
  source_file: SOURCE,
  sheet: sheet.name,
  records: rows.length,
  columns: headers.length,
  torque: {
    states: torqueStateCounts,
    targets: torqueRecords.length,
    same_model_candidates: torqueSameModelCandidates,
    motor_model_candidates: torqueMotorModelCandidates,
    internally_resolved_candidates: torqueInternalResolved,
    remaining_for_external_ocr: torqueRecords.length - torqueInternalResolved,
  },
  same_model_fill_candidates: sameModelCandidates.length,
  same_model_by_field: Object.fromEntries(
    [...sameModelCandidates.reduce((map, row) => map.set(row.field, (map.get(row.field) || 0) + 1), new Map())]
      .sort((a, b) => b[1] - a[1]),
  ),
  consumption_formula_candidates: consumptionCandidates.length,
  battery_80kwh_records: packs80.length,
  battery_80kwh_mass_derivable: packs80.filter((row) => row.derived_pack_mass_kg !== "").length,
  battery_80kwh_research_groups: batteryGroups.length,
  unresolved_motor_research_groups: unresolvedMotorGroups.length,
  contaminated_cells: contaminationRecords.length,
  contamination_by_field: Object.fromEntries(
    [...contaminationRecords.reduce((map, row) => map.set(row.field, (map.get(row.field) || 0) + 1), new Map())]
      .sort((a, b) => b[1] - a[1]),
  ),
  guardrails: [
    "不覆盖原始非空值",
    "同型号补全仅在全部已知值唯一一致时生成 C 级候选",
    "电耗仅用同一行容量与纯电续航计算",
    "80±5 kWh 包质量仅用同一行容量与系统能量密度推导，并标记非称重实测",
    "电池包尺寸必须有直接 L×W×H 证据",
    "营销用 800V 不视为精确系统电压",
  ],
};
await fs.writeFile(path.join(OUTPUT_DIR, "inventory_summary.json"), `${JSON.stringify(summary, null, 2)}\n`, "utf8");
await fs.writeFile(path.join(OUTPUT_DIR, "analysis_records.json"), `${JSON.stringify({
  copyright: COPYRIGHT,
  contamination_records: contaminationRecords,
  same_model_candidates: sameModelCandidates,
  consumption_candidates: consumptionCandidates,
  torque_records: torqueRecords,
  battery_80kwh_records: packs80,
  motor_research_groups: unresolvedMotorGroups,
  battery_research_groups: batteryGroups,
}, null, 2)}\n`, "utf8");
console.log(JSON.stringify(summary, null, 2));
