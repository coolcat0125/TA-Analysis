import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const root = path.resolve(process.cwd(), "..");
const payload = JSON.parse(
  await fs.readFile(path.join(root, "data", "delivery_payload.json"), "utf8"),
);
const outputDir = path.join(
  root,
  "outputs",
  "01a147a8-938a-7ff3-a15c-94b7fcac7a5e",
);
await fs.mkdir(outputDir, { recursive: true });

const font = "Microsoft YaHei";
const colors = {
  ink: "#1F2937",
  title: "#111827",
  muted: "#6B7280",
  header: "#1F4E78",
  headerText: "#FFFFFF",
  accent: "#DCE6F1",
  band: "#F7F9FC",
  border: "#D9D9D9",
  note: "#64748B",
  missing: "#FFF2CC",
  good: "#E2F0D9",
};

const workbook = Workbook.create();

function styleTitle(sheet, title, context, width = 10) {
  sheet.showGridLines = false;
  sheet.getRange(`A2`).values = [[title]];
  sheet.getRange(`A2:A2`).format.font = {
    name: font,
    size: 16,
    bold: true,
    color: colors.title,
  };
  sheet.getRange(`A3`).values = [[context]];
  sheet.getRange(`A3:${String.fromCharCode(65 + width)}3`).format.font = {
    name: font,
    size: 10,
    italic: true,
    color: colors.muted,
  };
}

function styleHeader(range, fill = colors.header, textColor = colors.headerText) {
  range.format = {
    fill,
    font: { name: font, size: 10, bold: true, color: textColor },
    verticalAlignment: "center",
    horizontalAlignment: "center",
    wrapText: true,
    borders: { preset: "all", style: "thin", color: colors.border },
  };
}

function styleTable(range, options = {}) {
  const { numberFormats = {}, centerColumns = [], startColumn = "A" } = options;
  range.format = {
    font: { name: font, size: 10, color: colors.ink },
    verticalAlignment: "center",
    wrapText: true,
    borders: { preset: "all", style: "thin", color: colors.border },
  };
  for (const [col, format] of Object.entries(numberFormats)) {
    range.getColumn(columnNumber(col) - columnNumber(startColumn)).format.numberFormat = format;
  }
  for (const col of centerColumns) {
    range.getColumn(columnNumber(col) - columnNumber(startColumn)).format.horizontalAlignment = "center";
  }
}

function columnNumber(letter) {
  let value = 0;
  for (const char of letter) {
    value = value * 26 + char.charCodeAt(0) - 64;
  }
  return value;
}

function setWidths(sheet, widths) {
  for (const [column, width] of Object.entries(widths)) {
    sheet.getRange(`${column}:${column}`).format.columnWidth = width;
  }
}

function counts(records) {
  const route = { "纯电(BEV)": 0, "插电混动(PHEV)": 0, "增程(EREV)": 0 };
  for (const record of records) route[record["技术路线"]] += 1;
  return route;
}

function writeDetailSheet(name, records, tableName, tabColor) {
  const sheet = workbook.worksheets.add(name);
  sheet.tabColor = tabColor;
  styleTitle(
    sheet,
    `${name}逐车型扫描`,
    "乘用车口径；空白表示官方公开清单未披露该字段，不以其他来源补写。",
    30,
  );
  const headers = [
    "批次",
    "状态",
    "技术路线",
    "是否增程",
    "企业名称",
    "产品商标",
    "通用名称",
    "产品型号",
    "产品名称",
    "车长(mm)",
    "车宽(mm)",
    "车高(mm)",
    "轴距(mm)",
    "整备质量(kg)",
    "总质量(kg)",
    "最高车速(km/h)",
    "额定载客(人)",
    "纯电续航(km)",
    "电池容量(kWh)",
    "电池类型",
    "电芯供应商",
    "电池包供应商",
    "电机峰值功率数列(kW)",
    "电机总功率(kW)",
    "驱动形式",
    "发动机排量(mL)",
    "发动机功率(kW)",
    "综合油耗(L/100km)",
    "型号级新增",
    "来源URL",
    "缺失关键字段",
  ];
  sheet.getRange("A3:AE3").values = [headers];
  styleHeader(sheet.getRange("A3:AE3"));
  const rows = records.map((record) =>
    headers.map((header) => record[header] ?? null),
  );
  if (rows.length) {
    sheet.getRange(`A4:AE${3 + rows.length}`).values = rows;
    styleTable(sheet.getRange(`A4:AE${3 + rows.length}`), {
      numberFormats: {
        A: "0",
        J: "#,##0",
        K: "#,##0",
        L: "#,##0",
        M: "#,##0",
        N: "#,##0",
        O: "#,##0",
        P: "#,##0",
        R: "#,##0",
        S: "0.0",
        X: "0.0",
        Z: "#,##0",
        AA: "0.0",
        AB: "0.0",
      },
      centerColumns: ["A", "B", "C", "D", "P", "R", "S", "X", "Y", "AC"],
    });
    sheet
      .getRange(`X4:X${3 + rows.length}`)
      .conditionalFormats.add("cellIs", {
        operator: "greaterThanOrEqual",
        formula: 200,
        format: { fill: colors.good, font: { bold: true, color: "#1F2937" } },
      });
    sheet
      .getRange(`AE4:AE${3 + rows.length}`)
      .conditionalFormats.add("notContainsBlanks", {
        format: { fill: colors.missing, color: "#7C4A03" },
      });
  }
  sheet.freezePanes.freezeRows(3);
  sheet.freezePanes.freezeColumns(9);
  setWidths(sheet, {
    A: 7,
    B: 10,
    C: 16,
    D: 9,
    E: 26,
    F: 18,
    G: 16,
    H: 20,
    I: 24,
    J: 10,
    K: 10,
    L: 10,
    M: 10,
    N: 12,
    O: 11,
    P: 11,
    Q: 12,
    R: 12,
    S: 12,
    T: 20,
    U: 30,
    V: 30,
    W: 20,
    X: 14,
    Y: 11,
    Z: 12,
    AA: 12,
    AB: 12,
    AC: 18,
    AD: 46,
    AE: 24,
  });
  return { sheet, dataStart: 4, dataEnd: 3 + rows.length, tableName };
}

const sheet411 = writeDetailSheet(
  "411正式车型",
  payload.records["411"],
  "Batch411",
  "#1F4E78",
);
const sheet412 = writeDetailSheet(
  "412公示车型",
  payload.records["412"],
  "Batch412",
  "#2E75B6",
);
const overall410 = payload.batch_summary.overall_410;

const summary = workbook.worksheets.add("扫描总览");
styleTitle(
  summary,
  "第411批正式公告与第412批公示 NEV乘用车扫描",
  "数据截至 2026-10-17；411为正式公告，412为公示候选口径。",
  11,
);
summary.tabColor = "#0F4C81";

summary.getRange("A5:K5").values = [
  [
    "批次",
    "状态",
    "日期",
    "NEV乘用车数",
    "纯电",
    "插混/增程",
    "型号级新增",
    "平均车长(mm)",
    "平均电机功率(kW)",
    "平均续航(km)",
    "平均电池容量(kWh)",
  ],
];
styleHeader(summary.getRange("A5:K5"));
summary.getRange("A6:K8").values = [
  [
    410,
    "正式公告基准",
    "2026-09-18",
    overall410["数量"],
    47,
    23,
    null,
    overall410["平均车长(mm)"],
    overall410["平均电机功率(kW)"],
    overall410["平均续航(km)"],
    overall410["平均电池容量(kWh)"],
  ],
  [411, "正式公告", "2026-10-16", null, null, null, null, null, null, null, null],
  [412, "公示", "2026-10-16", null, null, null, null, null, null, null, null],
];
const last411 = sheet411.dataEnd;
const last412 = sheet412.dataEnd;
const formula411 = (formula) => `=${formula.replaceAll("R", "'411正式车型'!$R$4:$R$" + last411)}`;
summary.getRange("D7").formulas = [[`=COUNTA('411正式车型'!$H$4:$H$${last411})`]];
summary.getRange("E7").formulas = [[`=COUNTIF('411正式车型'!$C$4:$C$${last411},"纯电(BEV)")`]];
summary.getRange("F7").formulas = [[`=COUNTIF('411正式车型'!$C$4:$C$${last411},"插电混动(PHEV)")+COUNTIF('411正式车型'!$C$4:$C$${last411},"增程(EREV)")`]];
summary.getRange("G7").formulas = [[`=COUNTIF('411正式车型'!$AC$4:$AC$${last411},"型号级新增")`]];
summary.getRange("H7").formulas = [[`=IF(COUNT('411正式车型'!$J$4:$J$${last411})=0,"n.a.",AVERAGE('411正式车型'!$J$4:$J$${last411}))`]];
summary.getRange("I7").formulas = [[`=IF(COUNT('411正式车型'!$X$4:$X$${last411})=0,"n.a.",AVERAGE('411正式车型'!$X$4:$X$${last411}))`]];
summary.getRange("J7").formulas = [[`=IF(COUNT('411正式车型'!$R$4:$R$${last411})=0,"n.a.",AVERAGE('411正式车型'!$R$4:$R$${last411}))`]];
summary.getRange("K7").formulas = [[`=IF(COUNT('411正式车型'!$S$4:$S$${last411})=0,"n.a.",AVERAGE('411正式车型'!$S$4:$S$${last411}))`]];
summary.getRange("D8").formulas = [[`=COUNTA('412公示车型'!$H$4:$H$${last412})`]];
summary.getRange("E8").formulas = [[`=COUNTIF('412公示车型'!$C$4:$C$${last412},"纯电(BEV)")`]];
summary.getRange("F8").formulas = [[`=COUNTIF('412公示车型'!$C$4:$C$${last412},"插电混动(PHEV)")+COUNTIF('412公示车型'!$C$4:$C$${last412},"增程(EREV)")`]];
summary.getRange("G8").formulas = [[`=COUNTIF('412公示车型'!$AC$4:$AC$${last412},"型号级新增")`]];
summary.getRange("H8").formulas = [[`=IF(COUNT('412公示车型'!$J$4:$J$${last412})=0,"n.a.",AVERAGE('412公示车型'!$J$4:$J$${last412}))`]];
summary.getRange("I8").formulas = [[`=IF(COUNT('412公示车型'!$X$4:$X$${last412})=0,"n.a.",AVERAGE('412公示车型'!$X$4:$X$${last412}))`]];
summary.getRange("J8").formulas = [[`=IF(COUNT('412公示车型'!$R$4:$R$${last412})=0,"n.a.",AVERAGE('412公示车型'!$R$4:$R$${last412}))`]];
summary.getRange("K8").formulas = [[`=IF(COUNT('412公示车型'!$S$4:$S$${last412})=0,"n.a.",AVERAGE('412公示车型'!$S$4:$S$${last412}))`]];
styleTable(summary.getRange("A6:K8"), {
  numberFormats: { D: "#,##0", E: "#,##0", F: "#,##0", G: "#,##0", H: "#,##0.0", I: "#,##0.0", J: "#,##0.0", K: "#,##0.0" },
  centerColumns: ["A", "B", "C", "D", "E", "F", "G"],
});

summary.getRange("A10:C10").values = [["技术路线", "410正式基准", "411正式 / 412公示"]];
styleHeader(summary.getRange("A10:C10"));
summary.getRange("A11:C13").values = [
  ["纯电(BEV)", "47 / 606.4km / 73.8kWh", "411：23 / 593.5km / 72.4kWh；412：37 / 待正式公告"],
  ["插电混动(PHEV)", "16 / 224.9km / 43.4kWh", "411：21 / 248.5km / 55.5kWh；412：15 / 225.1km / 48.3kWh"],
  ["增程(EREV)", "7 / 255.1km / 54.5kWh", "411：3 / 239.0km / 50.1kWh；412：6 / 257.0km / 54.7kWh"],
];
styleTable(summary.getRange("A11:C13"), { centerColumns: ["A"] });

summary.getRange("E10:L10").values = [[
  "判断",
  "说明",
  "",
  "",
  "",
  "",
  "",
  "",
]];
styleHeader(summary.getRange("E10:F10"));
summary.getRange("E11:F15").values = [
  ["结构判断", "412公示的纯电占比更高，但续航与电池容量要等正式购置税目录后才能定论。"],
  ["参数判断", "411正式公告的PHEV平均纯电续航和电池容量明显高于410基准，长续航插混仍在加码。"],
  ["驱动判断", "412纯电样本平均电机功率约270.5kW，200kW及以上车型占有效功率样本的70.3%。"],
  ["数据缺口", "412公示纯电车型尚未统一披露续航、容量和电池供应商，不能按0处理。"],
  ["建议", "下一轮按正式公告附件更新412，并单独跟踪功率、双电机/多电机和电池包供应商。"],
];
styleTable(summary.getRange("E11:F15"), { centerColumns: ["E"], startColumn: "E" });
summary.getRange("A17:J17").values = [[
  "技术路线",
  "410数量",
  "411数量",
  "412数量",
  "410续航",
  "411续航",
  "412续航",
  "410容量",
  "411容量",
  "412容量",
]];
styleHeader(summary.getRange("A17:J17"));
summary.getRange("A18:A20").values = [["纯电(BEV)"], ["插电混动(PHEV)"], ["增程(EREV)"]];
const routeRows = [
  { route: "纯电(BEV)", row: 18, b410: 47, r410: 606.4, c410: 73.8 },
  { route: "插电混动(PHEV)", row: 19, b410: 16, r410: 224.9, c410: 43.4 },
  { route: "增程(EREV)", row: 20, b410: 7, r410: 255.1, c410: 54.5 },
];
for (const item of routeRows) {
  summary.getRange(`B${item.row}:D${item.row}`).values = [[item.b410, null, null]];
  summary.getRange(`B${item.row}`).values = [[item.b410]];
  summary.getRange(`C${item.row}`).formulas = [[`=COUNTIF('411正式车型'!$C$4:$C$${last411},"${item.route}")`]];
  summary.getRange(`D${item.row}`).formulas = [[`=COUNTIF('412公示车型'!$C$4:$C$${last412},"${item.route}")`]];
  summary.getRange(`E${item.row}`).values = [[item.r410]];
  summary.getRange(`F${item.row}`).formulas = [[`=IF(COUNTIFS('411正式车型'!$C$4:$C$${last411},"${item.route}",'411正式车型'!$R$4:$R$${last411},">0")=0,"n.a.",SUMIFS('411正式车型'!$R$4:$R$${last411},'411正式车型'!$C$4:$C$${last411},"${item.route}",'411正式车型'!$R$4:$R$${last411},">0")/COUNTIFS('411正式车型'!$C$4:$C$${last411},"${item.route}",'411正式车型'!$R$4:$R$${last411},">0"))`]];
  summary.getRange(`G${item.row}`).formulas = [[`=IF(COUNTIFS('412公示车型'!$C$4:$C$${last412},"${item.route}",'412公示车型'!$R$4:$R$${last412},">0")=0,"n.a.",SUMIFS('412公示车型'!$R$4:$R$${last412},'412公示车型'!$C$4:$C$${last412},"${item.route}",'412公示车型'!$R$4:$R$${last412},">0")/COUNTIFS('412公示车型'!$C$4:$C$${last412},"${item.route}",'412公示车型'!$R$4:$R$${last412},">0"))`]];
  summary.getRange(`H${item.row}`).values = [[item.c410]];
  summary.getRange(`I${item.row}`).formulas = [[`=IF(COUNTIFS('411正式车型'!$C$4:$C$${last411},"${item.route}",'411正式车型'!$S$4:$S$${last411},">0")=0,"n.a.",SUMIFS('411正式车型'!$S$4:$S$${last411},'411正式车型'!$C$4:$C$${last411},"${item.route}",'411正式车型'!$S$4:$S$${last411},">0")/COUNTIFS('411正式车型'!$C$4:$C$${last411},"${item.route}",'411正式车型'!$S$4:$S$${last411},">0"))`]];
  summary.getRange(`J${item.row}`).formulas = [[`=IF(COUNTIFS('412公示车型'!$C$4:$C$${last412},"${item.route}",'412公示车型'!$S$4:$S$${last412},">0")=0,"n.a.",SUMIFS('412公示车型'!$S$4:$S$${last412},'412公示车型'!$C$4:$C$${last412},"${item.route}",'412公示车型'!$S$4:$S$${last412},">0")/COUNTIFS('412公示车型'!$C$4:$C$${last412},"${item.route}",'412公示车型'!$S$4:$S$${last412},">0"))`]];
}
styleTable(summary.getRange("A18:J20"), {
  numberFormats: { B: "#,##0", C: "#,##0", D: "#,##0", E: "#,##0.0", F: "#,##0.0", G: "#,##0.0", H: "#,##0.0", I: "#,##0.0", J: "#,##0.0" },
  centerColumns: ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J"],
});
summary.getRange("A22").values = [["备注：410为权威底表基准；411为正式公告附件与官方车型公示详情；412为公示候选数据。"]];
summary.getRange("A22").format.font = { name: font, size: 10, italic: true, color: colors.note };
setWidths(summary, { A: 18, B: 12, C: 14, D: 14, E: 16, F: 18, G: 18, H: 18, I: 18, J: 18, K: 17 });

const driver = workbook.worksheets.add("驱动系统对比");
styleTitle(
  driver,
  "批次驱动系统对比",
  "均值只基于官方公开字段；功率和驱动形式的有效样本数单独列示。",
  12,
);
driver.tabColor = "#4472C4";
const driverHeaders = [
  "批次",
  "技术路线",
  "数量",
  "功率有效样本",
  "平均电机功率(kW)",
  "200kW及以上占比",
  "多电机有效样本",
  "多电机占比",
  "平均续航(km)",
  "平均电池容量(kWh)",
  "平均车长(mm)",
  "平均整备质量(kg)",
];
driver.getRange("A3:L3").values = [driverHeaders];
styleHeader(driver.getRange("A3:L3"));
const driverRows = [];
for (const batch of ["410", "411", "412"]) {
  for (const row of payload.batch_summary[batch] || []) {
    driverRows.push([
      Number(batch),
      row["技术路线"],
      row["数量"],
      row["功率有效样本"] ?? null,
      row["平均电机功率(kW)"],
      row["200kW及以上占比"],
      row["多电机有效样本"] ?? null,
      row["多电机占比"],
      row["平均续航(km)"],
      row["平均电池容量(kWh)"],
      row["平均车长(mm)"],
      row["平均整备质量(kg)"],
    ]);
  }
}
driver.getRange(`A4:L${3 + driverRows.length}`).values = driverRows;
styleTable(driver.getRange(`A4:L${3 + driverRows.length}`), {
  numberFormats: {
    A: "0",
    C: "#,##0",
    D: "#,##0",
    E: "#,##0.0",
    F: "0.0%",
    G: "#,##0",
    H: "0.0%",
    I: "#,##0.0",
    J: "#,##0.0",
    K: "#,##0.0",
    L: "#,##0.0",
  },
  centerColumns: ["A", "B", "C", "D", "F", "G", "H"],
});
driver.freezePanes.freezeRows(3);
setWidths(driver, { A: 8, B: 16, C: 9, D: 13, E: 16, F: 15, G: 15, H: 13, I: 14, J: 16, K: 14, L: 16 });

const evidence = workbook.worksheets.add("证据与口径");
styleTitle(evidence, "证据与口径", "记录官方来源、字段覆盖和正式公告更新要求。", 5);
evidence.tabColor = "#A6A6A6";
evidence.getRange("A3:E3").values = [["项目", "内容", "数值/说明", "来源", "备注"]];
styleHeader(evidence.getRange("A3:E3"));
const evidenceRows = [
  ["411状态", "正式公告", "2026-10-16发布，公告成文日期2026-10-15", "工业和信息化部", "正式公告"],
  ["412状态", "公示", "2026-10-17至2026-10-23", "工业和信息化部装备工业一司", "等待正式公告"],
  ["411总产品申报", "1019个新产品申报；其中新能源288个型号", "汽车产品607个、摩托车412个", "第411批公示通知", "本表仅取乘用车"],
  ["412总产品申报", "1074个新产品申报；新能源349个型号", "汽车产品687个、摩托车387个", "第412批公示通知", "本表识别58个NEV乘用车型号"],
  ["411乘用车扫描", "47个NEV乘用车型号", "纯电23、插混21、增程3", "411正式公告附件+官方详情页", "型号级口径"],
  ["412乘用车扫描", "58个NEV乘用车型号", "纯电37、插混15、增程6", "412公示清单+官方详情页", "公示候选口径"],
  ["412续航/容量", "未统一披露", "正式购置税目录发布后补齐", "工信部", "空白不等于0"],
  ["电机功率覆盖", "411：28/47；412：50/58", "无公开功率字段的车型保留空白", "官方详情页", "均值只按有值样本"],
];
evidence.getRange(`A4:E${3 + evidenceRows.length}`).values = evidenceRows;
styleTable(evidence.getRange(`A4:E${3 + evidenceRows.length}`), { centerColumns: ["A"] });
evidence.getRange("A13").values = [["字段口径"]];
evidence.getRange("A13:E13").format.font = { name: font, size: 12, bold: true, color: colors.title };
evidence.getRange("A14:E20").values = [
  ["技术路线", "纯电(BEV)、插电混动(PHEV)、增程(EREV)分别统计", "EREV按名称或型号中的增程/REEV识别", "本表口径", "避免把PHEV和EREV混为一组"],
  ["车长/轴距/整备", "取公告详情页首个或主值", "多配置车型保留原始配置文本", "公告详情页", "均值只用于批次趋势"],
  ["功率", "电机总功率按公开的电机峰值功率序列求和", "若只披露额定功率则仅保留原字段", "公告详情页", "不按媒体参数补写"],
  ["电池", "容量优先取正式公告附件；类型和供应商取公示详情页", "412纯电容量待正式公告补齐", "公告附件/详情页", "优先公示原文"],
  ["批次新增", "按产品型号与341~410权威底表比对", "型号未出现标记为型号级新增", "341~410底表", "不等同于全新车系"],
  ["缺失值", "n.a.或空白表示未披露或不可得", "不得将缺失值当0参与均值", "本表口径", "报告与底表一致"],
  ["更新要求", "正式公告发布后替换412公示数据", "同步更新报告、驱动对比和看板", "项目流程", "正式公告优先"],
];
styleTable(evidence.getRange("A14:E20"), { centerColumns: ["A"] });
setWidths(evidence, { A: 16, B: 30, C: 42, D: 24, E: 30 });

workbook.recalculate();

const inspect = await workbook.inspect({
  kind: "workbook,sheet,table",
  maxChars: 5000,
  tableMaxRows: 8,
  tableMaxCols: 10,
});
console.log(inspect.ndjson);

for (const sheetName of workbook.worksheets.items.map((sheet) => sheet.name)) {
  const preview = await workbook.render({
    sheetName,
    autoCrop: "all",
    scale: 1,
    format: "png",
  });
  await fs.writeFile(
    path.join(outputDir, `${sheetName}.png`),
    new Uint8Array(await preview.arrayBuffer()),
  );
}

const output = await SpreadsheetFile.exportXlsx(workbook);
const target = path.join(outputDir, "411-412公告NEV车型扫描底表.xlsx");
await output.save(target);
console.log(`saved ${target}`);
