import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile, Workbook } from "@oai/artifact-tool";

function argument(name) {
  const index = process.argv.indexOf(name);
  if (index === -1 || !process.argv[index + 1]) throw new Error(`Missing ${name}`);
  return process.argv[index + 1];
}

const inputPath = argument("--input");
const analysisPath = argument("--analysis-path");
const outputPath = argument("--output");
const department = argument("--department");
const headers = ["处室名称", "职责", "事项", "职能类型", "有没有数字化", "流程/待确认问题", "泳道图", "系统（标号）", "流程依据"];
const allowedStatuses = new Set(["有", "待确认", "无", "不纳入本次梳理范畴"]);
const font = "Microsoft YaHei";
const black = "#000000";
const headerFill = "#BFBFBF";
const groupFill = "#F2F2F2";

function clean(value) { return String(value ?? "").replace(/\r/g, "").replace(/[ \t]+/g, " ").trim(); }
function formatBasis(value) {
  return clean(value).replace(/\s*(?=节点\s*\d+【)/gu, "\n").replace(/\n{2,}/g, "\n").trim();
}

function applyStyle(sheet, lastRow, groups, records) {
  const header = sheet.getRange("A1:I1");
  header.format = { fill: headerFill, font: { name: font, size: 10, bold: true, color: black }, horizontalAlignment: "center", verticalAlignment: "center", wrapText: true, borders: { preset: "all", style: "thin", color: black } };
  header.format.rowHeight = 30;
  const body = sheet.getRange(`A2:I${lastRow}`);
  body.format = { font: { name: font, size: 10, color: black }, verticalAlignment: "top", wrapText: true, borders: { preset: "all", style: "thin", color: black } };
  sheet.getRange(`A2:A${lastRow}`).format.horizontalAlignment = "center";
  sheet.getRange(`D2:E${lastRow}`).format.horizontalAlignment = "center";
  groups.forEach((group, index) => { if (index % 2 === 1) sheet.getRange(`A${group.start}:I${group.end}`).format.fill = groupFill; });
  for (let row = 2; row <= lastRow; row += 1) {
    const record = records[row - 2];
    const longestText = Math.max(record.process.length, record.basis.length);
    sheet.getRange(`A${row}:I${row}`).format.rowHeight = Math.max(180, Math.min(620, 120 + Math.ceil(longestText / 70) * 18));
  }
  sheet.getRange("A:A").format.columnWidth = 18;
  sheet.getRange("B:B").format.columnWidth = 54;
  sheet.getRange("C:C").format.columnWidth = 42;
  sheet.getRange("D:D").format.columnWidth = 12;
  sheet.getRange("E:E").format.columnWidth = 15;
  sheet.getRange("F:F").format.columnWidth = 64;
  sheet.getRange("G:G").format.columnWidth = 16;
  sheet.getRange("H:H").format.columnWidth = 36;
  sheet.getRange("I:I").format.columnWidth = 78;
  sheet.freezePanes.freezeRows(1);
  sheet.showGridLines = false;
}

const input = await SpreadsheetFile.importXlsx(await FileBlob.load(inputPath));
const source = input.worksheets.getItemAt(0).getUsedRange().values;
if (source[0].map(clean).join("|") !== headers.slice(0, 4).join("|")) throw new Error("Stage 1 reviewed Excel must contain A-D headers.");
const analysisPayload = JSON.parse(await fs.readFile(analysisPath, "utf8"));
const analysis = Array.isArray(analysisPayload) ? analysisPayload : analysisPayload.items;
if (!Array.isArray(analysis)) throw new Error("Internal analysis must be an array or { items: [...] }.");
const byMatter = new Map(analysis.map((item) => [clean(item.matter), item]));
let currentDepartment = "";
let currentResponsibility = "";
const records = [];
for (const row of source.slice(1)) {
  const matter = clean(row[2]);
  if (!matter) continue;
  currentDepartment = clean(row[0]) || currentDepartment;
  currentResponsibility = clean(row[1]) || currentResponsibility;
  if (currentDepartment !== department) throw new Error("Reviewed stage 1 Excel contains an unexpected department.");
  const item = byMatter.get(matter);
  const status = clean(item?.digital_status ?? item?.status);
  if (!item || !allowedStatuses.has(status)) throw new Error(`Missing or invalid prompt analysis for: ${matter}`);
  if (!clean(item.process) || !clean(item.basis)) throw new Error(`Process or basis is empty for: ${matter}`);
  if (status === "有" && !clean(item.systems)) throw new Error(`E=有 requires H systems for: ${matter}`);
  if (status === "有" && ((clean(item.basis).match(/节点\s*\d+【/gu) ?? []).length < 3 || !clean(item.basis).includes("《"))) throw new Error(`E=有 requires multiple sourced I-column nodes for: ${matter}`);
  if (/【(?:系统定位|资料范围核验|范围核验|边界说明|资料核查)】/u.test(clean(item.basis))) throw new Error(`I-column node title is too generic for: ${matter}`);
  const basis = formatBasis(item.basis);
  records.push({ department: currentDepartment, responsibility: currentResponsibility, matter, type: clean(row[3]), status, process: clean(item.process), systems: clean(item.systems), basis });
}
if (records.length !== byMatter.size) throw new Error("Analysis must contain exactly the current stage 1 matters.");
const groups = [];
let start = 2;
while (start <= records.length + 1) {
  let end = start;
  while (end < records.length + 1 && records[end - 2].responsibility === records[end - 1].responsibility) end += 1;
  groups.push({ start, end });
  start = end + 1;
}
const workbook = Workbook.create();
const sheet = workbook.worksheets.add("数字化与业务流程");
sheet.getRange(`A1:I${records.length + 1}`).values = [headers, ...records.map((r) => [r.department, r.responsibility, r.matter, r.type, r.status, r.process, "", r.systems, r.basis])];
applyStyle(sheet, records.length + 1, groups, records);
sheet.mergeCells(`A2:A${records.length + 1}`);
for (const group of groups) if (group.end > group.start) sheet.mergeCells(`B${group.start}:B${group.end}`);
workbook.recalculate();
await fs.mkdir(path.dirname(outputPath), { recursive: true });
const xlsx = await SpreadsheetFile.exportXlsx(workbook);
await xlsx.save(outputPath);
const reopened = await SpreadsheetFile.importXlsx(await FileBlob.load(outputPath));
const values = reopened.worksheets.getItem("数字化与业务流程").getUsedRange().values;
if (values.length - 1 !== records.length || values[0].map(clean).join("|") !== headers.join("|")) throw new Error("Export validation failed.");
if (values.slice(1).some((row) => !allowedStatuses.has(clean(row[4])) || clean(row[6]) !== "" || !clean(row[5]) || !clean(row[8]))) throw new Error("Stage 2 field validation failed.");
const statusCounts = records.reduce((acc, item) => {
  acc[item.status] = (acc[item.status] || 0) + 1;
  return acc;
}, {});
console.log(JSON.stringify({ matterCount: records.length, statusCounts, outputPath }, null, 2));
