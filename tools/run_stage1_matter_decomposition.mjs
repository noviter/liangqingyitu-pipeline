import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile, Workbook } from "@oai/artifact-tool";

function readArgument(name) {
  const index = process.argv.indexOf(name);
  if (index === -1 || !process.argv[index + 1]) throw new Error(`Missing required argument: ${name}`);
  return process.argv[index + 1];
}

const inputPath = readArgument("--input");
const outputDir = readArgument("--output-dir");
const expectedDepartment = readArgument("--department");
const runDate = readArgument("--run-date");
const headers = ["处室名称", "职责", "事项", "职能类型"];
const fontName = "Microsoft YaHei";
const headerFill = "#BFBFBF";
const groupFill = "#F2F2F2";
const black = "#000000";

function cleanText(value) {
  return String(value ?? "").replace(/\r/g, "").replace(/[ \t]+/g, " ").trim();
}

function trimTerminalPunctuation(text) {
  return text.replace(/[；。]$/u, "");
}

function extractFirstSheetContext(values) {
  let unit = "";
  let department = "";
  let responsibilityText = "";
  for (let row = 0; row < values.length; row += 1) {
    for (let column = 0; column < values[row].length; column += 1) {
      const cell = cleanText(values[row][column]);
      if (!cell) continue;
      if (cell === "单位" && values[row][column + 1]) unit = cleanText(values[row][column + 1]);
      if ((cell === "部门" || cell === "处室") && values[row][column + 1]) department = cleanText(values[row][column + 1]);
      if (cell.includes("处室职能")) {
        for (let nextRow = row + 1; nextRow < values.length; nextRow += 1) {
          for (let nextColumn = 0; nextColumn < values[nextRow].length; nextColumn += 1) {
            const candidate = cleanText(values[nextRow][nextColumn]);
            if (candidate.includes("（一）") || candidate.includes("(一)")) {
              responsibilityText = candidate;
              break;
            }
          }
          if (responsibilityText) break;
        }
      }
    }
  }
  return { unit, department, responsibilityText };
}

function splitResponsibilities(sourceText) {
  return sourceText.split(/(?=（[一二三四五六七八九十]+）)/u).map(cleanText).filter(Boolean);
}

function classifyMatter(matter) {
  if (/^(参与拟订|参与编制|参与)/u.test(matter)) return "参与";
  if (/^(组织实施|组织开展)/u.test(matter)) return "负责";
  if (/^(组织协调|统筹协调)/u.test(matter)) return "组织";
  if (/^(监督实施|监督执行|负责指导|承担指导|承担督导|承担监督)/u.test(matter)) return "督导";
  if (/(指导|检查|监督|督促|考核|备案|督导|督察)/u.test(matter)) return "督导";
  if (/(组织|统筹|牵头|召集|推进)/u.test(matter)) return "组织";
  if (/(参与|协助|辅助|配合|提出意见|提出建议)/u.test(matter)) return "参与";
  if (/(承担|承办|负责|实施|执行|拟订|编制|起草|制定|管理|建设|审批|贯彻落实|贯彻执行)/u.test(matter)) return "负责";
  return "";
}

function makeMatter(responsibility, matter) {
  const cleanedMatter = trimTerminalPunctuation(cleanText(matter));
  return { responsibility, matter: cleanedMatter, functionalType: classifyMatter(cleanedMatter) };
}

function decomposeResponsibility(responsibility) {
  const body = trimTerminalPunctuation(cleanText(responsibility)).replace(/^（[一二三四五六七八九十]+）/u, "");
  const matters = [];
  let match = body.match(/^贯彻执行(.+)，参与拟订(.+)并督促落实$/u);
  if (match) {
    matters.push(makeMatter(responsibility, `贯彻执行${match[1]}`));
    matters.push(makeMatter(responsibility, `参与拟订${match[2]}`));
    matters.push(makeMatter(responsibility, `督促落实${match[2]}`));
    return { matters, issue: "" };
  }
  match = body.match(/^承担编制(.+)、(.+项目)备案、计划、绩效考核等事务性工作$/u);
  if (match) {
    matters.push(makeMatter(responsibility, `承担编制${match[1]}事务性工作`));
    matters.push(makeMatter(responsibility, `承担${match[2]}备案事务性工作`));
    matters.push(makeMatter(responsibility, `承担${match[2]}计划事务性工作`));
    matters.push(makeMatter(responsibility, `承担${match[2]}绩效考核事务性工作`));
    return { matters, issue: "" };
  }
  match = body.match(/^承担指导(.+公路)(.+)、(.+)、(.+)和(.+)事务性工作$/u);
  if (match) {
    for (const item of [match[2], match[3], match[4], match[5]]) matters.push(makeMatter(responsibility, `承担指导${match[1]}${item}事务性工作`));
    return { matters, issue: "" };
  }
  match = body.match(/^承担指导(.+)、(.+)事务性工作$/u);
  if (match) {
    const scope = match[1].match(/^(.*公路)(.+)$/u);
    if (scope) {
      matters.push(makeMatter(responsibility, `承担指导${scope[1]}${scope[2]}事务性工作`));
      matters.push(makeMatter(responsibility, `承担指导${scope[1]}${match[2]}事务性工作`));
      return { matters, issue: "" };
    }
  }
  match = body.match(/^承担(.+)、(.+)、(.+)等事务性工作$/u);
  if (match) {
    const scope = match[1].match(/^(.*公路)(.+)$/u);
    if (scope) {
      matters.push(makeMatter(responsibility, `承担${scope[1]}${scope[2]}事务性工作`));
      matters.push(makeMatter(responsibility, `承担${scope[1]}${match[2]}事务性工作`));
      matters.push(makeMatter(responsibility, `承担${scope[1]}${match[3]}事务性工作`));
      return { matters, issue: "" };
    }
  }
  match = body.match(/^参与(.+公路)(.+)、(.+)及(.+)等事务性工作$/u);
  if (match) {
    for (const item of [match[2], match[3], match[4]]) matters.push(makeMatter(responsibility, `参与${match[1]}${item}事务性工作`));
    return { matters, issue: "" };
  }
  if (/^(承担|承办|负责|参与|协助|辅助|组织|统筹|贯彻执行)/u.test(body)) return { matters: [makeMatter(responsibility, body)], issue: "" };
  return { matters: [], issue: "无法从职责原文安全识别完整动作和业务对象，需人工确认拆解边界。" };
}

function applyTableStyle(sheet, lastRow, lastColumn, groups = []) {
  const header = sheet.getRange(`A1:${lastColumn}1`);
  header.format = { fill: headerFill, font: { name: fontName, size: 10, bold: true, color: black }, horizontalAlignment: "center", verticalAlignment: "center", wrapText: true, borders: { preset: "all", style: "thin", color: black } };
  header.format.rowHeight = 28;
  if (lastRow < 2) return;
  const body = sheet.getRange(`A2:${lastColumn}${lastRow}`);
  body.format = { font: { name: fontName, size: 10, color: black }, verticalAlignment: "center", wrapText: true, borders: { preset: "all", style: "thin", color: black } };
  sheet.getRange(`A2:A${lastRow}`).format.horizontalAlignment = "center";
  sheet.getRange(`D2:D${lastRow}`).format.horizontalAlignment = "center";
  sheet.getRange(`B2:C${lastRow}`).format.horizontalAlignment = "left";
  groups.forEach((group, index) => { if (index % 2 === 1) sheet.getRange(`A${group.start}:D${group.end}`).format.fill = groupFill; });
  for (let row = 2; row <= lastRow; row += 1) sheet.getRange(`A${row}:${lastColumn}${row}`).format.rowHeight = 42;
  sheet.freezePanes.freezeRows(1);
  sheet.showGridLines = false;
}

function setMatterColumnWidths(sheet) {
  sheet.getRange("A:A").format.columnWidth = 18;
  sheet.getRange("B:B").format.columnWidth = 58;
  sheet.getRange("C:C").format.columnWidth = 45;
  sheet.getRange("D:D").format.columnWidth = 12;
}

function createMatterWorkbook(records, department) {
  const workbook = Workbook.create();
  const sheet = workbook.worksheets.add("业务事项清单");
  const values = [headers, ...records.map((record) => [department, record.responsibility, record.matter, record.functionalType])];
  sheet.getRange(`A1:D${values.length}`).values = values;
  const groups = [];
  let start = 2;
  while (start <= records.length + 1) {
    let end = start;
    while (end < records.length + 1 && records[end - 2].responsibility === records[end - 1].responsibility) end += 1;
    groups.push({ start, end });
    start = end + 1;
  }
  applyTableStyle(sheet, values.length, "D", groups);
  setMatterColumnWidths(sheet);
  if (records.length > 1) sheet.mergeCells(`A2:A${records.length + 1}`);
  for (const group of groups) if (group.end > group.start) sheet.mergeCells(`B${group.start}:B${group.end}`);
  workbook.recalculate();
  return workbook;
}

function createPendingWorkbook(issues, department) {
  const workbook = Workbook.create();
  const sheet = workbook.worksheets.add("待确认说明");
  const pendingHeaders = ["处室名称", "职责", "事项或分类问题", "待确认原因"];
  const values = [pendingHeaders, ...issues.map((issue) => [department, issue.responsibility, issue.matter || "", issue.reason])];
  sheet.getRange(`A1:D${values.length}`).values = values;
  applyTableStyle(sheet, values.length, "D");
  sheet.getRange("A:A").format.columnWidth = 18;
  sheet.getRange("B:B").format.columnWidth = 58;
  sheet.getRange("C:C").format.columnWidth = 38;
  sheet.getRange("D:D").format.columnWidth = 48;
  workbook.recalculate();
  return workbook;
}

function createInputManifest(inputFile, sheetName, department, responsibilityCount, dateText) {
  const workbook = Workbook.create();
  const sheet = workbook.worksheets.add("输入资料清单");
  const manifestHeaders = ["职责文件", "工作表", "处室名称", "职责数量", "运行日期", "适用范围"];
  sheet.getRange("A1:F2").values = [manifestHeaders, [inputFile, sheetName, department, responsibilityCount, dateText, "阶段 1：三定方案业务拆解与分类。仅读取职责文件第一个工作表。"]];
  applyTableStyle(sheet, 2, "F");
  sheet.getRange("A:A").format.columnWidth = 48;
  sheet.getRange("B:B").format.columnWidth = 16;
  sheet.getRange("C:C").format.columnWidth = 20;
  sheet.getRange("D:D").format.columnWidth = 12;
  sheet.getRange("E:E").format.columnWidth = 14;
  sheet.getRange("F:F").format.columnWidth = 48;
  workbook.recalculate();
  return workbook;
}

async function saveWorkbook(workbook, destination) {
  const output = await SpreadsheetFile.exportXlsx(workbook);
  await output.save(destination);
}

async function validateWorkbook(filePath, expectedSheet, expectedHeaders, expectedRows, expectsFourColumns) {
  const imported = await SpreadsheetFile.importXlsx(await FileBlob.load(filePath));
  const sheet = imported.worksheets.getItem(expectedSheet);
  const values = sheet.getUsedRange().values;
  if (values[0].join("|") !== expectedHeaders.join("|")) throw new Error(`${path.basename(filePath)} has unexpected headers.`);
  if (expectsFourColumns && values[0].length !== 4) throw new Error(`${path.basename(filePath)} must contain exactly four columns.`);
  if (expectedRows !== null && values.length - 1 !== expectedRows) throw new Error(`${path.basename(filePath)} has an unexpected row count.`);
}

const sourceWorkbook = await SpreadsheetFile.importXlsx(await FileBlob.load(inputPath));
const firstSheet = sourceWorkbook.worksheets.getItemAt(0);
const firstSheetName = firstSheet.name;
const context = extractFirstSheetContext(firstSheet.getUsedRange().values);
if (!context.unit || !context.department || !context.responsibilityText) throw new Error("The first worksheet does not contain a readable unit, department, and responsibility text.");
if (context.department !== expectedDepartment) throw new Error(`Department mismatch: source is ${context.department}, argument is ${expectedDepartment}.`);

const responsibilities = splitResponsibilities(context.responsibilityText);
if (responsibilities.length === 0) throw new Error("No responsibilities were found in the first worksheet.");
const records = [];
const issues = [];
for (const responsibility of responsibilities) {
  const result = decomposeResponsibility(responsibility);
  if (result.issue) {
    issues.push({ responsibility, matter: "", reason: result.issue });
    continue;
  }
  for (const record of result.matters) {
    if (!record.functionalType) issues.push({ responsibility, matter: record.matter, reason: "无法依据原文和既定分类规则确定职能类型。" });
    else records.push(record);
  }
}
if (records.length === 0) throw new Error("No safely classified matters were produced.");

await fs.mkdir(outputDir, { recursive: true });
const matterPath = path.join(outputDir, `01_${expectedDepartment}_三定方案业务事项清单_人工核验前.xlsx`);
const pendingPath = path.join(outputDir, "阶段1待确认说明.xlsx");
const manifestPath = path.join(outputDir, "输入资料清单.xlsx");
const notesPath = path.join(outputDir, "阶段1运行说明.md");
await saveWorkbook(createMatterWorkbook(records, expectedDepartment), matterPath);
await saveWorkbook(createPendingWorkbook(issues, expectedDepartment), pendingPath);
await saveWorkbook(createInputManifest(inputPath, firstSheetName, expectedDepartment, responsibilities.length, runDate), manifestPath);
await fs.writeFile(notesPath, [
  "# 阶段 1 运行说明", "", `- 输入职责文件：${inputPath}`, `- 读取工作表：${firstSheetName}（第一个工作表）`, `- 单位：${context.unit}`, `- 处室：${expectedDepartment}`,
  `- 来源职责数量：${responsibilities.length}`, `- 生成事项数量：${records.length}`, `- 待人工确认问题数量：${issues.length}`,
  "- 适用范围：仅执行阶段 1 的三定方案业务拆解与清单整理；未读取人工样例业务内容，未使用系统资料。",
  "- 规则来源：PDF Prompt1、Prompt2 的规则与输出规范。",
  "- 停止状态：阶段 1 已停止，等待人工审核确认。阶段 2 和阶段 3 未执行。", "",
].join("\n"), "utf8");

await validateWorkbook(matterPath, "业务事项清单", headers, records.length, true);
await validateWorkbook(pendingPath, "待确认说明", ["处室名称", "职责", "事项或分类问题", "待确认原因"], issues.length, true);
await validateWorkbook(manifestPath, "输入资料清单", ["职责文件", "工作表", "处室名称", "职责数量", "运行日期", "适用范围"], 1, false);
console.log(JSON.stringify({ sourceWorksheet: firstSheetName, unit: context.unit, department: expectedDepartment, responsibilityCount: responsibilities.length, matterCount: records.length, pendingCount: issues.length, outputDir }, null, 2));
