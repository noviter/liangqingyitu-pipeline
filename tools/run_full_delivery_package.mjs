import fs from "node:fs/promises";
import path from "node:path";
import { spawn } from "node:child_process";
import sharp from "sharp";
import { FileBlob, SpreadsheetFile, Workbook } from "@oai/artifact-tool";

function optionalArgument(name, fallback = "") {
  const index = process.argv.indexOf(name);
  return index === -1 ? fallback : process.argv[index + 1] ?? fallback;
}

function argument(name) {
  const value = optionalArgument(name);
  if (!value) throw new Error(`Missing ${name}`);
  return value;
}

function stamp(date = new Date()) {
  const pad = (value) => String(value).padStart(2, "0");
  return `${date.getFullYear()}${pad(date.getMonth() + 1)}${pad(date.getDate())}_${pad(date.getHours())}${pad(date.getMinutes())}${pad(date.getSeconds())}`;
}

function clean(value) {
  return String(value ?? "").replace(/\r/g, "").replace(/[ \t]+/g, " ").trim();
}

async function runNode(script, args) {
  await new Promise((resolve, reject) => {
    const child = spawn(process.execPath, [script, ...args], { stdio: "inherit", shell: false });
    child.on("error", reject);
    child.on("exit", (code) => {
      if (code === 0) resolve();
      else reject(new Error(`${script} exited with code ${code}`));
    });
  });
}

async function copyFile(source, destination) {
  await fs.mkdir(path.dirname(destination), { recursive: true });
  await fs.copyFile(source, destination);
}

async function createRevisionTemplate(destination, department) {
  const workbook = Workbook.create();
  const sheet = workbook.worksheets.add("人工修订意见");
  const headers = ["处室名称", "阶段", "事项", "需修订位置", "修订意见", "依据材料", "提出人", "处理状态"];
  const rows = [
    [department, "阶段1/阶段2/阶段3", "", "如：E列数字化结论、F列流程、泳道图节点", "", "", "", "待处理"],
  ];
  sheet.getRange("A1:H2").values = [headers, ...rows];
  sheet.getRange("A1:H1").format = {
    fill: "#BFBFBF",
    font: { name: "Microsoft YaHei", size: 10, bold: true, color: "#000000" },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    wrapText: true,
    borders: { preset: "all", style: "thin", color: "#000000" },
  };
  sheet.getRange("A2:H2").format = {
    font: { name: "Microsoft YaHei", size: 10, color: "#000000" },
    verticalAlignment: "top",
    wrapText: true,
    borders: { preset: "all", style: "thin", color: "#000000" },
  };
  sheet.getRange("A:A").format.columnWidth = 20;
  sheet.getRange("B:B").format.columnWidth = 18;
  sheet.getRange("C:C").format.columnWidth = 42;
  sheet.getRange("D:D").format.columnWidth = 36;
  sheet.getRange("E:E").format.columnWidth = 54;
  sheet.getRange("F:F").format.columnWidth = 44;
  sheet.getRange("G:H").format.columnWidth = 16;
  sheet.freezePanes.freezeRows(1);
  sheet.showGridLines = false;
  workbook.recalculate();
  const file = await SpreadsheetFile.exportXlsx(workbook);
  await file.save(destination);
}

async function inspectWorkbook(filePath) {
  const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(filePath));
  const sheet = workbook.worksheets.getItemAt(0);
  const values = sheet.getUsedRange().values;
  const headers = values[0].map(clean);
  const statusIndex = headers.indexOf("有没有数字化");
  const statuses = {};
  for (const row of values.slice(1)) {
    const status = clean(row[statusIndex]);
    if (status) statuses[status] = (statuses[status] || 0) + 1;
  }
  return { sheet: sheet.name, rowCount: values.length - 1, headers, statuses };
}

async function inspectImages(imageDir) {
  const files = (await fs.readdir(imageDir)).filter((name) => name.toLowerCase().endsWith(".png"));
  const images = [];
  for (const file of files) {
    const fullPath = path.join(imageDir, file);
    const stat = await fs.stat(fullPath);
    const metadata = await sharp(fullPath).metadata();
    images.push({ file, bytes: stat.size, width: metadata.width, height: metadata.height });
  }
  return images;
}

async function main() {
  const department = argument("--department");
  const input = argument("--input");
  const index = optionalArgument("--index", "工作区/系统资料索引.json");
  const runsDir = optionalArgument("--runs-dir", "runs");
  const runDate = optionalArgument("--run-date", new Date().toISOString().slice(0, 10));
  const stage2Mode = optionalArgument("--stage2-mode", "auto");
  const systemCodes = optionalArgument("--system-codes");
  const runDir = path.join(runsDir, `${department}_${stamp()}`);
  const internalDir = path.join(runDir, "_internal");
  const inspectDir = path.join(internalDir, "inspect");

  await fs.mkdir(inspectDir, { recursive: true });

  await runNode("tools/run_stage1_matter_decomposition.mjs", [
    "--input", input,
    "--output-dir", internalDir,
    "--department", department,
    "--run-date", runDate,
  ]);
  const rawStage1 = path.join(internalDir, `01_${department}_三定方案业务事项清单_人工核验前.xlsx`);
  const stage1 = path.join(internalDir, "阶段1_事项清单.xlsx");
  await copyFile(rawStage1, stage1);

  const context = path.join(internalDir, "阶段2_上下文检索包.json");
  const retrieveArgs = [
    "--input", stage1,
    "--index", index,
    "--output", context,
    "--max-candidates", "14",
  ];
  if (systemCodes) retrieveArgs.push("--system-codes", systemCodes);
  await runNode("tools/retrieve_stage2_context.mjs", retrieveArgs);

  const analysis = path.join(internalDir, "阶段2_结构化分析.json");
  await runNode("tools/run_stage2_prompt_analysis.mjs", [
    "--context", context,
    "--mode", stage2Mode,
    "--output", analysis,
  ]);

  const stage2 = path.join(internalDir, "阶段2_数字化与业务流程.xlsx");
  await runNode("tools/run_stage2_digital_process.mjs", [
    "--input", stage1,
    "--analysis-path", analysis,
    "--department", department,
    "--output", stage2,
  ]);

  await runNode("tools/run_stage3_swimlane.mjs", [
    "--input", stage2,
    "--output-dir", internalDir,
    "--department", department,
  ]);
  const rawDrawio = path.join(internalDir, `03_${department}.drawio`);
  const rawStage3 = path.join(internalDir, `03_${department}_泳道图核验版.xlsx`);
  const rawPending = path.join(internalDir, `03_${department}_泳道图待核验清单.xlsx`);
  const stage3Drawio = path.join(internalDir, "阶段3.drawio");
  const stage3 = path.join(internalDir, "阶段3_泳道图核验版.xlsx");
  const stage3Pending = path.join(internalDir, "阶段3_泳道图待核验清单.xlsx");
  await copyFile(rawDrawio, stage3Drawio);
  await copyFile(rawStage3, stage3);
  await copyFile(rawPending, stage3Pending);

  const delivery = path.join(runDir, `${department}_两清一图_交付版.xlsx`);
  const revision = path.join(runDir, "人工修订意见模板.xlsx");
  await copyFile(stage3, delivery);
  await createRevisionTemplate(revision, department);

  const drawioText = await fs.readFile(stage3Drawio, "utf8");
  const imageDir = path.join(internalDir, "泳道图预览");
  const stage2Inspect = await inspectWorkbook(stage2);
  const deliveryInspect = await inspectWorkbook(delivery);
  const imageInspect = await inspectImages(imageDir);
  const analysisItems = JSON.parse(await fs.readFile(analysis, "utf8"));
  const summary = {
    department,
    runDir,
    delivery,
    revision,
    internalDir,
    stage1: await inspectWorkbook(stage1),
    stage2: stage2Inspect,
    deliveryWorkbook: deliveryInspect,
    stage2StatusCounts: (analysisItems.items ?? analysisItems).reduce((acc, item) => {
      acc[item.status] = (acc[item.status] || 0) + 1;
      return acc;
    }, {}),
    stage3: {
      drawioPages: (drawioText.match(/<diagram /g) ?? []).length,
      previewImageCount: imageInspect.length,
      nonEmptyPreviewImages: imageInspect.filter((item) => item.bytes > 0 && item.width && item.height).length,
      sampledImages: imageInspect.slice(0, 3),
    },
  };
  await fs.writeFile(path.join(inspectDir, "运行包检查摘要.json"), JSON.stringify(summary, null, 2), "utf8");
  console.log(JSON.stringify(summary, null, 2));
}

await main();
