import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

function optionalArgument(name, fallback = "") {
  const index = process.argv.indexOf(name);
  return index === -1 ? fallback : process.argv[index + 1] ?? fallback;
}

function argument(name) {
  const value = optionalArgument(name);
  if (!value) throw new Error(`Missing ${name}`);
  return value;
}

const inputPath = argument("--input");
const indexPath = optionalArgument("--index", "工作区/系统资料索引.json");
const outputPath = argument("--output");
const maxCandidates = Number(optionalArgument("--max-candidates", "14"));
const maxExcerptChars = Number(optionalArgument("--max-excerpt-chars", "7000"));
const systemCodes = new Set(optionalArgument("--system-codes").split(/[，,;；\s]+/u).filter(Boolean));
const matterRegex = optionalArgument("--matter-regex");
const matterFilter = matterRegex ? new RegExp(matterRegex, "u") : null;

function clean(value) {
  return String(value ?? "").replace(/\r/g, "").replace(/[ \t]+/g, " ").trim();
}

function basename(value) {
  return path.basename(String(value ?? "").replaceAll("\\", "/"));
}

function unique(values) {
  return [...new Set(values.filter(Boolean))];
}

function isPolicyDuty(text) {
  const value = clean(text);
  return /^(贯彻执行|参与拟订|参与编制|提出意见|协助起草|督促落实|推动落实)/u.test(value)
    && /(方针|政策|法律法规|法规|规章制度|制度|标准|规范|发展规划|办法)/u.test(value)
    && !/(审批|许可|备案|调查|填报|报送|审核|验收|监测|处置|整改|归档|系统)/u.test(value);
}

function addMatterHints(matter, terms) {
  const text = clean(matter);
  const hints = [
    [/交通流量|交通情况|交调/u, ["交通流量", "交通情况调查", "调查数据", "定期上报", "审批归档"]],
    [/项目备案|备案/u, ["养护工程备案", "项目备案", "项目上报", "备案库"]],
    [/项目计划|计划/u, ["年度计划", "计划编制", "计划下达", "项目库"]],
    [/风险隐患|安全生产/u, ["风险隐患", "隐患排查", "安全生产", "整改", "销号", "风险源"]],
    [/应急|保通/u, ["应急处置", "事件报送", "应急响应", "处置效果"]],
    [/技术状况|路况/u, ["技术状况", "路况评定", "基础设施技术状况"]],
    [/统计|报表/u, ["统计报表", "汇总", "统计分析", "报告"]],
    [/绩效考核|考核/u, ["绩效考核", "考核", "评分", "评价"]],
  ];
  for (const [pattern, extra] of hints) if (pattern.test(text)) for (const item of extra) terms.add(item);
}

function termsFor(matter, responsibility) {
  const normalized = clean(`${matter} ${responsibility}`)
    .replace(/^(承担指导|承担编制|承担|参与国省干线公路|参与|督促落实|贯彻执行|承办|负责|组织)/u, "")
    .replace(/事务性工作/g, " ")
    .replace(/[，,。；;、（）()“”《》]/g, " ");
  const terms = new Set();
  for (const token of normalized.split(/\s+/u)) if (token.length >= 3) terms.add(token);
  const compact = normalized.replace(/\s+/g, "");
  for (let length = Math.min(12, compact.length); length >= 4; length -= 1) {
    for (let start = 0; start <= compact.length - length; start += 2) terms.add(compact.slice(start, start + length));
  }
  addMatterHints(matter, terms);
  const priorityTerms = ["交通流量", "交通情况调查", "调查数据", "定期上报", "审批归档", "养护工程备案", "项目备案", "项目上报", "备案库", "年度计划", "计划编制", "计划下达", "风险隐患", "隐患排查", "安全生产", "整改", "销号"].filter((term) => terms.has(term));
  return unique([...priorityTerms, ...[...terms].sort((a, b) => b.length - a.length)]).slice(0, 36);
}

function isHeading(line) {
  return /^(第?[一二三四五六七八九十]+[、.]|第[一二三四五六七八九十0-9]+章|[（(][一二三四五六七八九十0-9]+[)）]|\d+(?:\.\d+){0,4}[、.．]|[一二三四五六七八九十]+、)/u.test(clean(line));
}

function nearestHeading(lines, lineIndex) {
  for (let index = lineIndex; index >= Math.max(0, lineIndex - 180); index -= 1) {
    const line = clean(lines[index]);
    if (line && isHeading(line)) return line.slice(0, 120);
  }
  return "";
}

function expandSection(lines, lineIndex) {
  let start = Math.max(0, lineIndex - 30);
  for (let index = lineIndex; index >= Math.max(0, lineIndex - 180); index -= 1) {
    if (isHeading(lines[index])) {
      start = index;
      break;
    }
  }
  let end = Math.min(lines.length, lineIndex + 90);
  for (let index = lineIndex + 1; index < Math.min(lines.length, lineIndex + 260); index += 1) {
    if (isHeading(lines[index])) {
      end = index;
      break;
    }
  }
  return {
    start_line: start + 1,
    end_line: end,
    heading: nearestHeading(lines, lineIndex),
    excerpt: lines.slice(start, end).join("\n").slice(0, maxExcerptChars),
  };
}

function scoreWindow(window, matched) {
  const actionScore = (window.match(/填报|录入|提交|上报|审核|审批|退回|补正|保存|归档|办结|查询|统计|汇总|生成|下达|反馈|整改|销号/gu) ?? []).length * 12;
  const sourceScore = (window.match(/系统|模块|功能|流程|业务|用户|部门|单位|数据/gu) ?? []).length * 4;
  return matched.reduce((sum, term) => sum + term.length ** 2, 0) + actionScore + sourceScore;
}

function candidatesFor(source, terms) {
  const candidates = [];
  for (let index = 0; index < source.lines.length; index += 1) {
    const windowStart = Math.max(0, index - 8);
    const windowEnd = Math.min(source.lines.length, index + 12);
    const window = source.lines.slice(windowStart, windowEnd).join("\n");
    const matched = terms.filter((term) => window.includes(term));
    if (!matched.length) continue;
    const section = expandSection(source.lines, index);
    candidates.push({
      score: scoreWindow(window, unique(matched)),
      matched_terms: unique(matched).slice(0, 16),
      line: index + 1,
      section_start_line: section.start_line,
      section_end_line: section.end_line,
      section_heading: section.heading,
      excerpt: section.excerpt,
    });
  }
  const seen = new Set();
  return candidates
    .sort((a, b) => b.score - a.score)
    .filter((candidate) => {
      const key = `${candidate.section_start_line}-${candidate.section_end_line}`;
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
}

async function readMatters(filePath) {
  const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(filePath));
  const values = workbook.worksheets.getItemAt(0).getUsedRange().values;
  const headers = values[0].map(clean);
  const matterIndex = headers.indexOf("事项");
  const departmentIndex = headers.indexOf("处室名称");
  const responsibilityIndex = headers.indexOf("职责");
  const typeIndex = headers.indexOf("职能类型");
  if (matterIndex === -1 || departmentIndex === -1 || responsibilityIndex === -1 || typeIndex === -1) {
    throw new Error("Stage 1 Excel must contain 处室名称、职责、事项、职能类型 headers.");
  }
  let department = "";
  let responsibility = "";
  const rows = [];
  for (const row of values.slice(1)) {
    department = clean(row[departmentIndex]) || department;
    responsibility = clean(row[responsibilityIndex]) || responsibility;
    const matter = clean(row[matterIndex]);
    if (!matter) continue;
    if (matterFilter && !matterFilter.test(`${matter} ${responsibility}`)) continue;
    rows.push({ department, responsibility, matter, functional_type: clean(row[typeIndex]) });
  }
  return rows;
}

async function loadSources() {
  const index = JSON.parse(await fs.readFile(indexPath, "utf8"));
  const root = path.dirname(path.dirname(indexPath));
  const sources = [];
  for (const item of index) {
    if (!item.text_file || item.extract_status !== "ready") continue;
    if (systemCodes.size && !systemCodes.has(String(item.system_code))) continue;
    try {
      const text = await fs.readFile(path.resolve(root, item.text_file), "utf8");
      sources.push({
        system_code: String(item.system_code),
        system_name: item.system_name,
        source_file: item.source_file,
        source_file_name: basename(item.source_file),
        text_file: item.text_file,
        lines: text.split(/\r?\n/),
      });
    } catch {
      // Unreadable extracted text cannot become evidence.
    }
  }
  return sources;
}

const matters = await readMatters(inputPath);
const sources = await loadSources();
if (!sources.length && !matters.every((item) => isPolicyDuty(`${item.matter} ${item.responsibility}`))) {
  throw new Error("No readable system materials remain after applying the requested system scope.");
}

const result = matters.map((record, matter_index) => {
  const policyDuty = isPolicyDuty(`${record.matter} ${record.responsibility}`);
  const terms = termsFor(record.matter, record.responsibility);
  const candidates = policyDuty ? [] : sources.flatMap((source) => candidatesFor(source, terms).map((candidate) => ({
    system_code: source.system_code,
    system_name: source.system_name,
    source_file: source.source_file,
    source_file_name: source.source_file_name,
    text_file: source.text_file,
    ...candidate,
  }))).sort((a, b) => b.score - a.score).slice(0, maxCandidates);
  return {
    matter_index,
    ...record,
    is_policy_duty: policyDuty,
    system_scope: [...systemCodes],
    query_terms: terms,
    candidates,
  };
});

await fs.mkdir(path.dirname(outputPath), { recursive: true });
await fs.writeFile(outputPath, JSON.stringify({
  schema: "stage2_context_package/v1",
  input: inputPath,
  index: indexPath,
  generated_at: new Date().toISOString(),
  matter_count: result.length,
  source_count: sources.length,
  items: result,
}, null, 2), "utf8");
console.log(JSON.stringify({ matterCount: result.length, sourceCount: sources.length, scopedSystemCodes: [...systemCodes], outputPath }, null, 2));
