import fs from "node:fs/promises";
import path from "node:path";

function optionalArgument(name, fallback = "") {
  const index = process.argv.indexOf(name);
  return index === -1 ? fallback : process.argv[index + 1] ?? fallback;
}

function argument(name) {
  const value = optionalArgument(name);
  if (!value) throw new Error(`Missing ${name}`);
  return value;
}

const contextPath = argument("--context");
const outputPath = argument("--output");
const requestedMode = optionalArgument("--mode", "auto");
const fixturePath = optionalArgument("--fixture-path");
const model = optionalArgument("--model", process.env.OPENAI_MODEL || "gpt-4.1-mini");
const apiKeyEnv = optionalArgument("--api-key-env", "OPENAI_API_KEY");
const apiKey = process.env[apiKeyEnv] || "";
const apiBase = optionalArgument("--api-base", process.env.OPENAI_BASE_URL || "https://api.openai.com/v1");
const apiUrl = optionalArgument("--api-url", `${apiBase.replace(/\/$/u, "")}/chat/completions`);

const allowedStatuses = new Set(["有", "待确认", "无", "不纳入本次梳理范畴"]);
const genericNodeTitles = /【(?:系统定位|资料范围核验|范围核验|边界说明|资料核查|系统模块|业务范围补充)】/u;
const actionWords = ["审核", "退回", "补正", "上报", "归档", "办结", "审批"];

function clean(value) {
  return String(value ?? "").replace(/\r/g, "").replace(/[ \t]+/g, " ").trim();
}

function normalizeLines(value) {
  return clean(value).replace(/\s*(?=节点\s*\d+【)/gu, "\n").replace(/\n{2,}/g, "\n").trim();
}

function node(no, title, body) {
  return `节点${no}【${title}】：${body}`;
}

function sentencePieces(text) {
  return clean(text).split(/(?<=[。；;])/u).map(clean).filter(Boolean);
}

function summarizeCandidate(candidate, matter) {
  const pieces = sentencePieces(candidate.excerpt);
  const preferred = pieces.filter((piece) => /(填报|录入|提交|上报|审核|审批|退回|补正|保存|归档|办结|查询|统计|汇总|生成|下达|反馈|整改|销号|调查|备案|监测|风险|隐患|安全)/u.test(piece));
  const source = (preferred.length ? preferred : pieces).slice(0, 3).join("；") || `资料片段与“${matter}”存在文本关联。`;
  return source.replace(/\s+/g, " ").slice(0, 260);
}

function titleForCandidate(candidate, matter) {
  const text = `${candidate.section_heading} ${candidate.excerpt}`;
  const rules = [
    [/调查数据|交通量|交调/u, "调查数据采集维护"],
    [/校验.*保存|保存.*校验/u, "数据校验保存"],
    [/定期上报|上报审批|审批归档|归档/u, "数据上报审批归档"],
    [/汇总|统计|报表/u, "数据汇总统计"],
    [/备案库|项目备案|备案/u, "项目备案填报"],
    [/年度计划|计划编制|计划下达/u, "项目计划编制下达"],
    [/风险隐患|隐患排查|整改|销号/u, "风险隐患排查整改"],
    [/安全生产|养护作业安全/u, "安全生产检查整改"],
    [/事件报送|应急处置|应急响应/u, "事件报送应急处置"],
    [/技术状况|路况评定/u, "技术状况数据评定"],
    [/查询|展示|看板/u, "数据查询展示"],
  ];
  for (const [pattern, title] of rules) if (pattern.test(text)) return title;
  const heading = clean(candidate.section_heading).replace(/^\d+(?:\.\d+)*[、.．]?\s*/u, "").slice(0, 16);
  return heading || clean(matter).slice(0, 14) || "业务资料片段";
}

function candidateRef(candidate, summary) {
  const heading = clean(candidate.section_heading) ? `“${clean(candidate.section_heading)}”` : "相关段落";
  return `依据《${candidate.source_file_name}》${heading}（提取文本第${candidate.section_start_line}-${candidate.section_end_line}行），${summary}`;
}

function systemList(candidates) {
  const values = candidates.map((item) => `${item.system_name}（${item.system_code}）`);
  return [...new Set(values)].join("；");
}

function isDirectEvidence(candidate) {
  const text = `${candidate.section_heading}\n${candidate.excerpt}`;
  const actionScore = (text.match(/填报|录入|提交|上报|审核|审批|退回|补正|保存|归档|办结|查询|统计|汇总|生成|下达|反馈|整改|销号/gu) ?? []).length;
  const tableOnly = /数据库设计|数据表|字段|主键|外键/u.test(text) && actionScore < 2;
  return actionScore >= 2 && !tableOnly;
}

function riskOrSafetyNeedsConfirmation(item, candidates) {
  if (!/风险隐患|安全生产/u.test(item.matter)) return false;
  return !candidates.some((candidate) => {
    const text = candidate.excerpt;
    return /(隐患排查|安全生产|检查)/u.test(text) && /(责任单位|整改责任|复查|销号|闭环|检查记录|发现主体)/u.test(text);
  });
}

function policyResult(item) {
  return {
    matter: item.matter,
    digital_status: "无",
    status: "无",
    process: "该事项为政策制度类职责，本轮按业务规则判定为无数字化，不展开流程。",
    systems: "",
    basis: [
      node(1, "事项来源职责", `依据阶段1职责原文“${item.responsibility}”，本事项为“${item.matter}”。`),
      node(2, "政策制度类判断", "依据阶段2业务规则：贯彻执行、参与拟订/编制、提出意见、协助起草、督促落实政策制度且未对应具体业务事项办理流程的事项，固定判定为无数字化，H列留空，不进入阶段3。"),
    ].join("\n"),
    quality_notes: ["policy_duty_fixed_no", "skipped_retrieval_by_rule"],
  };
}

function pendingResult(item, reason, candidates = []) {
  const basis = [
    node(1, "事项来源职责", `依据阶段1职责原文“${item.responsibility}”，本事项为“${item.matter}”。`),
  ];
  candidates.slice(0, 3).forEach((candidate, index) => {
    basis.push(node(index + 2, titleForCandidate(candidate, item.matter), candidateRef(candidate, summarizeCandidate(candidate, item.matter))));
  });
  basis.push(node(basis.length + 1, "缺失或待确认", reason));
  const systems = systemList(candidates.filter(isDirectEvidence).slice(0, 3));
  return {
    matter: item.matter,
    digital_status: "待确认",
    status: "待确认",
    process: `流程线索：${candidates.length ? "现有资料仅能确认部分相关系统功能或数据线索，尚不能串联为完整办理流程。" : "本轮候选资料未形成可回溯的系统办理流程证据。"}待确认：${reason}`,
    systems,
    basis: basis.join("\n"),
    quality_notes: ["dry_run_pending", candidates.length ? "partial_context_found" : "no_context_found"],
  };
}

function dryRunAnalyze(item) {
  if (item.is_policy_duty) return policyResult(item);
  if (/领导交办|其他事项/u.test(item.matter)) return {
    matter: item.matter,
    digital_status: "不纳入本次梳理范畴",
    status: "不纳入本次梳理范畴",
    process: "该事项为兜底类职责，未限定具体业务对象和办理流程，本轮不纳入阶段2数字化流程展开。",
    systems: "",
    basis: [node(1, "事项来源职责", `依据阶段1职责原文“${item.responsibility}”，本事项为“${item.matter}”。`), node(2, "兜底事项范围判断", "该事项未指向具体业务对象、材料、主体、状态或成果，不能因系统资料出现相近表述而拼接为数字化流程；需待人工明确具体交办事项后另行梳理。")].join("\\n"),
    quality_notes: ["catch_all_excluded"],
  };
  const candidates = (item.candidates ?? []).filter((candidate) => clean(candidate.excerpt)).slice(0, 6);
  if (!candidates.length) {
    return pendingResult(item, "需补充实际办理渠道、主责主体、申请或填报材料、处理动作、状态结果、退回补正和处室参与边界。");
  }
  if (riskOrSafetyNeedsConfirmation(item, candidates)) {
    return pendingResult(item, "现有资料未直接覆盖风险隐患或安全生产事项的发现主体、整改责任单位、复查销号、闭环反馈和本处室指导节点，不能把应急或监管功能直接拼成已确认流程。", candidates);
  }
  const direct = candidates.filter(isDirectEvidence);
  if (direct.length < 2) {
    return pendingResult(item, "现有资料不足以同时证明办理主体、材料或数据、操作顺序、状态转换和办结结果；需人工核验适用范围及处室实际权限。", candidates);
  }

  const selected = direct.slice(0, 5);
  const basis = [
    node(1, "事项来源职责", `依据阶段1职责原文“${item.responsibility}”，本事项为“${item.matter}”，需按该职责边界核验系统资料适用范围。`),
  ];
  selected.forEach((candidate, index) => {
    basis.push(node(index + 2, titleForCandidate(candidate, item.matter), candidateRef(candidate, summarizeCandidate(candidate, item.matter))));
  });
  basis.push(node(basis.length + 1, "缺失或待确认", "需人工核验本处室在该事项中的具体办理权限、退回补正对象、异常回流节点、最终成果签发或归档责任，以及不同公路范围之间是否存在口径差异。"));
  const summaries = selected.slice(0, 4).map((candidate) => summarizeCandidate(candidate, item.matter));
  return {
    matter: item.matter,
    digital_status: "有",
    status: "有",
    process: `流程：相关业务人员围绕“${item.matter}”通过${systemList(selected)}办理或支撑事项处理。现有资料可确认：${summaries.join("；")}。上述依据能够说明系统内存在材料或数据维护、处理、查询、统计、上报或归档等节点，可作为阶段3绘图的业务线索。待确认：本处室实际承担的发起、审核、反馈或成果确认边界，以及退回补正、异常处置、最终归档责任是否在现行制度中另有规定。`,
    systems: systemList(selected),
    basis: basis.join("\n"),
    quality_notes: ["dry_run_extract_only", "replace_with_model_or_human_review_before_final_delivery"],
  };
}

function buildPrompt(item) {
  const contextText = (item.candidates ?? []).map((candidate, index) => [
    `【候选${index + 1}】系统：${candidate.system_name}（${candidate.system_code}）`,
    `文件：${candidate.source_file}`,
    `章节/标题：${candidate.section_heading || "未识别"}`,
    `提取文本行号：${candidate.section_start_line}-${candidate.section_end_line}`,
    `命中词：${(candidate.matched_terms ?? []).join("、")}`,
    "原文：",
    candidate.excerpt,
  ].join("\n")).join("\n\n---\n\n");
  return [
    "你是“两清一图”阶段2数字化支撑与业务流程梳理执行器。请只基于给定事项、职责和候选资料片段，逐事项生成结构化 JSON，不得输出 Markdown。",
    "",
    "业务规则：先形成 I 列流程依据，再由 I 列归纳 F 列流程，最后填写 E/H。不得因为事项名称命中关键词就直接判“有”。制度政策类事项固定判“无”。数据表只能证明字段及关联，不能单独证明角色、审批和状态转换。保存成功、上报确认提示、审核通过分别表述。不同业务范围不得强行迁用。",
    "",
    "输出 JSON schema：",
    "{\"matter\":\"事项名称\",\"digital_status\":\"有|待确认|无|不纳入本次梳理范畴\",\"process\":\"F列：流程/待确认问题\",\"systems\":\"H列：系统（标号）\",\"basis\":\"I列：流程依据\",\"quality_notes\":[\"内部质量说明\"]}",
    "",
    `处室：${item.department}`,
    `职责：${item.responsibility}`,
    `事项：${item.matter}`,
    `职能类型：${item.functional_type}`,
    `是否命中制度政策类规则：${item.is_policy_duty ? "是" : "否"}`,
    "",
    "候选资料片段：",
    contextText || "无候选资料片段。",
  ].join("\n");
}

async function callModel(item) {
  if (!globalThis.fetch) throw new Error("Current Node runtime does not provide fetch; use --mode dry-run or a newer Node runtime.");
  const response = await fetch(apiUrl, {
    method: "POST",
    headers: { "Content-Type": "application/json", Authorization: `Bearer ${apiKey}` },
    body: JSON.stringify({
      model,
      temperature: 0.1,
      response_format: { type: "json_object" },
      messages: [
        { role: "system", content: "你严格执行两清一图阶段2规则，只返回单个 JSON 对象。" },
        { role: "user", content: buildPrompt(item) },
      ],
    }),
  });
  if (!response.ok) throw new Error(`Model request failed for ${item.matter}: ${response.status} ${await response.text()}`);
  const payload = await response.json();
  const content = payload.choices?.[0]?.message?.content;
  if (!content) throw new Error(`Model returned no content for ${item.matter}`);
  return JSON.parse(content);
}

function normalizeResult(item, raw) {
  const status = clean(raw.digital_status ?? raw.status);
  return {
    matter: item.matter,
    digital_status: status,
    status,
    process: clean(raw.process),
    systems: clean(raw.systems),
    basis: normalizeLines(raw.basis),
    quality_notes: Array.isArray(raw.quality_notes) ? raw.quality_notes.map(clean).filter(Boolean) : [],
  };
}

async function fixtureAnalyze(items) {
  if (!fixturePath) throw new Error("--fixture-path is required in fixture mode.");
  const fixture = JSON.parse(await fs.readFile(fixturePath, "utf8"));
  const fixtureItems = Array.isArray(fixture) ? fixture : fixture.items;
  if (!Array.isArray(fixtureItems)) throw new Error("Fixture must be an array or { items: [...] }.");
  const byMatter = new Map(fixtureItems.map((item) => [clean(item.matter), item]));
  return items.map((item) => {
    const matched = byMatter.get(clean(item.matter));
    if (!matched) throw new Error(`Fixture is missing matter: ${item.matter}`);
    return normalizeResult(item, matched);
  });
}

function validateResult(item, result) {
  if (clean(result.matter) !== clean(item.matter)) throw new Error(`Analysis matter mismatch: ${item.matter}`);
  if (!allowedStatuses.has(result.status)) throw new Error(`Invalid digital_status for ${item.matter}: ${result.status}`);
  if (!result.process || !result.basis) throw new Error(`Process or basis is empty for ${item.matter}`);
  if (genericNodeTitles.test(result.basis)) throw new Error(`Generic I-column node title remains for ${item.matter}`);
  if (item.is_policy_duty && (result.status !== "无" || result.systems)) throw new Error(`Policy duty must be 无 with empty systems: ${item.matter}`);
  if (result.status === "有") {
    if (!result.systems) throw new Error(`E=有 requires systems for ${item.matter}`);
    const nodeCount = (result.basis.match(/节点\s*\d+【/gu) ?? []).length;
    if (nodeCount < 3 || !result.basis.includes("《")) throw new Error(`E=有 requires multiple sourced nodes for ${item.matter}`);
  }
  const knownCodes = new Set((item.candidates ?? []).map((candidate) => String(candidate.system_code)));
  for (const match of result.systems.matchAll(/（([0-9][0-9+]*?)）/gu)) {
    if (!knownCodes.has(match[1])) throw new Error(`Unknown system code ${match[1]} in H column for ${item.matter}`);
  }
  for (const word of actionWords) {
    if (result.process.includes(word) && !result.basis.includes(word) && !result.process.includes("待确认")) {
      throw new Error(`Action ${word} appears in F without I support or pending note for ${item.matter}`);
    }
  }
}

const context = JSON.parse(await fs.readFile(contextPath, "utf8"));
const items = Array.isArray(context) ? context : context.items;
if (!Array.isArray(items)) throw new Error("Context package must be an array or { items: [...] }.");

let mode = requestedMode;
if (mode === "auto") mode = apiKey ? "model" : "dry-run";
if (!["model", "dry-run", "fixture"].includes(mode)) throw new Error("--mode must be auto, model, dry-run, or fixture.");
if (mode === "model" && !apiKey) throw new Error(`${apiKeyEnv} is required for --mode model.`);

let analysis;
if (mode === "fixture") {
  analysis = await fixtureAnalyze(items);
} else if (mode === "model") {
  analysis = [];
  for (const item of items) {
    const result = normalizeResult(item, await callModel(item));
    validateResult(item, result);
    analysis.push(result);
  }
} else {
  analysis = items.map((item) => normalizeResult(item, dryRunAnalyze(item)));
}

for (let index = 0; index < items.length; index += 1) validateResult(items[index], analysis[index]);

await fs.mkdir(path.dirname(outputPath), { recursive: true });
await fs.writeFile(outputPath, JSON.stringify({
  schema: "stage2_prompt_analysis/v1",
  mode,
  model: mode === "model" ? model : "",
  context: contextPath,
  generated_at: new Date().toISOString(),
  items: analysis,
}, null, 2), "utf8");
const statusCounts = analysis.reduce((acc, item) => {
  acc[item.status] = (acc[item.status] || 0) + 1;
  return acc;
}, {});
console.log(JSON.stringify({ mode, matterCount: analysis.length, statusCounts, outputPath }, null, 2));
