import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

console.warn("[legacy] build_stage2_rule_fusion_analysis.mjs is retained only for comparison. Formal stage 2 uses retrieve_stage2_context.mjs -> run_stage2_prompt_analysis.mjs -> run_stage2_digital_process.mjs.");

function optionalArgument(name, fallback = "") {
  const index = process.argv.indexOf(name);
  return index === -1 ? fallback : process.argv[index + 1] ?? fallback;
}
function argument(name) {
  const value = optionalArgument(name);
  if (!value) throw new Error(`Missing ${name}`);
  return value;
}

const stage1Path = optionalArgument("--stage1");
const seedStage2Path = optionalArgument("--seed-stage2");
const indexPath = optionalArgument("--index", "工作区/系统资料索引.json");
const outputPath = argument("--output");
const department = optionalArgument("--department", "养护技术服务处");

function clean(value) {
  return String(value ?? "").replace(/\r/g, "").replace(/[ \t]+/g, " ").trim();
}
function basename(value) {
  return path.basename(String(value ?? "").replaceAll("\\", "/"));
}
function node(no, title, body) {
  return `节点${no}【${title}】：${body}`;
}
function pendingProcess(reason, questions) {
  return `流程线索：${reason}待确认：${questions}`;
}
function pendingBasis(matter, responsibility, reason, questions) {
  return [
    node(1, "事项来源", `依据阶段1职责原文“${responsibility}”，可确认本事项为“${matter}”。`),
    node(2, "资料范围核验", `${reason}现有系统资料不能直接证明完整办理流程、办理主体和办结结果。`),
    node(3, "缺失或待确认", `待确认：${questions}`)
  ].join("\n");
}

async function readRows(inputPath) {
  const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(inputPath));
  const values = workbook.worksheets.getItemAt(0).getUsedRange().values;
  const rows = [];
  let currentDepartment = "";
  let currentResponsibility = "";
  for (const row of values.slice(1)) {
    currentDepartment = clean(row[0]) || currentDepartment;
    currentResponsibility = clean(row[1]) || currentResponsibility;
    const matter = clean(row[2]);
    if (!matter) continue;
    if (currentDepartment && currentDepartment !== department) throw new Error(`Unexpected department: ${currentDepartment}`);
    rows.push({ matter, responsibility: currentResponsibility, type: clean(row[3]) });
  }
  return rows;
}

async function loadSources() {
  const index = JSON.parse(await fs.readFile(indexPath, "utf8"));
  const root = path.dirname(path.dirname(indexPath));
  const sources = [];
  for (const item of index) {
    if (!item.text_file || item.extract_status !== "ready") continue;
    try {
      const textPath = path.resolve(root, item.text_file);
      const text = await fs.readFile(textPath, "utf8");
      sources.push({ ...item, textPath, text, lines: text.split(/\r?\n/) });
    } catch {
      // Unreadable text cannot become evidence.
    }
  }
  return sources;
}

function nearestHeading(lines, index) {
  for (let i = index; i >= Math.max(0, index - 160); i -= 1) {
    const line = clean(lines[i]);
    if (/^(\d+(?:\.\d+){1,4}|第[一二三四五六七八九十]+章|[（(][一二三四五六七八九十0-9]+[)）])/.test(line)) return line.slice(0, 80);
  }
  return "";
}
function findRef(sources, { codes = [], patterns = [] }) {
  const scoped = sources.filter((source) => !codes.length || codes.includes(String(source.system_code)));
  for (const source of scoped) {
    for (let i = 0; i < source.lines.length; i += 1) {
      const window = clean(source.lines.slice(i, i + 8).join(" "));
      if (window && patterns.every((pattern) => pattern.test(window))) {
        return {
          code: String(source.system_code),
          system: source.system_name,
          file: source.source_file,
          fileName: basename(source.source_file),
          line: i + 1,
          heading: nearestHeading(source.lines, i)
        };
      }
    }
  }
  return null;
}
function ref(found, summary) {
  if (!found) return `现有资料未明确，${summary}`;
  const heading = found.heading ? `“${found.heading}”` : "相关段落";
  return `依据《${found.fileName}》${heading}（提取文本第${found.line}行），${summary}`;
}
function systemList(refs, fallback = "") {
  const pairs = refs.filter(Boolean).map((item) => `${item.system}（${item.code}）`);
  return [...new Set(pairs)].join("；") || fallback;
}

function buildProjectFiling(record, sources) {
  const base = findRef(sources, { codes: ["4"], patterns: [/养护工程一体化业务流/, /五年规划项目库|年度建议项目库/, /养护工程备案库/] });
  const actors = findRef(sources, { codes: ["4"], patterns: [/市、县级单位|市级用户/, /养护工程备案库|养护工程管理/, /上报/] });
  const filing = findRef(sources, { codes: ["4"], patterns: [/养护工程备案库的主要功能/, /年度实施养护工程备案|项目上报/] });
  const select = findRef(sources, { codes: ["4"], patterns: [/市、县级单位/, /年度建议计划项目库/, /列入该年度的养护工程备案库/] });
  const query = findRef(sources, { codes: ["4"], patterns: [/养护工程备案库支持组合条件查询/, /计划类型/] });
  const management = findRef(sources, { codes: ["4"], patterns: [/养护工程管理的主要功能/, /项目开工/, /交.*工验收|交.*竣.*工验收|完工信息/] });
  const refs = [base, actors, filing, select, query, management];
  return {
    status: "有",
    systems: systemList(refs, "江西省公路养护综合监管系统（4）"),
    process: "流程：市、县级单位根据年度建议计划项目库筛选养护工程项目，补充实施年份等备案信息，将项目列入年度养护工程备案库；对建议计划库外项目，可结合实际填报项目基本信息，并通过计划类型等字段区分计划内、计划外。市、县级单位确认备案项目后通过系统上报，省级单位及省交通厅用户查阅已上报项目；后续已上报项目进入养护工程管理，由市、县级单位办理开工、进度、完工和竣工验收资料填报，形成工程过程和归档数据。待确认：普通国省干线实际填报单位、上报接收或审核权限、退回补正规则及养护技术服务处具体事务性操作。",
    basis: [
      node(1, "事项来源", `依据阶段1职责原文“${record.responsibility}”，可确认本事项为国省干线公路养护工程项目备案事务性工作。`),
      node(2, "业务流范围", ref(base, "养护工程一体化业务流包含年度建议项目库、五年规划项目库、养护工程备案库、养护工程管理和养护评价等模块。")),
      node(3, "主体和数据责任", ref(actors, "市级用户及地方高速公路管理部门负责填报或导入规划、备案和工程进度数据并上报，省级和省厅用户负责查阅监管。")),
      node(4, "备案项目填报", ref(select, "市、县级单位可筛选年度建议计划项目库项目，添加实施年份后列入年度养护工程备案库。")),
      node(5, "备案查询和计划类型", ref(query, "备案库可按行政区划、路线行政等级、路线编码、实施年份、计划类型等组合查询并导出。")),
      node(6, "后续工程管理", ref(management, "已上报备案项目衔接工程管理，市、县级单位办理开工、进度上报、完工和交竣工资料填报。")),
      node(7, "缺失或待确认", "现有资料未直接说明养护技术服务处在备案中的审批、退回或具体办理权限，退回补正对象和回流节点需人工核验。")
    ].join("\n")
  };
}

function buildProjectPlan(record, sources) {
  const planSystem = findRef(sources, { codes: ["4"], patterns: [/江西省普通公路计划项目管理系统/, /计划项目信息数据/, /投资计划/] });
  const forward = findRef(sources, { codes: ["4"], patterns: [/县级与市级用户/, /前期工作阶段/, /省级审核/] });
  const annual = findRef(sources, { codes: ["4"], patterns: [/计划管理主要功能|计划管理.*年度计划编制/, /年度计划编制/, /审查|审批/] });
  const handoff = findRef(sources, { codes: ["4"], patterns: [/省级用户完成计划下达/, /自动进入工程管理模块/] });
  const business = findRef(sources, { codes: ["4"], patterns: [/普通公路建设养护业务计划内外工程项目/, /数据规范化/] });
  const refs = [planSystem, forward, annual, handoff, business];
  return {
    status: "有",
    systems: systemList(refs, "江西省公路养护综合监管系统（4）"),
    process: "流程：县级、市级用户填报本年度进入前期工作阶段的项目立项信息、施工图信息，提交省级用户审核；省级用户审核后，在五年项目库中标识已开展前期的工程项目。年度计划环节在前期项目和项目库基础上开展普通公路项目年度计划编制、审查、审批和上报，并可根据实际情况办理计划变更。省级用户完成计划下达后，项目自动进入工程管理模块，后续办理开工、月报进度、完工和竣工验收信息。待确认：养护技术服务处在本事项中的实际收件、核验、报送和成果交接，以及不通过时的退回对象和回流节点。",
    basis: [
      node(1, "事项来源", `依据阶段1职责原文“${record.responsibility}”，可确认本事项为国省干线公路养护工程项目计划事务性工作。`),
      node(2, "系统定位", ref(planSystem, "普通公路计划项目管理系统实现计划项目信息追溯查阅，并为安排省级普通公路投资计划提供信息支撑。")),
      node(3, "前期项目填报与审核", ref(forward, "县级与市级用户填报进入前期工作阶段的立项和施工图信息，系统提供省级审核信息功能，并在五年项目库中标识已开展前期的工程项目。")),
      node(4, "年度计划办理", ref(annual, "计划管理实现普通公路项目年度计划编制、审查、审批和按期上报，并支持计划变更。")),
      node(5, "计划下达与工程衔接", ref(handoff, "省级用户完成计划下达后，项目自动进入工程管理模块，后续管理开工、月报进度、完工和竣工验收信息。")),
      node(6, "业务范围补充", ref(business, "系统建设目标包含普通公路建设养护计划内外工程项目数据规范化、精细化、系统化管理。")),
      node(7, "缺失或待确认", "现有资料未直接说明养护技术服务处的收件、审查、退回或下达权限；不通过时退回对象和回流节点需人工核验。")
    ].join("\n")
  };
}

function buildMaintenance(record, sources) {
  const scope = findRef(sources, { codes: ["4"], patterns: [/日常养护/, /养护工程|桥隧管理|路况评定/] });
  const data = findRef(sources, { codes: ["4"], patterns: [/数据工程建设/, /业务数据库|主题数据库|共享数据库/] });
  const dashboard = findRef(sources, { codes: ["4"], patterns: [/统计报表专题|养护专题看板|GIS可视化/, /监管|分析|展示/] });
  const refs = [scope, data, dashboard];
  return {
    status: "有",
    systems: systemList(refs, "江西省公路养护综合监管系统（4）"),
    process: "流程：省、市、县及相关养护管理单位围绕普通国省干线日常养护、养护工程、桥隧管理、路况评定等业务采集和维护养护业务数据；系统对业务数据进行集中管理、实时更新、交换共享和专题展示，形成统计报表、养护专题看板、GIS可视化或监管分析结果。业务管理单位据此开展查询、监管、分析和指导支撑。待确认：养护技术服务处实际指导对象、材料来源、反馈渠道、问题跟踪方式和指导结果留痕。",
    basis: [
      node(1, "事项来源", `依据阶段1职责原文“${record.responsibility}”，可确认本事项为国省干线公路养护管理指导事务性工作。`),
      node(2, "业务覆盖范围", ref(scope, "系统资料覆盖日常养护、养护工程、桥隧管理、路况评定等养护业务域。")),
      node(3, "数据集中管理", ref(data, "数据工程形成基础数据库、业务数据库、主题数据库和共享数据库，支撑养护业务数据集中管理与交换共享。")),
      node(4, "监管和展示支撑", ref(dashboard, "系统通过专题看板、统计报表、GIS可视化或工作面板提供查询、监管、分析和展示能力。")),
      node(5, "缺失或待确认", "现有资料能证明数据和监管支撑，未直接规定养护技术服务处发起指导、下发意见、反馈闭环或系统权限。")
    ].join("\n")
  };
}

function buildRoadMonitoring(record, sources) {
  const subject = findRef(sources, { codes: ["8"], patterns: [/路网运行监测管理系统主要面向路网运行监测业务人员/] });
  const content = findRef(sources, { codes: ["8"], patterns: [/视频|交通运行状况/, /公路事件|基础设施技术状况/, /预测预警/] });
  const exchange = findRef(sources, { codes: ["8"], patterns: [/数据共享交换/, /流量数据|事件数据|基础设施技术状况数据/] });
  const refs = [subject, content, exchange];
  return {
    status: "有",
    systems: systemList(refs, "江西省普通干线路网运行监测与应急处置平台（二期）（8）"),
    process: "流程：路网运行监测业务人员通过平台整合公路管理部门、社会机构等多源路网运行监测信息，汇聚视频、交通运行、交通情况调查、公路事件、基础设施技术状况、气象环境和设备运行等数据；平台开展实时监测、查询展示和预测预警，并通过数据交换形成路网运行监测数据和报告支撑。业务管理单位基于监测、预警和报告开展研判指导。待确认：预警触发条件、指导对象、闭环反馈和养护技术服务处具体操作权限。",
    basis: [
      node(1, "事项来源", `依据阶段1职责原文“${record.responsibility}”，可确认本事项为国省干线公路路网运行监测指导事务性工作。`),
      node(2, "监测主体", ref(subject, "路网运行监测管理系统主要面向路网运行监测业务人员，并整合多源路网运行监测信息资源。")),
      node(3, "监测内容", ref(content, "系统对视频、交通运行状况、公路事件、基础设施技术状况、气象环境和设备运行状况等进行实时监测和预测预警。")),
      node(4, "数据交换", ref(exchange, "系统通过数据共享交换获取视频图像、流量、拥堵、事件、基础设施技术状况和气象环境等数据。")),
      node(5, "缺失或待确认", "现有资料未指定养护技术服务处的收件、研判、下发或反馈权限。")
    ].join("\n")
  };
}

function buildTechnicalCondition(record, sources) {
  const maintenance = findRef(sources, { codes: ["4"], patterns: [/路况评定|技术状况/, /日常养护|桥隧管理/] });
  const monitor = findRef(sources, { codes: ["8"], patterns: [/基础设施技术状况/, /实时监测|预测预警/] });
  const exchange = findRef(sources, { codes: ["8"], patterns: [/基础设施技术状况数据/, /数据共享交换|获取/] });
  const refs = [maintenance, monitor, exchange];
  return {
    status: "有",
    systems: systemList(refs, "江西省公路养护综合监管系统（4）；江西省普通干线路网运行监测与应急处置平台（二期）（8）"),
    process: "流程：养护业务人员围绕路况评定、桥隧管理和基础设施技术状况等业务维护或接入技术状况数据；路网运行监测平台汇聚基础设施技术状况等多源数据，开展监测、展示、分析或预警；养护管理单位使用技术状况看板、专题报告、监测预警或统计结果支撑养护管理和路网运行研判。待确认：检测采集主体、数据质量校核、补正机制、正式报告签发主体和养护技术服务处参与环节。",
    basis: [
      node(1, "事项来源", `依据阶段1职责原文“${record.responsibility}”，可确认本事项为国省干线公路技术状况监测事务性工作。`),
      node(2, "养护业务范围", ref(maintenance, "养护系统资料覆盖路况评定、桥隧管理等与技术状况相关的业务数据。")),
      node(3, "监测数据", ref(monitor, "路网运行监测系统将基础设施技术状况纳入实时监测和预测预警范围。")),
      node(4, "数据来源", ref(exchange, "平台可通过数据共享交换获取基础设施技术状况数据等多源数据。")),
      node(5, "缺失或待确认", "现有资料未明确检测主体、质量责任、补正机制、报告形成主体及本处室事务性办理权限。")
    ].join("\n")
  };
}

function buildTrafficSurvey(record, sources) {
  const input = findRef(sources, { codes: ["5"], patterns: [/调查数据管理/, /人工调查数据|设备数据|导入/] });
  const validate = findRef(sources, { codes: ["5"], patterns: [/数据校验通过后.*保存|校验通过后.*保存/] });
  const summary = findRef(sources, { codes: ["5"], patterns: [/小时|月|季度|年度/, /汇总生成|周期数据/] });
  const approval = findRef(sources, { codes: ["5"], patterns: [/定期上报审批过程/, /校验上报|审批归档/] });
  const reject = findRef(sources, { codes: ["5"], patterns: [/上级部门/, /审批后/, /退回处理/] });
  const refs = [input, validate, summary, approval, reject];
  return {
    status: "有",
    systems: systemList(refs, "江西省公路交通情况调查系统（5）"),
    process: "流程：交调业务人员维护调查站点、设备等基础信息，录入、编辑或导入连续站点、间隙站点、比重站点等交通量数据；系统对人工调查数据、设备数据和导入数据进行校验，校验通过后保存并汇总形成小时、日、月、季度、半年、年度等周期数据。下级部门完成月度、季度或年度数据填报后，经校验上报至上级部门；上级部门审批，不满足要求的数据退回处理，满足要求的数据归档处理，上报后相应数据不能再修改，归档后统计报表使用历史数据。待确认：调查站点分工、现行系统版本衔接、养护技术服务处参与节点和退回后的具体重报路径。",
    basis: [
      node(1, "事项来源", `依据阶段1职责原文“${record.responsibility}”，可确认本事项为国省干线公路交通流量调查事务性工作。`),
      node(2, "调查数据填报", ref(input, "系统支持人工调查数据录入填报、设备数据审核修订、异常数据处理和设备数据导入。")),
      node(3, "数据校验保存", ref(validate, "交通量数据校验通过后保存并重新加载显示，已上报数据不能修改或清除。")),
      node(4, "汇总统计", ref(summary, "系统基于基础数据汇总生成小时、日、月、季度、半年和年度等周期数据。")),
      node(5, "报送审批归档", ref(approval, "定期上报审批包括月度、季度、年度报送，主体过程为校验上报和审批归档。")),
      node(6, "退回处理", ref(reject, "下级部门填报后上报至上级部门，上级部门审批后对不满足要求的数据退回处理，对满足要求的数据归档处理。")),
      node(7, "缺失或待确认", "资料未点名养护技术服务处在交调系统中的录入、审核、报送或退回后重报权限。")
    ].join("\n")
  };
}

function buildStatistics(record, sources) {
  const source = findRef(sources, { codes: ["4"], patterns: [/日常养护|养护工程|桥隧管理|路况评定/, /统计报表|专题/] });
  const data = findRef(sources, { codes: ["4"], patterns: [/数据工程建设/, /集中管理|交换共享/] });
  const report = findRef(sources, { codes: ["4"], patterns: [/多维统计分析|统计报表|评价分析报告/] });
  const refs = [source, data, report];
  return {
    status: "有",
    systems: systemList(refs, "江西省公路养护综合监管系统（4）"),
    process: "流程：省、市、县及相关养护管理单位围绕日常养护、养护工程、桥隧管理、路况评定等业务持续形成养护业务数据；系统通过数据工程实现养护业务数据集中管理、实时更新和交换共享，并通过统计报表专题、养护专题看板、养护工程专题报告等方式开展查询、汇总、统计和展示。业务人员基于统计结果开展监管分析、报表编制或材料支撑。待确认：数据来源责任、统计口径、审核报送主体、对外报送规则及养护技术服务处具体事务性工作。",
    basis: [
      node(1, "事项来源", `依据阶段1职责原文“${record.responsibility}”，可确认本事项为国省干线公路养护相关统计事务性工作。`),
      node(2, "业务数据来源", ref(source, "养护业务覆盖日常养护、养护工程、桥隧管理、路况评定等业务域，并包含统计报表或专题应用。")),
      node(3, "数据集中管理", ref(data, "数据工程支撑养护业务数据集中管理、实时更新和交换共享。")),
      node(4, "统计结果", ref(report, "系统支持多维统计分析、统计报表、评价分析报告等结果输出。")),
      node(5, "缺失或待确认", "资料未直接规定统计审核、对外报送、报部口径及养护技术服务处权限。")
    ].join("\n")
  };
}

function buildEmergency(record, sources) {
  const monitor = findRef(sources, { codes: ["8"], patterns: [/路网运行监测管理系统主要面向/, /公路事件|气象环境/] });
  const report = findRef(sources, { codes: ["8"], patterns: [/事件报送管理/, /转发报送|发送事件信息/] });
  const response = findRef(sources, { codes: ["8"], patterns: [/应急预案模板匹配|处置预案/, /风险源|应急响应/] });
  const evaluation = findRef(sources, { codes: ["8"], patterns: [/处置效果评估/, /灾害损失评估|突发事件统计分析/] });
  const refs = [monitor, report, response, evaluation];
  return {
    status: "有",
    systems: systemList(refs, "江西省普通干线路网运行监测与应急处置平台（二期）（8）"),
    process: "流程：路网运行监测平台汇聚视频、交通运行、交通事件、基础设施技术状况、气象环境和设备运行等信息，发现或接入公路运行事件后，值守管理子系统开展事件报送和事件转接；应急处置系统根据事件突发类型匹配应急预案、关联风险源并开展应急调度，处置过程形成事件信息、处置指令、现场图片视频文字信息和处置效果评估等记录。业务管理单位据此开展应急保通研判、调度和结果评估。待确认：事件分级、指挥主体、保通责任单位、线下处置与系统记录衔接，以及养护技术服务处具体指导方式。",
    basis: [
      node(1, "事项来源", `依据阶段1职责原文“${record.responsibility}”，可确认本事项为国省干线公路应急保通指导事务性工作。`),
      node(2, "运行监测基础", ref(monitor, "平台面向路网运行监测业务人员，整合视频、交通运行状况、公路事件、基础设施技术状况、气象环境等信息。")),
      node(3, "事件报送和转接", ref(report, "值守管理子系统支持事件报送和事件转接，实现事件信息的接入管理和转发报送。")),
      node(4, "应急处置", ref(response, "系统可根据事件突发类型匹配应急预案，关联风险源，并启动对应处置流程。")),
      node(5, "评估统计", ref(evaluation, "系统支持处置效果评估、灾害损失评估和突发事件统计分析。")),
      node(6, "缺失或待确认", "系统资料未直接规定养护技术服务处的指令下达、现场抢通、复盘或反馈权限。")
    ].join("\n")
  };
}

function classify(record) {
  const matter = record.matter;
  if (/工程项目备案/u.test(matter)) return "projectFiling";
  if (/工程项目计划/u.test(matter)) return "projectPlan";
  if (/养护管理事务性工作/u.test(matter) && !/发展纲要|项目|技术状况|统计/u.test(matter)) return "maintenance";
  if (/路网运行监测/u.test(matter)) return "roadMonitoring";
  if (/技术状况监测/u.test(matter)) return "technicalCondition";
  if (/交通流量调查/u.test(matter)) return "trafficSurvey";
  if (/养护相关统计/u.test(matter)) return "statistics";
  if (/应急保通/u.test(matter)) return "emergency";
  return "pending";
}

function analyze(record, sources) {
  const kind = classify(record);
  if (kind === "projectFiling") return buildProjectFiling(record, sources);
  if (kind === "projectPlan") return buildProjectPlan(record, sources);
  if (kind === "maintenance") return buildMaintenance(record, sources);
  if (kind === "roadMonitoring") return buildRoadMonitoring(record, sources);
  if (kind === "technicalCondition") return buildTechnicalCondition(record, sources);
  if (kind === "trafficSurvey") return buildTrafficSurvey(record, sources);
  if (kind === "statistics") return buildStatistics(record, sources);
  if (kind === "emergency") return buildEmergency(record, sources);

  const matter = record.matter;
  let reason = "本轮指定系统资料未检索到与该事项直接对应的办理功能、材料流转、审核规则和办结结果。";
  let questions = "实际办理渠道、主责主体、申请或填报材料、审核规则、退回补正、结果文书或系统记录。";
  if (/发展纲要/u.test(matter)) {
    reason = "系统资料可提供项目库、计划、工程进度和养护业务数据支撑，但未说明发展纲要的编制、论证、会商、审签和发布流程。";
    questions = "编制牵头主体、使用的数据材料、审核程序、成果文稿及养护技术服务处事务性工作。";
  } else if (/绩效考核/u.test(matter)) {
    reason = "系统资料可见养护考核或建养业务数据考核功能线索，但未形成养护工程项目绩效考核对象、指标、评分、确认和结果使用的完整流程。";
    questions = "考核对象、指标来源、主责单位、评分或申诉规则、结果反馈和本处室事务性工作。";
  } else if (/风险隐患排查整治/u.test(matter)) {
    reason = "路网运行与应急资料涉及事件、风险和预警能力，但未形成普通国省干线养护风险隐患排查、整改、复查、销号的同一业务流程。";
    questions = "隐患发现主体、整治责任单位、整改材料、复查销号和本处室指导节点。";
  } else if (/应急体系建设/u.test(matter)) {
    reason = "现有平台可支撑应急资源、应急指挥和事件信息处理，但未说明应急体系建设的规划、编制、评审、发布、演练或更新流程。";
    questions = "体系建设牵头主体、建设材料、审核机制、成果形式和本处室指导工作。";
  } else if (/安全生产/u.test(matter)) {
    reason = "系统资料涉及养护作业、应急和监管功能，但未检索到养护作业安全生产指导、检查、整改、复查或结果管理的直接业务流程。";
    questions = "安全检查对象、检查表单、整改时限、复查销号、主责单位和本处室指导节点。";
  } else if (/资质审批/u.test(matter)) {
    questions = "实际办理渠道、主责审批部门、申请材料、补正规则、结果文书和本处室事务性工作。";
  } else if (/造价/u.test(matter)) {
    reason = "养护工程系统可提供项目、计划、进度和工程资料等数据支撑，但未检索到养护造价编制、审核、调整或成果发布的直接流程。";
    questions = "造价数据来源、参与主体、审核权限、成果类型和本处室参与方式。";
  } else if (/市场管理/u.test(matter)) {
    questions = "实际关联系统、市场管理对象、处置规则、结果记录和本处室参与环节。";
  } else if (/信用体系/u.test(matter)) {
    questions = "信用信息采集来源、评价主体、异议处理、结果使用和本处室参与方式。";
  } else if (/领导交办/u.test(matter)) {
    reason = "该职责为兜底事项，未限定具体业务对象；现有资料无法判断是否属于本次系统范围或是否存在统一数字化办理过程。";
    questions = "实际交办事项范围、交办渠道、承办材料、反馈成果及是否纳入后续梳理。";
  }
  return {
    status: "待确认",
    systems: "",
    process: pendingProcess(reason, questions),
    basis: pendingBasis(matter, record.responsibility, reason, questions)
  };
}

const sourcePath = stage1Path || seedStage2Path;
if (!sourcePath) throw new Error("Provide --stage1 or --seed-stage2.");
const rows = await readRows(sourcePath);
const sources = await loadSources();
const analysis = rows.map((record) => ({
  matter: record.matter,
  ...analyze(record, sources),
  source_note: "阶段2证据驱动重跑：按事项类型收敛系统资料并由I列证据生成F/E/H。",
  responsibility: record.responsibility
}));

await fs.mkdir(path.dirname(outputPath), { recursive: true });
await fs.writeFile(outputPath, JSON.stringify(analysis, null, 2), "utf8");
const statusCounts = analysis.reduce((acc, item) => {
  acc[item.status] = (acc[item.status] || 0) + 1;
  return acc;
}, {});
console.log(JSON.stringify({ sourcePath, indexPath, outputPath, matterCount: analysis.length, statusCounts }, null, 2));
