import fs from "node:fs/promises";
import path from "node:path";
import sharp from "sharp";
import { FileBlob, SpreadsheetFile, Workbook } from "@oai/artifact-tool";

function argument(name) {
  const i = process.argv.indexOf(name);
  if (i === -1 || !process.argv[i + 1]) throw new Error(`Missing ${name}`);
  return process.argv[i + 1];
}
function optionalArgument(name, fallback = "") {
  const i = process.argv.indexOf(name);
  return i === -1 ? fallback : process.argv[i + 1] ?? fallback;
}

const inputPath = argument("--input");
const outputDir = argument("--output-dir");
const department = argument("--department");
const label = optionalArgument("--label");
const suffix = label ? `_${label}` : "";
const outputBase = `03_${department}${suffix}`;
const drawioPath = path.join(outputDir, `${outputBase}.drawio`);
const reviewedPath = path.join(outputDir, `${outputBase}_泳道图核验版.xlsx`);
const pendingPath = path.join(outputDir, `${outputBase}_泳道图待核验清单.xlsx`);
const imageDir = path.join(outputDir, `泳道图预览${suffix}`);
const headers = ["处室名称", "职责", "事项", "职能类型", "有没有数字化", "流程/待确认问题", "泳道图", "系统（标号）", "流程依据"];
const black = "#000000";
const gray = "#BFBFBF";
const lightGray = "#F2F2F2";
const font = "Microsoft YaHei";

function clean(value) {
  return String(value ?? "").replace(/\r/g, "").replace(/[ \t]+/g, " ").trim();
}
function xml(value) {
  return String(value).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}
function truncate(value, length = 42) {
  const text = clean(value);
  return text.length > length ? `${text.slice(0, length - 1)}…` : text;
}
function wrap(value, width) {
  const text = clean(value);
  if (!text) return [""];
  const chunks = [];
  for (let i = 0; i < text.length; i += width) chunks.push(text.slice(i, i + width));
  return chunks;
}
function drawioText(value, width = 14) {
  return xml(wrap(value, width).join("<br>"));
}
function cell(id, value, x, y, width, height, style, textWidth = 14) {
  return `<mxCell id="${id}" value="${drawioText(value, textWidth)}" style="${style}" vertex="1" parent="1"><mxGeometry x="${x}" y="${y}" width="${width}" height="${height}" as="geometry"/></mxCell>`;
}
function edge(id, source, target, points = [], value = "", dashed = false) {
  const style = `edgeStyle=orthogonalEdgeStyle;rounded=0;orthogonalLoop=1;jettySize=auto;html=1;endArrow=block;endFill=1;strokeColor=${black};fontFamily=${font};fontSize=11;${dashed ? "dashed=1;" : ""}`;
  const pointXml = points.length ? `<Array as="points">${points.map((p) => `<mxPoint x="${p.x}" y="${p.y}"/>`).join("")}</Array>` : "";
  return `<mxCell id="${id}" value="${xml(value)}" style="${style}" edge="1" parent="1" source="${source}" target="${target}"><mxGeometry relative="1" as="geometry">${pointXml}</mxGeometry></mxCell>`;
}

function parseNodes(basis) {
  return String(basis ?? "")
    .split(/\n+/)
    .map((paragraph) => {
      const match = clean(paragraph).match(/^节点\s*(\d+)【([^】]+)】：([\s\S]*)$/u);
      return match ? { number: Number(match[1]), label: match[2], detail: match[3] } : null;
    })
    .filter(Boolean)
    .sort((a, b) => a.number - b.number);
}

function systemsFrom(value) {
  const text = clean(value);
  if (!text) return [];
  return text.split(/[；;]/u).map((item) => clean(item.replace(/（\d+）$/u, ""))).filter(Boolean);
}

function actorFor(node, record) {
  const label = clean(node.label);
  const text = `${node.label} ${node.detail}`;
  if (/主体和数据责任|备案项目填报|备案查询|备案项目选取|计划内外标识|前期项目填报|后续工程管理/u.test(label)) return "养护单位/项目管理单位";
  if (/项目上报|项目库建立|系统定位|年度计划办理|计划下达/u.test(label)) return "厅公路管理处/上级养护管理部门";
  if (/业务流范围|业务覆盖范围|养护业务范围|业务数据来源/u.test(label)) return /统计/u.test(record.matter) ? "养护管理单位" : inferredPrimaryActor(record);
  if (/监管和展示支撑|监管与数据支撑|数据集中管理|统计结果|业务使用|结果和使用/u.test(label)) return inferredManagementActor(record);
  if (/监测主体|监测内容|监测数据|数据来源|数据交换|运行监测基础|事件报送/u.test(label)) return inferredPrimaryActor(record);
  if (/应急处置|评估统计|评估和统计|报送审批归档|报送审批与归档/u.test(label)) return inferredManagementActor(record);
  if (/调查数据填报|数据校验保存|数据校验与保存|汇总统计/u.test(label)) return "交调业务人员";
  if (/养护技术服务处|省综交中心/u.test(text)) return "养护技术服务处（综交中心）";
  if (/省交通厅|省级单位及省交通厅|省级用户|省级单位/u.test(text)) return "厅公路管理处/上级养护管理部门";
  if (/市、县级单位|市、县单位|市县级单位|县级与市级用户|县级、市级用户|市、县级用户/u.test(text)) return "养护单位/项目管理单位";
  if (/交调业务人员|交调工作业务人员/u.test(text)) return "交调业务人员";
  if (/路网运行监测业务人员|值守/u.test(text)) return "路网监测/值守人员";
  if (/应急处置人员/u.test(text)) return "应急处置人员/保通单位";
  if (/交通流量调查/u.test(record.matter) && /报送审批|归档/u.test(node.label)) return "交调管理/审批人员";
  if (/交通流量调查/u.test(record.matter)) return "交调业务人员";
  if (/技术状况监测/u.test(record.matter) && /结果|使用/u.test(node.label)) return "养护管理单位";
  if (/技术状况监测/u.test(record.matter)) return "养护业务人员";
  if (/业务管理单位/u.test(record.process)) return "业务管理单位";
  if (/养护管理单位|各级养护管理单位/u.test(record.process)) return "各级养护管理单位";
  return inferredPrimaryActor(record);
}

function inferredPrimaryActor(record) {
  if (/交通流量调查/u.test(record.matter)) return "交调业务人员";
  if (/技术状况监测/u.test(record.matter)) return "养护业务人员";
  if (/路网运行监测|应急保通/u.test(record.matter)) return "路网监测/值守人员";
  if (/统计|养护管理/u.test(record.matter)) return "养护管理单位";
  return "养护单位/项目管理单位";
}

function inferredManagementActor(record) {
  if (/交通流量调查/u.test(record.matter)) return "交调管理/审批人员";
  if (/路网运行监测|应急保通|统计|养护管理/u.test(record.matter)) return "业务管理单位";
  if (/技术状况监测/u.test(record.matter)) return "养护管理单位";
  return "厅公路管理处/上级养护管理部门";
}

function orderedActors(record, actors) {
  const merged = actors.filter(Boolean);
  const result = [...new Set(merged)].filter((actor) => !/(系统|平台|数据库|相关业务主体)/u.test(actor));
  return result.length ? result : [inferredPrimaryActor(record)];
}

function isDrawableNode(node) {
  return !/(事项来源|处室边界|边界|缺失|待确认|证据边界|资料范围核验|范围核验|系统模块|用户角色)/u.test(node.label);
}

function activityName(node) {
  const label = clean(node.label);
  const replacements = [
    [/业务流范围/u, "确认养护工程业务流"],
    [/主体和数据责任/u, "填报导入并上报数据"],
    [/备案项目填报/u, "筛选项目并填报备案"],
    [/备案查询和计划类型/u, "查询备案并标识计划类型"],
    [/备案项目选取与填报/u, "筛选项目并填报备案"],
    [/计划内外标识/u, "标识计划内外项目"],
    [/项目上报与查阅/u, "上报项目并查阅"],
    [/后续工程管理/u, "办理工程过程管理"],
    [/系统定位/u, "确认计划系统范围"],
    [/项目库建立与审查/u, "建立审查项目库"],
    [/前期项目填报/u, "填报前期项目信息"],
    [/年度计划办理/u, "编制审查年度计划"],
    [/计划下达与工程衔接/u, "下达计划并衔接工程"],
    [/业务范围补充/u, "补充业务范围"],
    [/业务覆盖范围/u, "采集维护养护业务数据"],
    [/监管和展示支撑/u, "查询监管养护数据"],
    [/监管与数据支撑/u, "查询监管养护数据"],
    [/数据集中管理/u, "集中管理共享数据"],
    [/监测主体/u, "汇聚路网监测数据"],
    [/监测主体和数据来源/u, "汇聚路网监测数据"],
    [/监测内容/u, "开展实时监测预警"],
    [/数据交换/u, "交换数据并提供结果"],
    [/数据交换与结果提供/u, "交换数据并提供结果"],
    [/养护业务范围/u, "维护技术状况相关数据"],
    [/监测数据/u, "监测基础设施技术状况"],
    [/路网监测数据/u, "监测基础设施技术状况"],
    [/数据来源/u, "接入技术状况数据"],
    [/技术状况数据来源/u, "汇聚技术状况数据"],
    [/调查数据填报/u, "录入或导入交调数据"],
    [/数据校验保存/u, "校验保存交通量数据"],
    [/数据校验与保存/u, "校验保存交通量数据"],
    [/汇总统计/u, "汇总周期交通量"],
    [/报送审批归档/u, "校验上报并归档"],
    [/报送审批与归档/u, "校验上报并归档"],
    [/业务数据来源/u, "形成养护业务数据"],
    [/统计结果/u, "生成统计报表和分析"],
    [/业务使用/u, "使用统计结果支撑监管"],
    [/运行监测基础/u, "汇聚运行和事件信息"],
    [/事件报送和转接/u, "报送转接事件信息"],
    [/应急处置/u, "匹配预案并应急调度"],
    [/评估统计/u, "评估统计处置结果"],
    [/评估和统计/u, "评估统计处置结果"],
    [/结果和使用/u, "形成监测分析结果"],
  ];
  for (const [pattern, replacement] of replacements) if (pattern.test(label)) return replacement;
  return truncate(label.replace(/^(办理或|流程|数字化|数据)/u, ""), 18);
}

function deriveResult(record, activities) {
  const text = `${record.process}\n${record.basis}`.replace(/待确认：[\s\S]*$/u, "");
  if (/技术状况监测/u.test(record.matter)) return "形成技术状况监测支撑";
  const exact = [
    "为省厅提供普通干线公路网运行监测数据",
    "形成工程过程和归档数据",
    "形成交通量数据及统计结果",
    "形成技术状况监测、预警、展示或专题分析结果",
    "形成养护业务数据集中管理和统计分析结果",
    "形成事件信息、处置指令和处置效果评估记录"
  ].find((item) => text.includes(item));
  if (exact) return truncate(exact, 34);
  const patterns = [
    /形成([^。；\n]{4,42}(?:数据|信息|结果|报告|报表|台账|记录|支撑))/u,
    /(项目自动进入工程管理模块[^。；\n]{0,32})/u,
    /(供[^。；\n]{4,42}(?:使用|参考|支撑))/u,
  ];
  for (const pattern of patterns) {
    const match = text.match(pattern);
    if (match) return truncate(match[0], 34);
  }
  return activities.length ? `完成${activities.at(-1).name}` : "";
}

function openQuestions(record) {
  const found = record.process.match(/待确认：([^。\n]+[。]?)/u);
  return found ? clean(found[1]) : "";
}

function buildFlow(record) {
  const rawActivities = record.nodes.filter(isDrawableNode).map((node) => ({
    id: "",
    name: activityName(node),
    actor: actorFor(node, record),
    source: `节点${node.number}【${node.label}】`
  }));
  const activities = rawActivities.filter((activity, index, all) => activity.name && all.findIndex((item) => item.name === activity.name && item.actor === activity.actor) === index);
  const actors = orderedActors(record, activities.map((activity) => activity.actor));
  const pools = systemsFrom(record.systems);
  const poolName = pools.length ? truncate(pools.join(" / "), 34) : "业务流程";
  const result = deriveResult(record, activities);
  const questions = openQuestions(record);
  const safe = activities.length >= 2 && Boolean(result) && actors.some((actor) => !/待确认/u.test(actor));
  return { poolName, actors, activities, result, questions, safe, systemNote: pools.join("；") };
}

function gatewayQuestion(activity) {
  const text = `${activity.name} ${activity.source}`;
  if (/校验/u.test(text)) return "是否校验通过";
  if (/审批/u.test(text)) return "是否审批通过";
  if (/审核/u.test(text)) return "是否审核通过";
  if (/审查/u.test(text)) return "是否审查通过";
  return "是否通过";
}

function shouldAddGateway(activity, record) {
  const supported = `${activity.name} ${activity.source}`.match(/审核|审查|审批|校验/u);
  if (!supported) return false;
  return /退回|补正|不通过|审批|审核|审查|校验/u.test(record.process);
}

function flowItems(record, flow) {
  const items = [];
  let gateways = 0;
  flow.activities.forEach((activity, index) => {
    activity.id = `activity${index + 1}`;
    items.push({ type: "activity", id: activity.id, actor: activity.actor, text: activity.name, source: activity.source });
    if (gateways < 1 && index > 0 && shouldAddGateway(activity, record)) {
      gateways += 1;
      const failActor = flow.actors[0] === activity.actor ? flow.actors[1] || flow.actors[0] : flow.actors[0];
      items.push({
        type: "gateway",
        id: `gateway${gateways}`,
        actor: activity.actor,
        text: gatewayQuestion(activity),
        failId: `gateway${gateways}Fail`,
        failActor,
        failText: /退回|补正|不通过/u.test(record.process) ? "退回补正路径待核验" : "未通过处理待核验",
        targetId: flow.activities[0]?.id || "activity1"
      });
    }
  });
  return items;
}

function layout(record, flow, items) {
  const baseLaneW = 426;
  const titleH = 38;
  const laneH = 48;
  const bodyTop = titleH + laneH;
  const rowH = 94;
  const nodeW = 230;
  const nodeH = 64;
  const gatewayW = 110;
  const gatewayH = 74;
  const lanes = Math.max(1, flow.actors.length);
  const minTitleWidth = Math.min(980, Math.max(560, (`业务事项：${record.matter}`).length * 18));
  const width = Math.max(lanes * baseLaneW, minTitleWidth);
  const laneW = width / lanes;
  const resultY = bodyTop + 48 + (items.length + 1) * rowH;
  const endY = resultY + 120;
  const bodyH = endY + 58 - bodyTop;
  const height = bodyTop + bodyH;
  const laneX = (actor) => Math.max(0, flow.actors.indexOf(actor)) * laneW;
  const laneCenter = (actor) => laneX(actor) + laneW / 2;
  const yFor = (index) => bodyTop + 48 + index * rowH;
  return { laneW, titleH, laneH, bodyTop, rowH, nodeW, nodeH, gatewayW, gatewayH, lanes, width, resultY, endY, bodyH, height, laneX, laneCenter, yFor };
}

function itemPosition(item, index, l) {
  const x = l.laneCenter(item.actor);
  const y = l.yFor(index + 1);
  return { x, y };
}

function pageXml(record, flow) {
  const items = flowItems(record, flow);
  const l = layout(record, flow, items);
  const cells = [
    "<mxCell id=\"0\"/>",
    "<mxCell id=\"1\" parent=\"0\"/>",
    cell("title", `业务事项：${record.matter}`, 0, 0, l.width, l.titleH, `rounded=0;whiteSpace=wrap;html=1;fillColor=#FFFFFF;strokeColor=${black};fontFamily=${font};fontSize=14;fontStyle=1;align=center;verticalAlign=middle;`, Math.max(24, Math.floor(l.width / 28))),
  ];
  flow.actors.forEach((actor, i) => {
    const x = i * l.laneW;
    cells.push(cell(`laneHead${i}`, actor, x, l.titleH, l.laneW, l.laneH, `rounded=0;whiteSpace=wrap;html=1;fillColor=#FFFFFF;strokeColor=${black};fontFamily=${font};fontSize=13;fontStyle=1;align=center;verticalAlign=middle;`, 17));
    cells.push(cell(`laneBody${i}`, "", x, l.bodyTop, l.laneW, l.bodyH, `rounded=0;whiteSpace=wrap;html=1;fillColor=#FFFFFF;strokeColor=${black};`));
  });
  const first = flow.activities[0];
  cells.push(cell("start", "", l.laneCenter(first.actor) - 17, l.yFor(0) - 17, 34, 34, `shape=ellipse;whiteSpace=wrap;html=1;fillColor=#FFFFFF;strokeColor=${black};strokeWidth=1;`));
  const positions = new Map();
  items.forEach((item, index) => positions.set(item.id, itemPosition(item, index, l)));
  let previous = { id: "start", type: "start" };
  let edgeNo = 0;
  items.forEach((item, index) => {
    const pos = positions.get(item.id);
    const edgeLabel = previous.type === "gateway" ? "是" : "";
    if (item.type === "activity") {
      cells.push(cell(item.id, item.text, pos.x - l.nodeW / 2, pos.y - l.nodeH / 2, l.nodeW, l.nodeH, `rounded=1;whiteSpace=wrap;html=1;fillColor=#FFFFFF;strokeColor=${black};fontFamily=${font};fontSize=13;align=center;verticalAlign=middle;spacing=6;arcSize=7;`, 12));
    } else {
      cells.push(cell(item.id, item.text, pos.x - l.gatewayW / 2, pos.y - l.gatewayH / 2, l.gatewayW, l.gatewayH, `shape=rhombus;whiteSpace=wrap;html=1;fillColor=#FFFFFF;strokeColor=${black};fontFamily=${font};fontSize=12;align=center;verticalAlign=middle;spacing=4;`, 7));
      const failX = l.laneCenter(item.failActor);
      cells.push(cell(item.failId, item.failText, failX - l.nodeW / 2, pos.y - l.nodeH / 2, l.nodeW, l.nodeH, `rounded=1;whiteSpace=wrap;html=1;fillColor=#FFFFFF;strokeColor=${black};fontFamily=${font};fontSize=12;align=center;verticalAlign=middle;spacing=6;arcSize=7;`, 12));
      cells.push(edge(`edgeNo${++edgeNo}`, item.id, item.failId, [], "否"));
      cells.push(edge(`edgeBack${edgeNo}`, item.failId, item.targetId, [], "重新提交"));
    }
    cells.push(edge(`edgeMain${index + 1}`, previous.id, item.id, [], edgeLabel));
    previous = { id: item.id, type: item.type };
  });
  const lastActor = items.at(-1)?.actor || first.actor;
  cells.push(cell("result", flow.result, l.laneCenter(lastActor) - l.nodeW / 2, l.resultY - l.nodeH / 2, l.nodeW, l.nodeH, `rounded=1;whiteSpace=wrap;html=1;fillColor=#FFFFFF;strokeColor=${black};strokeWidth=2;fontFamily=${font};fontSize=13;fontStyle=1;align=center;verticalAlign=middle;spacing=6;arcSize=7;`, 12));
  cells.push(edge("edgeResult", previous.id, "result", [], previous.type === "gateway" ? "是" : ""));
  cells.push(cell("end", "", l.laneCenter(lastActor) - 17, l.endY - 17, 34, 34, `shape=ellipse;whiteSpace=wrap;html=1;fillColor=#FFFFFF;strokeColor=${black};strokeWidth=3;`));
  cells.push(edge("edgeEnd", "result", "end"));
  const name = `行${String(record.row).padStart(2, "0")}-${record.matter}`;
  return `<diagram id="${xml(name)}" name="${xml(name)}"><mxGraphModel dx="${l.width}" dy="${l.height}" grid="1" gridSize="10" guides="1" tooltips="1" connect="1" arrows="1" fold="1" page="1" pageScale="1" pageWidth="${l.width}" pageHeight="${l.height}" math="0" shadow="0"><root>${cells.join("")}</root></mxGraphModel></diagram>`;
}

function svgText(value, x, y, className, width, lineHeight = 18) {
  return wrap(value, width).map((line, i) => `<text x="${x}" y="${y + i * lineHeight}" class="${className}">${xml(line)}</text>`).join("");
}
function svgLine(x1, y1, x2, y2, label = "") {
  const midY = Math.round((y1 + y2) / 2);
  const d = Math.abs(x1 - x2) < 4 ? `M${x1} ${y1} V${y2}` : `M${x1} ${y1} V${midY} H${x2} V${y2}`;
  const labelText = label ? `<text x="${Math.round((x1 + x2) / 2) + 8}" y="${Math.round((y1 + y2) / 2) - 6}" class="edgeLabel">${xml(label)}</text>` : "";
  return `<path d="${d}" class="arrow"/>${labelText}`;
}
function svgBackLine(x1, y1, x2, y2, label = "") {
  const elbowX = Math.min(x1, x2) - 44;
  const d = `M${x1} ${y1} H${elbowX} V${y2} H${x2}`;
  const labelText = label ? `<text x="${elbowX + 8}" y="${Math.round((y1 + y2) / 2)}" class="edgeLabel">${xml(label)}</text>` : "";
  return `<path d="${d}" class="arrow thin"/>${labelText}`;
}
function svg(record, flow) {
  const items = flowItems(record, flow);
  const l = layout(record, flow, items);
  const laneRects = flow.actors.map((actor, i) => {
    const x = i * l.laneW;
    return `<rect x="${x}" y="${l.titleH}" width="${l.laneW}" height="${l.laneH}" fill="#FFF" stroke="#000" stroke-width="2"/>${svgText(actor, x + l.laneW / 2, l.titleH + 30, "lane", 18)}<rect x="${x}" y="${l.bodyTop}" width="${l.laneW}" height="${l.bodyH}" fill="#FFF" stroke="#000" stroke-width="2"/>`;
  }).join("");
  const positions = new Map();
  items.forEach((item, index) => positions.set(item.id, itemPosition(item, index, l)));
  const first = flow.activities[0];
  const start = { x: l.laneCenter(first.actor), y: l.yFor(0) };
  let arrows = "";
  let previous = { id: "start", type: "start", x: start.x, y: start.y, bottom: start.y + 17 };
  const nodeShapes = [];
  items.forEach((item, index) => {
    const pos = positions.get(item.id);
    const label = previous.type === "gateway" ? "是" : "";
    const top = item.type === "gateway" ? pos.y - l.gatewayH / 2 : pos.y - l.nodeH / 2;
    arrows += svgLine(previous.x, previous.bottom, pos.x, top, label);
    if (item.type === "activity") {
      nodeShapes.push(`<rect x="${pos.x - l.nodeW / 2}" y="${pos.y - l.nodeH / 2}" width="${l.nodeW}" height="${l.nodeH}" rx="6" fill="#FFF" stroke="#000" stroke-width="2"/>${svgText(item.text, pos.x, pos.y - 5, "node", 12)}`);
      previous = { id: item.id, type: item.type, x: pos.x, y: pos.y, bottom: pos.y + l.nodeH / 2 };
    } else {
      const points = `${pos.x},${pos.y - l.gatewayH / 2} ${pos.x + l.gatewayW / 2},${pos.y} ${pos.x},${pos.y + l.gatewayH / 2} ${pos.x - l.gatewayW / 2},${pos.y}`;
      nodeShapes.push(`<polygon points="${points}" fill="#FFF" stroke="#000" stroke-width="2"/>${svgText(item.text, pos.x, pos.y - 9, "node small", 7)}`);
      const failX = l.laneCenter(item.failActor);
      const failY = pos.y;
      nodeShapes.push(`<rect x="${failX - l.nodeW / 2}" y="${failY - l.nodeH / 2}" width="${l.nodeW}" height="${l.nodeH}" rx="6" fill="#FFF" stroke="#000" stroke-width="2"/>${svgText(item.failText, failX, failY - 5, "node small", 12)}`);
      arrows += svgLine(pos.x - l.gatewayW / 2, pos.y, failX + l.nodeW / 2, failY, "否");
      const target = positions.get(item.targetId);
      if (target) arrows += svgBackLine(failX, failY - l.nodeH / 2, target.x - l.nodeW / 2, target.y, "重新提交");
      previous = { id: item.id, type: item.type, x: pos.x, y: pos.y, bottom: pos.y + l.gatewayH / 2 };
    }
  });
  const lastActor = items.at(-1)?.actor || first.actor;
  const resultX = l.laneCenter(lastActor);
  arrows += svgLine(previous.x, previous.bottom, resultX, l.resultY - l.nodeH / 2, previous.type === "gateway" ? "是" : "");
  arrows += svgLine(resultX, l.resultY + l.nodeH / 2, resultX, l.endY - 17);
  const titleWrap = Math.max(28, Math.floor(l.width / 32));
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${l.width}" height="${l.height}" viewBox="0 0 ${l.width} ${l.height}"><defs><marker id="arrow" markerWidth="10" markerHeight="10" refX="7" refY="3" orient="auto"><path d="M0,0 L0,6 L8,3 z" fill="#000"/></marker><style>.title{font-family:'Microsoft YaHei',Arial,sans-serif;font-size:17px;font-weight:700;text-anchor:middle}.lane{font-family:'Microsoft YaHei',Arial,sans-serif;font-size:17px;font-weight:700;text-anchor:middle}.node{font-family:'Microsoft YaHei',Arial,sans-serif;font-size:16px;text-anchor:middle}.small{font-size:14px}.edgeLabel{font-family:'Microsoft YaHei',Arial,sans-serif;font-size:13px;text-anchor:start}.arrow{fill:none;stroke:#000;stroke-width:2;marker-end:url(#arrow)}.thin{stroke-width:1.6}</style></defs><rect x="0" y="0" width="${l.width}" height="${l.titleH}" fill="#FFF" stroke="#000" stroke-width="2"/>${svgText(`业务事项：${record.matter}`, l.width / 2, 25, "title", titleWrap, 18)}${laneRects}<circle cx="${start.x}" cy="${start.y}" r="17" fill="#FFF" stroke="#000" stroke-width="1.5"/>${arrows}${nodeShapes.join("")}<rect x="${resultX - l.nodeW / 2}" y="${l.resultY - l.nodeH / 2}" width="${l.nodeW}" height="${l.nodeH}" rx="6" fill="#FFF" stroke="#000" stroke-width="3"/>${svgText(flow.result, resultX, l.resultY - 5, "node", 12)}<circle cx="${resultX}" cy="${l.endY}" r="17" fill="#FFF" stroke="#000" stroke-width="4"/></svg>`;
}

async function writePendingWorkbook(entries) {
  const wb = Workbook.create();
  const sheet = wb.worksheets.add("待核验清单");
  const header = ["事项", "来源职责", "待核验内容", "原因"];
  const rows = entries.map(({ record, flow, reason }) => [record.matter, record.responsibility, flow?.questions || "办理主体、主要流程、结果或处室交接", reason]);
  sheet.getRange(`A1:D${rows.length + 1}`).values = [header, ...rows];
  sheet.getRange("A1:D1").format = { fill: gray, font: { name: font, size: 10, bold: true, color: black }, horizontalAlignment: "center", verticalAlignment: "center", wrapText: true, borders: { preset: "all", style: "thin", color: black } };
  sheet.getRange("A1:D1").format.rowHeight = 28;
  if (rows.length) {
    sheet.getRange(`A2:D${rows.length + 1}`).format = { font: { name: font, size: 10, color: black }, verticalAlignment: "top", wrapText: true, borders: { preset: "all", style: "thin", color: black } };
    for (let r = 2; r <= rows.length + 1; r += 1) sheet.getRange(`A${r}:D${r}`).format.rowHeight = 96;
  }
  sheet.getRange("A:A").format.columnWidth = 42;
  sheet.getRange("B:B").format.columnWidth = 58;
  sheet.getRange("C:C").format.columnWidth = 54;
  sheet.getRange("D:D").format.columnWidth = 72;
  sheet.freezePanes.freezeRows(1);
  sheet.showGridLines = false;
  wb.recalculate();
  const file = await SpreadsheetFile.exportXlsx(wb);
  await file.save(pendingPath);
}

await fs.mkdir(outputDir, { recursive: true });
await fs.mkdir(imageDir, { recursive: true });

const input = await SpreadsheetFile.importXlsx(await FileBlob.load(inputPath));
const sheet = input.worksheets.getItemAt(0);
const values = sheet.getUsedRange().values;
if (values[0].map(clean).join("|") !== headers.join("|")) throw new Error("Stage 2 Excel headers are not A-I as required.");

let currentDepartment = "";
let currentResponsibility = "";
const records = [];
for (let i = 1; i < values.length; i += 1) {
  const row = values[i];
  currentDepartment = clean(row[0]) || currentDepartment;
  currentResponsibility = clean(row[1]) || currentResponsibility;
  const matter = clean(row[2]);
  if (!matter) continue;
  if (currentDepartment !== department) throw new Error(`Unexpected department on Excel row ${i + 1}`);
  records.push({
    row: i + 1,
    responsibility: currentResponsibility,
    matter,
    status: clean(row[4]),
    process: clean(row[5]),
    systems: clean(row[7]),
    basis: String(row[8] ?? ""),
    nodes: parseNodes(row[8])
  });
}

const digital = records.filter((record) => record.status === "有");
if (!digital.length) throw new Error("No digital matters were found in the stage 2 Excel.");

const assessed = digital.map((record) => ({ record, flow: buildFlow(record) }));
const drawn = assessed.filter(({ flow }) => flow.safe);
const pending = [
  ...assessed.filter(({ flow }) => !flow.safe).map(({ record, flow }) => ({ record, flow, reason: "F/I 未能形成至少两个连续业务动作、明确业务主体和结果节点，未出图。" })),
  ...assessed.filter(({ flow }) => flow.safe && flow.questions).map(({ record, flow }) => ({ record, flow, reason: "已绘制主流程；待确认内容不强行补画，需人工核验后再完善异常路径或处室权限。" }))
];

const diagrams = [];
for (const { record, flow } of drawn) {
  diagrams.push(pageXml(record, flow));
  const name = `行${String(record.row).padStart(2, "0")}-${record.matter}`;
  const imagePath = path.join(imageDir, `${name}.png`);
  const png = await sharp(Buffer.from(svg(record, flow))).png().toBuffer();
  await fs.writeFile(imagePath, png);
  const metadata = await sharp(png).metadata();
  if (!metadata.width || !metadata.height || metadata.width < 320 || metadata.height < 300) throw new Error(`Preview rendering failed for ${record.matter}`);
  record.imagePath = imagePath;
}

const drawio = `<?xml version="1.0" encoding="UTF-8"?><mxfile host="app.diagrams.net" modified="${new Date().toISOString()}" agent="Codex" version="24.7.17" type="device">${diagrams.join("")}</mxfile>`;
await fs.writeFile(drawioPath, drawio, "utf8");
if ((drawio.match(/<diagram /g) ?? []).length !== drawn.length) throw new Error("Draw.io page count validation failed.");
if (/laneHead\d+" value="[^"]*(系统|平台|数据库)/u.test(drawio)) throw new Error("Draw.io structural validation failed: system appeared as a lane.");

for (const { record } of drawn) {
  const data = await fs.readFile(record.imagePath);
  const metadata = await sharp(data).metadata();
  const displayWidth = 760;
  const displayHeight = Math.min(640, Math.round(displayWidth * (metadata.height || 1) / (metadata.width || displayWidth)));
  sheet.images.add({ dataUrl: `data:image/png;base64,${data.toString("base64")}`, anchor: { from: { row: record.row - 1, col: 6 }, extent: { widthPx: displayWidth, heightPx: displayHeight } } });
  sheet.getRange(`G${record.row}`).format.columnWidthPx = displayWidth + 20;
  sheet.getRange(`A${record.row}:I${record.row}`).format.rowHeightPx = displayHeight + 24;
}
sheet.getRange("G:G").format.horizontalAlignment = "center";
input.recalculate();
const reviewed = await SpreadsheetFile.exportXlsx(input);
await reviewed.save(reviewedPath);
await writePendingWorkbook(pending);

const reopened = await SpreadsheetFile.importXlsx(await FileBlob.load(reviewedPath));
const exported = reopened.worksheets.getItemAt(0).getUsedRange().values;
if (exported.length !== values.length || exported.some((row, r) => row.some((value, c) => clean(value) !== clean(values[r][c])))) {
  throw new Error("Reviewed workbook values changed outside image insertion.");
}

console.log(JSON.stringify({
  inputPath,
  drawioPath,
  reviewedPath,
  pendingPath,
  imageDir,
  digitalCount: digital.length,
  drawnCount: drawn.length,
  pendingCount: pending.length,
  pages: drawn.map(({ record }) => `行${String(record.row).padStart(2, "0")}-${record.matter}`)
}, null, 2));
