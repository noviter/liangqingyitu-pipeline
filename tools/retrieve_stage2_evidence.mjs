import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

function argument(name) {
  const index = process.argv.indexOf(name);
  if (index === -1 || !process.argv[index + 1]) throw new Error(`Missing ${name}`);
  return process.argv[index + 1];
}
function optionalArgument(name, fallback = "") {
  const index = process.argv.indexOf(name);
  return index === -1 ? fallback : process.argv[index + 1] ?? fallback;
}

const inputPath = argument("--input");
const indexPath = argument("--index");
const outputPath = argument("--output");
const maxCandidates = Number(argument("--max-candidates"));
const systemCodes = new Set(optionalArgument("--system-codes").split(/[，,;；\s]+/u).filter(Boolean));

function clean(value) { return String(value ?? "").replace(/\r/g, "").replace(/[ \t]+/g, " ").trim(); }
function unique(values) { return [...new Set(values.filter(Boolean))]; }
function termsFor(matter) {
  const text = clean(matter).replace(/^(承担指导|承担编制|承担国省干线公路|承担公路|参与国省干线公路|参与|督促落实|贯彻执行|承办)/u, "").replace(/事务性工作$/u, "");
  const terms = new Set([text]);
  for (let length = Math.min(10, text.length); length >= 4; length -= 1) {
    for (let start = 0; start <= text.length - length; start += 1) terms.add(text.slice(start, start + length));
  }
  return [...terms].sort((a, b) => b.length - a.length);
}
function heading(line) { return /^(第?[一二三四五六七八九十]+[、.]|[（(][一二三四五六七八九十0-9]+[)）]|\d+(?:\.\d+){0,3}[、.．]|[一二三四五六七八九十]+、)/u.test(clean(line)); }
function section(lines, lineIndex) {
  let start = Math.max(0, lineIndex - 24);
  for (let index = lineIndex; index >= Math.max(0, lineIndex - 140); index -= 1) if (heading(lines[index])) { start = index; break; }
  let end = Math.min(lines.length, lineIndex + 80);
  for (let index = lineIndex + 1; index < Math.min(lines.length, lineIndex + 220); index += 1) if (heading(lines[index])) { end = index; break; }
  return { start: start + 1, end, excerpt: lines.slice(start, end).join("\n").slice(0, 8000) };
}
function candidatesFor(text, terms) {
  const lines = text.split(/\r?\n/);
  const candidates = [];
  for (let index = 0; index < lines.length; index += 1) {
    const matched = terms.filter((term) => lines[index].includes(term));
    if (!matched.length) continue;
    const strongest = matched[0];
    const nearby = lines.slice(Math.max(0, index - 4), Math.min(lines.length, index + 5)).join("\n");
    const nearbyMatches = terms.filter((term) => nearby.includes(term));
    const extracted = section(lines, index);
    candidates.push({ score: strongest.length ** 3 + nearbyMatches.reduce((sum, term) => sum + term.length, 0), line: index + 1, matched: unique(nearbyMatches).slice(0, 12), section_start: extracted.start, section_end: extracted.end, excerpt: extracted.excerpt });
  }
  const seen = new Set();
  return candidates.filter((candidate) => {
    const key = `${candidate.section_start}-${candidate.section_end}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  }).sort((a, b) => b.score - a.score);
}

const input = await SpreadsheetFile.importXlsx(await FileBlob.load(inputPath));
const matters = input.worksheets.getItemAt(0).getUsedRange().values.slice(1).map((row) => clean(row[2])).filter(Boolean);
const index = JSON.parse(await fs.readFile(indexPath, "utf8"));
const root = path.dirname(path.dirname(indexPath));
const sources = [];
for (const item of index) {
  if (!item.text_file || !item.source_file || (systemCodes.size && !systemCodes.has(String(item.system_code)))) continue;
  try { sources.push({ ...item, text: await fs.readFile(path.resolve(root, item.text_file), "utf8") }); } catch { /* Unavailable extracted text cannot become evidence. */ }
}
if (!sources.length) throw new Error("No readable system materials remain after applying the requested system scope.");
const result = matters.map((matter) => {
  const terms = termsFor(matter);
  const candidates = [];
  for (const source of sources) {
    for (const candidate of candidatesFor(source.text, terms)) candidates.push({ system_code: String(source.system_code), system_name: source.system_name, source_file: source.source_file, ...candidate });
  }
  return { matter, system_scope: [...systemCodes], query_terms: terms.slice(0, 18), candidates: candidates.sort((a, b) => b.score - a.score).slice(0, maxCandidates) };
});
await fs.mkdir(path.dirname(outputPath), { recursive: true });
await fs.writeFile(outputPath, JSON.stringify(result, null, 2), "utf8");
console.log(JSON.stringify({ matterCount: result.length, sourceCount: sources.length, scopedSystemCodes: [...systemCodes], outputPath }, null, 2));
