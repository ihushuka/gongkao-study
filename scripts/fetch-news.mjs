import { mkdir, writeFile } from "node:fs/promises";

const now = new Date();
const parts = new Intl.DateTimeFormat("en-CA", { timeZone: "Asia/Shanghai", year: "numeric", month: "2-digit", day: "2-digit" }).formatToParts(now);
const pick = type => parts.find(x => x.type === type)?.value;
const today = `${pick("year")}-${pick("month")}-${pick("day")}`;
const shanghaiDate = value => {
  const date = value instanceof Date ? value : new Date(value);
  if (Number.isNaN(date.getTime())) return "";
  const formatted = new Intl.DateTimeFormat("en-CA", { timeZone: "Asia/Shanghai", year: "numeric", month: "2-digit", day: "2-digit" }).formatToParts(date);
  const get = type => formatted.find(x => x.type === type)?.value;
  return `${get("year")}-${get("month")}-${get("day")}`;
};
const clean = value => String(value || "").replace(/<!\[CDATA\[|\]\]>/g, "").replace(/<[^>]+>/g, "").replace(/&nbsp;|&#160;/g, " ").replace(/&amp;/g, "&").replace(/&quot;/g, '"').replace(/\s+/g, " ").trim();
const absolute = (url, base) => { try { return new URL(url, base).toString(); } catch { return base; } };
const fetchText = async url => { const controller = new AbortController(); const timer = setTimeout(() => controller.abort(), 25000); try { const response = await fetch(url, { signal: controller.signal, headers: { "user-agent": "Mozilla/5.0 ShoreStudyNews/1.0" } }); if (!response.ok) throw new Error(`${response.status} ${url}`); return await response.text(); } finally { clearTimeout(timer); } };
const unique = rows => Array.from(new Map(rows.filter(x => x.title && x.link).map(x => [x.title, x])).values()).slice(0, 8);

async function peopleNews() {
  const xml = await fetchText("https://www.people.com.cn/rss/politics.xml"), rows = [];
  for (const item of xml.match(/<item[\s\S]*?<\/item>/gi) || []) { const title = clean(item.match(/<title[^>]*>([\s\S]*?)<\/title>/i)?.[1]), link = clean(item.match(/<link[^>]*>([\s\S]*?)<\/link>/i)?.[1]), rawDate = clean(item.match(/<pubDate[^>]*>([\s\S]*?)<\/pubDate>/i)?.[1] || item.match(/<date[^>]*>([\s\S]*?)<\/date>/i)?.[1]); const date = rawDate ? shanghaiDate(rawDate) : ""; if (date === today) rows.push({ title, link, source: "人民日报·人民网", date }); }
  return unique(rows);
}

async function govNews() {
  const base = "https://www.gov.cn/yaowen/liebiao/", html = await fetchText(base), rows = [];
  const anchor = /<a\s+[^>]*href=["']([^"']+)["'][^>]*>([\s\S]*?)<\/a>/gi; let match;
  while ((match = anchor.exec(html))) { const around = html.slice(Math.max(0, match.index - 160), Math.min(html.length, anchor.lastIndex + 220)), date = around.match(/20\d{2}[-./年]\d{1,2}[-./月]\d{1,2}日?/)?.[0]?.replace(/[./年月]/g, "-").replace(/日$/, "").split("-").map((x, i) => i ? x.padStart(2, "0") : x).join("-"); const title = clean(match[2]); if (date === today && title.length >= 8) rows.push({ title, link: absolute(match[1], base), source: "中国政府网", date }); }
  return unique(rows);
}

const xuexiEndpoints = ["https://www.xuexi.cn/lgdata/1jscb6pu1n2.json", "https://www.xuexi.cn/lgdata/1crqb964p71.json", "https://www.xuexi.cn/lgdata/1ap1igfgdn2.json"];
function walk(value, rows) {
  if (!value || typeof value !== "object") return;
  if (Array.isArray(value)) { value.forEach(x => walk(x, rows)); return; }
  const title = clean(value.title || value.name || value.text), url = value.url || value.link || value.shareUrl || value.static_page_url, rawDate = String(value.publishTime || value.publish_time || value.date || value.time || value.auditTime || ""), explicit = rawDate.match(/20\d{2}[-./]\d{1,2}[-./]\d{1,2}/)?.[0], numericTime = /^\d{10,13}$/.test(rawDate) ? Number(rawDate) * (rawDate.length === 10 ? 1000 : 1) : null, normalized = explicit ? explicit.replace(/[./]/g, "-").split("-").map((x, i) => i ? x.padStart(2, "0") : x).join("-") : numericTime ? shanghaiDate(numericTime) : "";
  if (title.length >= 8 && url && normalized === today) rows.push({ title, link: absolute(url, "https://www.xuexi.cn/"), source: "学习强国", date: normalized });
  Object.values(value).forEach(x => walk(x, rows));
}
async function xuexiNews() { const rows = []; for (const endpoint of xuexiEndpoints) { try { walk(JSON.parse(await fetchText(endpoint)), rows); } catch (error) { console.warn(String(error)); } } return unique(rows); }

const settled = await Promise.allSettled([xuexiNews(), peopleNews(), govNews()]);
const value = index => settled[index].status === "fulfilled" ? settled[index].value : [];
const output = { date: today, updatedAt: now.toISOString(), sources: { xuexi: value(0), people: value(1), gov: value(2) } };
await mkdir("public", { recursive: true });
await writeFile("public/daily-news.json", JSON.stringify(output, null, 2) + "\n", "utf8");
console.log(`daily-news ${today}: xuexi=${output.sources.xuexi.length}, people=${output.sources.people.length}, gov=${output.sources.gov.length}`);
