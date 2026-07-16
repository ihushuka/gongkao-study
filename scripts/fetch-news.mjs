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
const absolute = (url, base) => { try { return new URL(url, base).toString(); } catch { return ""; } };
const fetchText = async url => {
  let lastError;
  for (let attempt = 0; attempt < 2; attempt += 1) {
    const controller = new AbortController(), timer = setTimeout(() => controller.abort(), 25000);
    try {
      const response = await fetch(url, { signal: controller.signal, headers: { "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/130 Safari/537.36", accept: "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8" } });
      if (!response.ok) throw new Error(`${response.status} ${url}`);
      return await response.text();
    } catch (error) { lastError = error; }
    finally { clearTimeout(timer); }
  }
  throw lastError;
};
const normalizeDateText = value => {
  const hit = String(value || "").match(/20\d{2}[-./年]\d{1,2}[-./月]\d{1,2}日?/);
  return hit ? hit[0].replace(/[./年月]/g, "-").replace(/日$/, "").split("-").map((x, i) => i ? x.padStart(2, "0") : x).join("-") : "";
};
const isArticleUrl = (link, source) => {
  try {
    const url = new URL(link), path = url.pathname.replace(/\/{2,}/g, "/").toLowerCase(), normalized = path.replace(/\/(index\.(?:s?html?|htm))?$/, "/");
    if (!/^https?:$/.test(url.protocol)) return false;
    if (["/", "/yaowen/", "/yaowen/liebiao/", "/politics/", "/news/", "/index.html", "/index.htm"].includes(path) || ["/", "/yaowen/", "/yaowen/liebiao/", "/politics/", "/news/"].includes(normalized)) return false;
    if (source === "gov") return /content_\d+\.(?:s?html?|htm)$/.test(path) || /\/20\d{4}\/\d{1,2}\/\d{1,2}\//.test(path) || (/\.(?:s?html?|htm)$/.test(path) && path.split("/").filter(Boolean).length >= 3);
    if (source === "people") return /\/n1\/20\d{2}\/\d{4}\//.test(path) || /content_\d+\.(?:s?html?|htm)$/.test(path) || (/\.(?:s?html?|htm)$/.test(path) && path.split("/").filter(Boolean).length >= 3);
    return url.searchParams.has("id") || /detail|article|static_page|lgpage/.test(path) || path.split("/").filter(Boolean).length >= 3;
  } catch { return false; }
};
const validTitle = title => title.length >= 8 && !/^(首页|更多|要闻|新闻|时政|国内|国际|中国政府网|人民日报|人民网|学习强国)(频道|栏目|首页)?$/.test(title);
const unique = (rows, source) => Array.from(new Map(rows.filter(x => validTitle(x.title) && x.link && isArticleUrl(x.link, source)).map(x => [x.link, x])).values()).slice(0, 8);

async function peopleNews() {
  const xml = await fetchText("https://www.people.com.cn/rss/politics.xml"), rows = [];
  for (const item of xml.match(/<item[\s\S]*?<\/item>/gi) || []) {
    const title = clean(item.match(/<title[^>]*>([\s\S]*?)<\/title>/i)?.[1]);
    const link = clean(item.match(/<link[^>]*>([\s\S]*?)<\/link>/i)?.[1]);
    const rawDate = clean(item.match(/<pubDate[^>]*>([\s\S]*?)<\/pubDate>/i)?.[1] || item.match(/<date[^>]*>([\s\S]*?)<\/date>/i)?.[1]);
    const date = rawDate ? shanghaiDate(rawDate) : "";
    if (date === today) rows.push({ title, link, source: "人民日报·人民网", date });
  }
  return unique(rows, "people");
}

async function govNews() {
  const base = "https://www.gov.cn/yaowen/liebiao/", html = await fetchText(base), rows = [];
  const blocks = html.match(/<li\b[\s\S]*?<\/li>/gi) || [];
  for (const block of blocks) {
    const date = normalizeDateText(block);
    if (date !== today) continue;
    const anchors = [...block.matchAll(/<a\s+[^>]*href=["']([^"']+)["'][^>]*>([\s\S]*?)<\/a>/gi)];
    for (const match of anchors) {
      const title = clean(match[2]), link = absolute(match[1], base);
      if (validTitle(title) && isArticleUrl(link, "gov")) { rows.push({ title, link, source: "中国政府网", date }); break; }
    }
  }
  // Some versions of the list page do not wrap every item in an li. Keep a
  // conservative fallback, but still reject home pages and channel pages.
  if (!rows.length) {
    const anchor = /<a\s+[^>]*href=["']([^"']+)["'][^>]*>([\s\S]*?)<\/a>/gi; let match;
    while ((match = anchor.exec(html))) {
      const around = html.slice(Math.max(0, match.index - 220), Math.min(html.length, anchor.lastIndex + 260));
      const date = normalizeDateText(around), title = clean(match[2]), link = absolute(match[1], base);
      if (date === today && validTitle(title) && isArticleUrl(link, "gov")) rows.push({ title, link, source: "中国政府网", date });
    }
  }
  return unique(rows, "gov");
}

const xuexiEndpoints = ["https://www.xuexi.cn/lgdata/1jscb6pu1n2.json", "https://www.xuexi.cn/lgdata/1crqb964p71.json", "https://www.xuexi.cn/lgdata/1ap1igfgdn2.json"];
function walk(value, rows) {
  if (!value || typeof value !== "object") return;
  if (Array.isArray(value)) { value.forEach(x => walk(x, rows)); return; }
  const title = clean(value.title || value.name || value.text), rawUrl = value.url || value.link || value.shareUrl || value.static_page_url, link = absolute(rawUrl, "https://www.xuexi.cn/"), rawDate = String(value.publishTime || value.publish_time || value.date || value.time || value.auditTime || ""), explicit = rawDate.match(/20\d{2}[-./]\d{1,2}[-./]\d{1,2}/)?.[0], numericTime = /^\d{10,13}$/.test(rawDate) ? Number(rawDate) * (rawDate.length === 10 ? 1000 : 1) : null, normalized = explicit ? explicit.replace(/[./]/g, "-").split("-").map((x, i) => i ? x.padStart(2, "0") : x).join("-") : numericTime ? shanghaiDate(numericTime) : "";
  if (validTitle(title) && link && normalized === today && isArticleUrl(link, "xuexi")) rows.push({ title, link, source: "学习强国", date: normalized });
  Object.values(value).forEach(x => walk(x, rows));
}
async function xuexiNews() { const rows = []; for (const endpoint of xuexiEndpoints) { try { walk(JSON.parse(await fetchText(endpoint)), rows); } catch (error) { console.warn(String(error)); } } return unique(rows, "xuexi"); }

const settled = await Promise.allSettled([xuexiNews(), peopleNews(), govNews()]);
const value = index => settled[index].status === "fulfilled" ? settled[index].value : [];
const output = { date: today, updatedAt: now.toISOString(), sources: { xuexi: value(0), people: value(1), gov: value(2) } };
await mkdir("public", { recursive: true });
await writeFile("public/daily-news.json", JSON.stringify(output, null, 2) + "\n", "utf8");
console.log(`daily-news ${today}: xuexi=${output.sources.xuexi.length}, people=${output.sources.people.length}, gov=${output.sources.gov.length}`);
