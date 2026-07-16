import { mkdir, readFile, writeFile } from "node:fs/promises";

const now = new Date();
const dateParts = new Intl.DateTimeFormat("en-CA", { timeZone: "Asia/Shanghai", year: "numeric", month: "2-digit", day: "2-digit" }).formatToParts(now);
const pick = type => dateParts.find(x => x.type === type)?.value;
const today = `${pick("year")}-${pick("month")}-${pick("day")}`;
const shanghaiDate = value => {
  const date = value instanceof Date ? value : new Date(value);
  if (Number.isNaN(date.getTime())) return "";
  const formatted = new Intl.DateTimeFormat("en-CA", { timeZone: "Asia/Shanghai", year: "numeric", month: "2-digit", day: "2-digit" }).formatToParts(date);
  const get = type => formatted.find(x => x.type === type)?.value;
  return `${get("year")}-${get("month")}-${get("day")}`;
};
const decodeEntities = value => String(value || "")
  .replace(/<!\[CDATA\[|\]\]>/g, "")
  .replace(/&nbsp;|&#160;/gi, " ")
  .replace(/&amp;/gi, "&")
  .replace(/&quot;/gi, '"')
  .replace(/&#39;|&apos;/gi, "'")
  .replace(/&lt;/gi, "<")
  .replace(/&gt;/gi, ">");
const clean = value => decodeEntities(value).replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim();
const absolute = (url, base) => { try { return new URL(url, base).toString(); } catch { return ""; } };
const normalizeDateText = value => {
  const source = String(value || "");
  const hit = source.match(/20\d{2}[-./年]\s*\d{1,2}[-./月]\s*\d{1,2}日?/);
  if (hit) return hit[0].replace(/\s+/g, "").replace(/[./年月]/g, "-").replace(/日$/, "").split("-").map((x, i) => i ? x.padStart(2, "0") : x).join("-");
  const compact = source.match(/20\d{6}/)?.[0];
  return compact ? `${compact.slice(0, 4)}-${compact.slice(4, 6)}-${compact.slice(6, 8)}` : "";
};
const fetchText = async url => {
  let lastError;
  for (let attempt = 0; attempt < 3; attempt += 1) {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), 30000);
    try {
      const response = await fetch(url, {
        signal: controller.signal,
        redirect: "follow",
        headers: {
          "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/130 Safari/537.36",
          accept: "text/html,application/xhtml+xml,application/xml,application/json;q=0.9,*/*;q=0.8",
          "accept-language": "zh-CN,zh;q=0.9,en;q=0.7"
        }
      });
      if (!response.ok) throw new Error(`${response.status} ${url}`);
      return await response.text();
    } catch (error) {
      lastError = error;
      if (attempt < 2) await new Promise(resolve => setTimeout(resolve, 900 * (attempt + 1)));
    } finally { clearTimeout(timer); }
  }
  throw lastError;
};
const anchorsFrom = (html, base) => {
  const rows = [];
  const pattern = /<a\b[^>]*href\s*=\s*["']([^"']+)["'][^>]*>([\s\S]*?)<\/a>/gi;
  let match;
  while ((match = pattern.exec(html))) {
    rows.push({ link: absolute(match[1], base), title: clean(match[2]), index: match.index, end: pattern.lastIndex });
  }
  return rows;
};
const validTitle = title => title.length >= 7 && !/^(首页|更多|更多>>?|要闻|新闻|时政|国内|国际|中国政府网|人民日报|人民网|学习强国|进入频道|点击进入)(频道|栏目|首页)?$/i.test(title);
const isArticleUrl = (link, source) => {
  try {
    const url = new URL(link);
    const host = url.hostname.toLowerCase();
    const path = url.pathname.replace(/\/{2,}/g, "/").toLowerCase();
    const normalized = path.replace(/\/(?:index\.(?:s?html?|htm))?$/, "/");
    if (!/^https?:$/.test(url.protocol)) return false;
    if (["/", "/gb/", "/yaowen/", "/yaowen/liebiao/", "/politics/", "/news/"].includes(normalized)) return false;
    if (source === "people" && !(host === "people.com.cn" || host.endsWith(".people.com.cn") || host.endsWith(".people.cn"))) return false;
    if (source === "gov" && !(host === "gov.cn" || host.endsWith(".gov.cn"))) return false;
    if (source === "xuexi" && !(host === "xuexi.cn" || host.endsWith(".xuexi.cn"))) return false;
    return Boolean(url.search) || /\.(?:s?html?|htm)$/.test(path) || /content|detail|article|static_page|lgpage/.test(path) || path.split("/").filter(Boolean).length >= 3;
  } catch { return false; }
};
const unique = (rows, source) => Array.from(new Map(rows
  .filter(x => validTitle(x.title) && x.link && isArticleUrl(x.link, source))
  .map(x => [x.link, x])).values()).slice(0, 10);

const peopleDateFromUrl = link => {
  try {
    const path = new URL(link).pathname;
    const hit = path.match(/\/n1\/(20\d{2})\/(\d{2})(\d{2})\//);
    return hit ? `${hit[1]}-${hit[2]}-${hit[3]}` : "";
  } catch { return ""; }
};

async function peopleNews() {
  const rows = [];
  const pages = ["https://www.people.com.cn/", "https://politics.people.com.cn/", "https://news.people.com.cn/"];
  for (const page of pages) {
    try {
      const html = await fetchText(page);
      for (const anchor of anchorsFrom(html, page)) {
        const urlDate = peopleDateFromUrl(anchor.link);
        const nearby = html.slice(Math.max(0, anchor.index - 260), Math.min(html.length, anchor.end + 260));
        const date = urlDate || normalizeDateText(nearby);
        if (date === today) rows.push({ title: anchor.title, link: anchor.link, source: "人民日报·人民网", date });
      }
    } catch (error) { console.warn(`people page failed: ${page}: ${String(error)}`); }
  }
  if (rows.length < 4) {
    for (const rssUrl of ["https://www.people.com.cn/rss/ywkx.xml", "https://www.people.com.cn/rss/politics.xml"]) {
      try {
        const xml = await fetchText(rssUrl);
        for (const item of xml.match(/<item[\s\S]*?<\/item>/gi) || []) {
          const title = clean(item.match(/<title[^>]*>([\s\S]*?)<\/title>/i)?.[1]);
          const link = clean(item.match(/<link[^>]*>([\s\S]*?)<\/link>/i)?.[1]);
          const rawDate = clean(item.match(/<pubDate[^>]*>([\s\S]*?)<\/pubDate>/i)?.[1] || item.match(/<date[^>]*>([\s\S]*?)<\/date>/i)?.[1]);
          const date = peopleDateFromUrl(link) || normalizeDateText(rawDate) || (rawDate ? shanghaiDate(rawDate) : "");
          if (date === today) rows.push({ title, link, source: "人民日报·人民网", date });
        }
      } catch (error) { console.warn(`people rss failed: ${rssUrl}: ${String(error)}`); }
    }
  }
  return unique(rows, "people");
}

const articlePageDate = html => {
  const preferred = [
    /(?:publishdate|pubdate|datepublished|发布时间|发稿时间)[^>\n]{0,120}?(20\d{2}[-./年]\s*\d{1,2}[-./月]\s*\d{1,2}日?)/i,
    /content\s*=\s*["'][^"']*?(20\d{2}[-./年]\s*\d{1,2}[-./月]\s*\d{1,2}日?)[^"']*["']/i
  ];
  for (const pattern of preferred) {
    const match = html.match(pattern)?.[1];
    if (match) return normalizeDateText(match);
  }
  return normalizeDateText(html.slice(0, 12000));
};

async function govNews() {
  const rows = [];
  const undated = [];
  const pages = ["https://www.gov.cn/", "https://www.gov.cn/yaowen/liebiao/", "https://www.gov.cn/yaowen/liebiao/home.htm"];
  for (const page of pages) {
    try {
      const html = await fetchText(page);
      for (const anchor of anchorsFrom(html, page)) {
        if (!validTitle(anchor.title) || !isArticleUrl(anchor.link, "gov")) continue;
        const nearby = html.slice(Math.max(0, anchor.index - 520), Math.min(html.length, anchor.end + 520));
        const date = normalizeDateText(nearby);
        if (date === today) rows.push({ title: anchor.title, link: anchor.link, source: "中国政府网", date });
        else if (!date && (anchor.link.includes(today.slice(0, 7).replace("-", "")) || page === "https://www.gov.cn/")) undated.push(anchor);
      }
    } catch (error) { console.warn(`gov page failed: ${page}: ${String(error)}`); }
  }
  if (rows.length < 5 && undated.length) {
    const candidates = Array.from(new Map(undated.map(x => [x.link, x])).values()).slice(0, 24);
    const checked = await Promise.allSettled(candidates.map(async anchor => {
      const html = await fetchText(anchor.link);
      const date = articlePageDate(html);
      return date === today ? { title: anchor.title, link: anchor.link, source: "中国政府网", date } : null;
    }));
    for (const result of checked) if (result.status === "fulfilled" && result.value) rows.push(result.value);
  }
  return unique(rows, "gov");
}

const xuexiEndpoints = [
  "https://www.xuexi.cn/lgdata/1jscb6pu1n2.json",
  "https://www.xuexi.cn/lgdata/1crqb964p71.json",
  "https://www.xuexi.cn/lgdata/1ap1igfgdn2.json"
];
function walk(value, rows) {
  if (!value || typeof value !== "object") return;
  if (Array.isArray(value)) { value.forEach(x => walk(x, rows)); return; }
  const title = clean(value.title || value.name || value.text);
  const rawUrl = value.url || value.link || value.shareUrl || value.static_page_url;
  const link = absolute(rawUrl, "https://www.xuexi.cn/");
  const rawDate = String(value.publishTime || value.publish_time || value.date || value.time || value.auditTime || "");
  const explicit = normalizeDateText(rawDate);
  const numericTime = /^\d{10,13}$/.test(rawDate) ? Number(rawDate) * (rawDate.length === 10 ? 1000 : 1) : null;
  const date = explicit || (numericTime ? shanghaiDate(numericTime) : "");
  if (date === today && validTitle(title) && isArticleUrl(link, "xuexi")) rows.push({ title, link, source: "学习强国", date });
  Object.values(value).forEach(x => walk(x, rows));
}
async function xuexiNews() {
  const rows = [];
  for (const endpoint of xuexiEndpoints) {
    try { walk(JSON.parse(await fetchText(endpoint)), rows); }
    catch (error) { console.warn(`xuexi endpoint failed: ${endpoint}: ${String(error)}`); }
  }
  return unique(rows, "xuexi");
}

let previous = null;
try { previous = JSON.parse(await readFile("public/daily-news.json", "utf8")); } catch { /* first run */ }
const settled = await Promise.allSettled([xuexiNews(), peopleNews(), govNews()]);
const sourceValue = (index, key) => {
  const fresh = settled[index].status === "fulfilled" ? settled[index].value : [];
  if (fresh.length) return fresh;
  if (previous?.date === today && Array.isArray(previous?.sources?.[key])) return previous.sources[key];
  return [];
};
const output = {
  date: today,
  updatedAt: now.toISOString(),
  sources: {
    xuexi: sourceValue(0, "xuexi"),
    people: sourceValue(1, "people"),
    gov: sourceValue(2, "gov")
  }
};
await mkdir("public", { recursive: true });
await writeFile("public/daily-news.json", JSON.stringify(output, null, 2) + "\n", "utf8");
console.log(`daily-news ${today}: xuexi=${output.sources.xuexi.length}, people=${output.sources.people.length}, gov=${output.sources.gov.length}`);
