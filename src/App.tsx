import { ChangeEvent, Dispatch, FormEvent, ReactNode, SetStateAction, useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";

type Tab = "home" | "daily" | "plan" | "study" | "practice" | "idiom" | "mistakes" | "review" | "reports" | "settings";
type SupervisionState = { delayCount?: number; postponedUntil?: string; reasons?: string[]; canceled?: boolean; promptedAt?: string; lostMinutes?: number; incompleteReason?: string };
type Task = { id: number; title: string; subject: string; minutes: number; done: boolean; date?: string; plannedStart?: string; plannedEnd?: string; supervision?: SupervisionState };
type DailyRoutine = { id: number; title: string; subject: string; minutes: number; month: string; plannedStart?: string; plannedEnd?: string; completedDates: string[]; skippedDates?: string[]; supervision?: Record<string, SupervisionState> };
type Practice = { id: number; date: string; source: string; module: string; correct: number; total: number; minutes: number };
type StudySession = { id: number; date: string; module: string; seconds: number; startTime?: string; endTime?: string; title?: string };
type ExamDates = { national: string; provincial: string };
type Mistake = { id: number; module: string; source: string; title: string; answer: string; mine: string; reason: string; image?: string; mastered: boolean };
type OcrConfidence = "高可信" | "待确认" | "疑似错误";
type Idiom = { id: number; word: string; meaning: string; example: string; level: "未掌握" | "模糊" | "已掌握"; source?: string; meaningHighlights?: string[]; ocrNote?: string; kind?: "成语" | "词语"; createdAt?: string; masteredAt?: string; ocrCrop?: string; exampleCrop?: string; candidates?: string[]; confidence?: OcrConfidence; uncertainPositions?: number[]; rawWord?: string };
type NewsItem = { title: string; link: string; source: string; date?: string; summary?: string };

const modules = ["资料分析", "判断推理", "言语理解", "数量关系", "常识判断"];
const sources = ["夸夸刷", "超大杯", "套卷", "粉笔题库", "真题", "错题重做"];
const commonIdioms = ["按部就班", "白驹过隙", "百折不挠", "抱残守缺", "杯弓蛇影", "鞭辟入里", "别出心裁", "不负众望", "不孚众望", "不刊之论", "不落窠臼", "不容置喙", "不置可否", "差强人意", "陈陈相因", "踌躇满志", "出类拔萃", "大相径庭", "独辟蹊径", "耳提面命", "方兴未艾", "分庭抗礼", "凤毛麟角", "浮光掠影", "高屋建瓴", "革故鼎新", "功亏一篑", "管中窥豹", "汗牛充栋", "好高骛远", "讳疾忌医", "匠心独运", "见微知著", "矫枉过正", "泾渭分明", "举重若轻", "开门见山", "刻舟求剑", "空穴来风", "苦心孤诣", "滥竽充数", "老生常谈", "临渊羡鱼", "洛阳纸贵", "买椟还珠", "明日黄花", "墨守成规", "南辕北辙", "泥沙俱下", "抛砖引玉", "披沙拣金", "平分秋色", "破釜沉舟", "潜移默化", "浅尝辄止", "曲高和寡", "趋之若鹜", "人浮于事", "如履薄冰", "舍本逐末", "甚嚣尘上", "矢志不渝", "首当其冲", "水到渠成", "司空见惯", "夙兴夜寐", "随波逐流", "昙花一现", "韬光养晦", "提纲挈领", "同日而语", "投鼠忌器", "推陈出新", "望其项背", "微言大义", "未雨绸缪", "蔚然成风", "无可厚非", "相得益彰", "相形见绌", "休戚相关", "循序渐进", "削足适履", "言简意赅", "一蹴而就", "一以贯之", "饮鸩止渴", "缘木求鱼", "责无旁贷", "真知灼见", "振聋发聩", "筚路蓝缕", "擘肌分理"];
const nav: { id: Tab; label: string; mark: string }[] = [
  { id: "home", label: "今日主页", mark: "今" },
  { id: "daily", label: "每日计划", mark: "日" },
  { id: "plan", label: "月度计划", mark: "计" },
  { id: "study", label: "学习计时", mark: "时" },
  { id: "practice", label: "刷题统计", mark: "练" },
  { id: "idiom", label: "成语积累", mark: "词" },
  { id: "mistakes", label: "错题本", mark: "错" },
  { id: "review", label: "复习中心", mark: "复" },
  { id: "reports", label: "学习报告", mark: "报" },
  { id: "settings", label: "考试与主题", mark: "设" },
];

const pad = (n: number) => String(n).padStart(2, "0");
const localISO = (date = new Date()) => `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`;
const addDays = (offset: number) => { const d = new Date(); d.setDate(d.getDate() + offset); return localISO(d); };
const normalizedDate = (value?: string) => {
  if (!value) return localISO();
  if (/^\d{4}-\d{2}-\d{2}$/.test(value)) return value;
  if (/^\d{2}-\d{2}$/.test(value)) return `${new Date().getFullYear()}-${value}`;
  return localISO();
};
const formatShortDate = (value: string) => normalizedDate(value).slice(5).replace("-", "/");
const formatClock = (seconds: number) => `${pad(Math.floor(seconds / 3600))}:${pad(Math.floor((seconds % 3600) / 60))}:${pad(seconds % 60)}`;
const formatHours = (seconds: number) => `${Math.floor(seconds / 3600)}h ${Math.floor((seconds % 3600) / 60)}m`;
const timeMinutes = (value?: string) => { const [h, m] = (value || "").split(":").map(Number); return Number.isFinite(h) && Number.isFinite(m) ? h * 60 + m : null; };
const validTimeRange = (start?: string, end?: string) => (timeMinutes(start) ?? 0) < (timeMinutes(end) ?? 0);
const routineApplies = (routine: DailyRoutine, date: string) => !(routine.skippedDates || []).includes(date) && !(/除周日|周日除外/.test(routine.title) && new Date(`${date}T12:00:00`).getDay() === 0);
const daysUntil = (value: string) => value ? Math.max(0, Math.ceil((new Date(`${value}T23:59:59`).getTime() - Date.now()) / 86400000)) : null;
const inferSubject = (title: string) => {
  if (/资料|速算/.test(title)) return "资料分析";
  if (/言语|成语|人民日报/.test(title)) return "言语理解";
  if (/判断|图推|逻辑/.test(title)) return "判断推理";
  if (/数量|数学/.test(title)) return "数量关系";
  if (/常识/.test(title)) return "常识判断";
  if (/错题|复盘|模考/.test(title)) return "错题复盘";
  return "其他计划";
};
const inferMinutes = (title: string) => Number(title.match(/(\d+)\s*(?:min|分钟)/i)?.[1] || 45);
const isDirectNewsArticle = (item: NewsItem, source: "xuexi" | "people" | "gov") => {
  const title = item.title.trim();
  if (!title || /^(首页|更多|要闻|新闻|时政|国内|国际|中国政府网|人民日报|人民网|学习强国)(频道|栏目|首页)?$/.test(title)) return false;
  try {
    const url = new URL(item.link);
    if (!/^https?:$/.test(url.protocol)) return false;
    const path = url.pathname.replace(/\/{2,}/g, "/").toLowerCase();
    const normalized = path.replace(/\/(index\.(?:s?html?|htm))?$/, "/");
    const generic = new Set(["/", "/yaowen/", "/yaowen/liebiao/", "/politics/", "/news/", "/index.html", "/index.htm"]);
    if (generic.has(path) || generic.has(normalized)) return false;
    if (source === "gov") return /content_\d+\.(?:s?html?|htm)$/.test(path) || /\/20\d{4}\/\d{1,2}\/\d{1,2}\//.test(path) || (/\.(?:s?html?|htm)$/.test(path) && path.split("/").filter(Boolean).length >= 3);
    if (source === "people") return /\/n1\/20\d{2}\/\d{4}\//.test(path) || /content_\d+\.(?:s?html?|htm)$/.test(path) || (/\.(?:s?html?|htm)$/.test(path) && path.split("/").filter(Boolean).length >= 3);
    return url.searchParams.has("id") || /detail|article|static_page|lgpage/.test(path) || path.split("/").filter(Boolean).length >= 3;
  } catch { return false; }
};

type TimelineLayoutEntry<T> = { item: T; lane: number; laneCount: number };
function layoutTimelineEntries<T>(items: T[], getStart: (item: T) => string | undefined, getEnd: (item: T) => string | undefined): TimelineLayoutEntry<T>[] {
  const normalized = items.map(item => {
    const start = timeMinutes(getStart(item));
    const rawEnd = timeMinutes(getEnd(item));
    if (start === null) return null;
    // A short card still needs enough visual height for two text lines. Treat that
    // visual footprint as occupied time so adjacent cards never paint over it.
    const end = Math.max(rawEnd ?? start + 45, start + 48);
    return { item, start, end };
  }).filter((entry): entry is { item: T; start: number; end: number } => Boolean(entry)).sort((a, b) => a.start - b.start || a.end - b.end);
  const result: TimelineLayoutEntry<T>[] = [];
  let cluster: typeof normalized = [];
  let clusterEnd = -Infinity;
  const flush = () => {
    if (!cluster.length) return;
    const laneEnds: number[] = [];
    const placed = cluster.map(entry => {
      let lane = laneEnds.findIndex(end => end <= entry.start);
      if (lane < 0) lane = laneEnds.length;
      laneEnds[lane] = entry.end;
      return { item: entry.item, lane };
    });
    const laneCount = Math.max(1, laneEnds.length);
    placed.forEach(entry => result.push({ ...entry, laneCount }));
    cluster = [];
    clusterEnd = -Infinity;
  };
  normalized.forEach(entry => {
    if (cluster.length && entry.start >= clusterEnd) flush();
    cluster.push(entry);
    clusterEnd = Math.max(clusterEnd, entry.end);
  });
  flush();
  return result;
}
const excelDate = (value: unknown) => {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return localISO(value);
  if (typeof value === "number" && value > 25000 && value < 80000) {
    const parsed = new Date(Date.UTC(1899, 11, 30) + Math.floor(value) * 86400000);
    return `${parsed.getUTCFullYear()}-${pad(parsed.getUTCMonth() + 1)}-${pad(parsed.getUTCDate())}`;
  }
  if (typeof value === "string") {
    const clean = value.trim();
    const direct = clean.match(/^(20\d{2})[./-](\d{1,2})[./-](\d{1,2})$/);
    if (direct) return `${direct[1]}-${pad(Number(direct[2]))}-${pad(Number(direct[3]))}`;
  }
  return "";
};
const cleanPlanText = (value: unknown) => typeof value === "string" ? value.trim().replace(/^[□☑✓√]+\s*/, "") : "";

function recognizePlanWorkbook(buffer: ArrayBuffer) {
  // Keep Excel dates as serial numbers. Converting them to JavaScript Date objects
  // can shift a day when the browser timezone differs from the workbook timezone.
  const workbook = XLSX.read(buffer, { type: "array", cellDates: false, cellText: false });
  let best: { tasks: Omit<Task, "id" | "done">[]; routines: Omit<DailyRoutine, "id" | "completedDates">[]; sheet: string; dated: number } = { tasks: [], routines: [], sheet: "", dated: 0 };
  const validTaskText = (text: string) => Boolean(text)
    && !/^(每日计划|工作计划|MONTHLY WORK PLAN|星期[一二三四五六日天]|周[一二三四五六日天]|一|二|三|四|五|六|日)$/.test(text)
    && !/^\d{1,4}$/.test(text)
    && !/^20\d{2}[./-]\d{1,2}[./-]\d{1,2}$/.test(text);

  workbook.SheetNames.forEach(sheetName => {
    const rows = XLSX.utils.sheet_to_json<unknown[]>(workbook.Sheets[sheetName], { header: 1, raw: true, defval: null });
    const dateRows = rows.map((row, rowIndex) => ({
      rowIndex,
      dates: row.map((cell, col) => ({ col, date: excelDate(cell) })).filter(x => x.date),
    })).filter(x => x.dates.length >= 2);

    const datedTasks: Omit<Task, "id" | "done">[] = [];
    dateRows.forEach(entry => {
      entry.dates.forEach(({ col, date }) => {
        // Find the next date row in this same column group. This prevents the small
        // helper calendar on the left from cutting into the main task calendar.
        const nextSameColumnRow = dateRows.find(next => next.rowIndex > entry.rowIndex && next.dates.some(candidate => Math.abs(candidate.col - col) <= 1));
        const stop = Math.min(nextSameColumnRow?.rowIndex ?? rows.length, entry.rowIndex + 5);
        const titles: string[] = [];
        for (let r = entry.rowIndex + 1; r < stop; r += 1) {
          // In this workbook each day uses a checkbox column followed by a text column.
          // Prefer the text column, but also support merged cells beginning on the date column.
          [col + 1, col].forEach(candidateCol => {
            const title = cleanPlanText(rows[r]?.[candidateCol]);
            if (validTaskText(title) && !titles.includes(title)) titles.push(title);
          });
        }
        titles.forEach(title => datedTasks.push({ title, subject: inferSubject(title), minutes: inferMinutes(title), date }));
      });
    });

    const planHeader = rows.findIndex(row => row.some(cell => cleanPlanText(cell) === "每日计划"));
    const recurringTitles = planHeader < 0 ? [] : rows.slice(planHeader + 1, planHeader + 9)
      .flatMap(row => row.map(cleanPlanText))
      .filter(validTaskText);
    const dates = dateRows.flatMap(x => x.dates.map(d => d.date));
    const monthCounts = dates.reduce<Record<string, number>>((acc, d) => { const m = d.slice(0, 7); acc[m] = (acc[m] || 0) + 1; return acc; }, {});
    const targetMonth = Object.entries(monthCounts).sort((a, b) => b[1] - a[1])[0]?.[0];
    const unique = Array.from(new Map(datedTasks.map(t => [`${t.date}|${t.title}`, t])).values());
    const routineTemplates = targetMonth ? Array.from(new Set(recurringTitles)).map(title => ({ title, subject: inferSubject(title), minutes: inferMinutes(title), month: targetMonth })) : [];
    if (unique.length + routineTemplates.length > best.tasks.length + best.routines.length) best = { tasks: unique, routines: routineTemplates, sheet: sheetName, dated: unique.length };
  });
  return best;
}

type CanvasBox = { x: number; y: number; width: number; height: number };
const cropCanvas = (source: HTMLCanvasElement, box: CanvasBox) => { const out = document.createElement("canvas"); out.width = Math.max(1, Math.round(box.width)); out.height = Math.max(1, Math.round(box.height)); out.getContext("2d")?.drawImage(source, box.x, box.y, box.width, box.height, 0, 0, out.width, out.height); return out; };
const previewDataUrl = (source: HTMLCanvasElement, maxWidth = 760) => { const scale = Math.min(1, maxWidth / source.width), out = document.createElement("canvas"); out.width = Math.max(1, Math.round(source.width * scale)); out.height = Math.max(1, Math.round(source.height * scale)); out.getContext("2d")?.drawImage(source, 0, 0, out.width, out.height); return out.toDataURL("image/jpeg", .76); };
const rowGroups = (rows: number[]) => rows.reduce<Array<[number, number]>>((groups, y) => { const last = groups[groups.length - 1]; if (!last || y > last[1] + 2) groups.push([y, y]); else last[1] = y; return groups; }, []);
function findGreenBoxes(canvas: HTMLCanvasElement) {
  const ctx = canvas.getContext("2d"); if (!ctx) return [];
  const { width, height } = canvas, pixels = ctx.getImageData(0, 0, width, height).data, greenRows: number[] = [];
  for (let y = 0; y < height; y += 2) { let count = 0; for (let x = Math.floor(width * .45); x < width; x += 2) { const i = (y * width + x) * 4, r = pixels[i], g = pixels[i + 1], b = pixels[i + 2]; if (g > 115 && g > r * 1.18 && g > b * 1.25) count += 1; } if (count > width * .09) greenRows.push(y); }
  const lines = rowGroups(greenRows).map(([a, b]) => Math.round((a + b) / 2)), boxes: CanvasBox[] = [];
  for (let i = 0; i + 1 < lines.length; i += 2) { const top = lines[i], bottom = lines[i + 1]; if (bottom - top < 70) continue; let minX = width, maxX = 0; for (const y of [top, bottom]) for (let x = Math.floor(width * .42); x < width; x += 2) { const p = (y * width + x) * 4, r = pixels[p], g = pixels[p + 1], b = pixels[p + 2]; if (g > 115 && g > r * 1.18 && g > b * 1.25) { minX = Math.min(minX, x); maxX = Math.max(maxX, x); } } if (maxX - minX > width * .25) boxes.push({ x: Math.max(0, minX - 8), y: Math.max(0, top - 6), width: Math.min(width - minX + 8, maxX - minX + 16), height: bottom - top + 12 }); }
  return boxes;
}
function findBlueBands(canvas: HTMLCanvasElement, box: CanvasBox) {
  const ctx = canvas.getContext("2d"); if (!ctx) return [];
  const { width } = canvas, pixels = ctx.getImageData(0, 0, canvas.width, canvas.height).data, rows: number[] = [];
  for (let y = Math.round(box.y); y < box.y + box.height; y += 2) { let count = 0; for (let x = Math.round(box.x); x < box.x + box.width; x += 2) { const i = (y * width + x) * 4, r = pixels[i], g = pixels[i + 1], b = pixels[i + 2]; if (b > 160 && g > 135 && r < 205 && b - r > 18 && g - r > 3) count += 1; } if (count > 8) rows.push(y); }
  const grouped = rowGroups(rows).filter(([a, b]) => b - a > 5), merged: Array<[number, number]> = [];
  grouped.forEach(([a, b]) => { const last = merged[merged.length - 1]; if (last && a - last[1] <= 34) last[1] = b; else merged.push([a, b]); });
  return merged.map(([a, b]) => { let minX = box.x + box.width, maxX = box.x; for (let y = a; y <= b; y += 2) for (let x = Math.round(box.x); x < box.x + box.width; x += 2) { const i = (y * width + x) * 4, r = pixels[i], g = pixels[i + 1], blue = pixels[i + 2]; if (blue > 160 && g > 135 && r < 205 && blue - r > 18 && g - r > 3) { minX = Math.min(minX, x); maxX = Math.max(maxX, x); } } return { x: Math.max(0, minX - 12), y: Math.max(0, a - 8), width: Math.max(40, maxX - minX + 24), height: b - a + 16 }; });
}
const parseGreenText = (text: string) => { const cleaned = text.replace(/\s+/g, ""); const rows: { word: string; meaning: string }[] = []; const regex = /[（(]\d+[)）]([^：:]{2,10})[：:]([\s\S]*?)(?=[（(]\d+[)）]|$)/g; let match: RegExpExecArray | null; while ((match = regex.exec(cleaned))) rows.push({ word: match[1].replace(/[^㐀-鿿]/g, ""), meaning: match[2].trim() }); return rows.filter(x => x.word.length >= 2); };
const characterDistance = (a: string, b: string) => { const left = Array.from(a), right = Array.from(b); const rows = Array.from({ length: left.length + 1 }, (_, i) => Array.from({ length: right.length + 1 }, (__, j) => i || j)); for (let i = 1; i <= left.length; i += 1) for (let j = 1; j <= right.length; j += 1) rows[i][j] = Math.min(rows[i - 1][j] + 1, rows[i][j - 1] + 1, rows[i - 1][j - 1] + (left[i - 1] === right[j - 1] ? 0 : 1)); return rows[left.length][right.length]; };
const findExample = (text: string, word: string) => {
  const lines = text.split(/\r?\n/).map(x => x.replace(/\s+/g, "").replace(/^\d+[、.．]?/, "")).filter(x => x.length >= word.length);
  const exactLine = lines.find(x => x.includes(word)); if (exactLine) return exactLine;
  const matches: { line: string; candidate: string; distance: number }[] = [];
  for (const line of lines) for (let i = 0; i <= line.length - word.length; i += 1) { const candidate = line.slice(i, i + word.length); if (/^[㐀-鿿]+$/.test(candidate)) matches.push({ line, candidate, distance: characterDistance(word, candidate) }); }
  const best = matches.sort((a, b) => a.distance - b.distance)[0];
  return best && best.distance <= 1 ? best.line.replace(best.candidate, word) : "待校对例句";
};
const exampleCandidates = (text: string, length: number) => { const clean = text.replace(/\s+/g, ""); const result = new Set<string>(); for (const segment of clean.split(/[^㐀-鿿]+/).filter(Boolean)) for (let i = 0; i <= segment.length - length; i += 1) result.add(segment.slice(i, i + length)); return Array.from(result); };
const correctIdiomFromExample = (rawWord: string, leftText: string) => {
  const word = rawWord.replace(/[^㐀-鿿]/g, ""); if (word.length < 2) return { word, note: "字数异常，请人工校对", candidates: [], confidence: "疑似错误" as OcrConfidence, uncertainPositions: Array.from({ length: word.length }, (_, i) => i) };
  const pool = Array.from(new Set([...commonIdioms.filter(x => x.length === word.length), ...exampleCandidates(leftText, word.length)]));
  const ranked = pool.map(candidate => ({ candidate, distance: characterDistance(word, candidate), inExample: leftText.replace(/\s+/g, "").includes(candidate), known: commonIdioms.includes(candidate) })).sort((a, b) => a.distance - b.distance || Number(b.inExample) - Number(a.inExample) || Number(b.known) - Number(a.known));
  const candidates = ranked.filter(x => x.distance <= 2 && (x.inExample || x.known)).slice(0, 3).map(x => x.candidate), cleanLeft = leftText.replace(/\s+/g, "");
  if (cleanLeft.includes(word) && commonIdioms.includes(word)) return { word, note: "左侧例句与本地成语词库均已确认", candidates: [word, ...candidates.filter(x => x !== word)].slice(0, 3), confidence: "高可信" as OcrConfidence, uncertainPositions: [] };
  if (cleanLeft.includes(word) || commonIdioms.includes(word)) return { word, note: cleanLeft.includes(word) ? "已与左侧例句核对" : "已通过本地成语词库核验", candidates: [word, ...candidates.filter(x => x !== word)].slice(0, 3), confidence: "高可信" as OcrConfidence, uncertainPositions: [] };
  const best = ranked[0]; if (best && best.distance <= 1 && (best.inExample || best.known)) { const uncertainPositions = Array.from(word).map((char, i) => char === Array.from(best.candidate)[i] ? -1 : i).filter(i => i >= 0); return { word: best.candidate, note: `疑似错字，已由${best.inExample ? "左侧例句" : "成语词库"}校正：${rawWord} → ${best.candidate}`, candidates, confidence: "待确认" as OcrConfidence, uncertainPositions }; }
  return { word, note: "词库与例句均未确认，请重点校对", candidates, confidence: "疑似错误" as OcrConfidence, uncertainPositions: Array.from({ length: word.length }, (_, i) => i) };
};
async function recognizePeopleDailyPdf(file: File, onProgress: (text: string, value: number) => void) {
  const pdfjs = await import("pdfjs-dist");
  pdfjs.GlobalWorkerOptions.workerSrc = new URL("pdfjs-dist/build/pdf.worker.min.mjs", import.meta.url).toString();
  const { createWorker } = await import("tesseract.js");
  onProgress("正在加载中文识别模型（首次使用时间较长）", 4);
  const worker = await createWorker("chi_sim", undefined, { langPath: new URL("./ocr/", window.location.href).toString(), logger: message => { if (message.status === "recognizing text") onProgress("正在识别页面文字", Math.min(95, 8 + Math.round((message.progress || 0) * 20))); } });
  const pdf = await pdfjs.getDocument({ data: new Uint8Array(await file.arrayBuffer()) }).promise;
  const sourceDate = file.name.match(/(\d{1,2})[._-](\d{1,2})/); const source = `人民日报${sourceDate ? `${Number(sourceDate[1])}.${Number(sourceDate[2])}` : file.name.replace(/\.pdf$/i, "")}`;
  const found: Idiom[] = [];
  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
    onProgress(`正在处理第 ${pageNumber}/${pdf.numPages} 页`, Math.round(pageNumber / pdf.numPages * 70));
    const page = await pdf.getPage(pageNumber), viewport = page.getViewport({ scale: 2.25 }), canvas = document.createElement("canvas"); canvas.width = Math.round(viewport.width); canvas.height = Math.round(viewport.height); const context = canvas.getContext("2d"); if (!context) continue; await page.render({ canvas, canvasContext: context, viewport }).promise;
    const left = cropCanvas(canvas, { x: 0, y: 0, width: canvas.width * .52, height: canvas.height }), exampleCrop = previewDataUrl(left, 620);
    await worker.setParameters({ tessedit_pageseg_mode: 6, preserve_interword_spaces: "1" } as never);
    const leftText = (await worker.recognize(left)).data.text;
    for (const box of findGreenBoxes(canvas)) {
      const greenCanvas = cropCanvas(canvas, box), ocrCrop = previewDataUrl(greenCanvas, 620), greenText = (await worker.recognize(greenCanvas)).data.text, entries = parseGreenText(greenText);
      const blueTexts: string[] = [];
      for (const blue of findBlueBands(canvas, box)) { const key = (await worker.recognize(cropCanvas(canvas, blue))).data.text.replace(/\s+/g, "").replace(/[^㐀-鿿，。；、]/g, ""); if (key.length >= 2) blueTexts.push(key); }
      entries.forEach((entry, index) => { const corrected = correctIdiomFromExample(entry.word, leftText); found.push({ id: Date.now() + found.length, rawWord: entry.word, word: corrected.word, meaning: entry.meaning, example: findExample(leftText, corrected.word), level: "未掌握", source, meaningHighlights: blueTexts.length ? [blueTexts[Math.min(blueTexts.length - 1, Math.floor(index * blueTexts.length / entries.length))]] : [], ocrNote: corrected.note, kind: corrected.word.length === 4 ? "成语" : "词语", ocrCrop, exampleCrop, candidates: corrected.candidates, confidence: corrected.confidence, uncertainPositions: corrected.uncertainPositions }); });
    }
  }
  await worker.terminate(); onProgress("识别完成，请校对后导入", 100);
  return Array.from(new Map(found.map(item => [`${item.source}|${item.word}`, item])).values());
}

async function verifyIdiomsOnline(rows: Idiom[], onProgress: (text: string, value: number) => void) {
  const checked: Idiom[] = [];
  for (let index = 0; index < rows.length; index += 1) {
    const item = rows[index]; onProgress(`正在智能查询成语 ${index + 1}/${rows.length}`, 96 + Math.round((index + 1) / Math.max(1, rows.length) * 3));
    try {
      const exactUrl = `https://zh.wiktionary.org/w/api.php?action=query&titles=${encodeURIComponent(item.word)}&format=json&formatversion=2&origin=*`;
      const exact = await fetch(exactUrl).then(response => response.ok ? response.json() : Promise.reject(new Error("dictionary unavailable")));
      if (!exact?.query?.pages?.[0]?.missing) { checked.push({ ...item, ocrNote: `${item.ocrNote || "本地校验完成"}；联网词典已确认存在`, confidence: item.confidence === "疑似错误" ? "待确认" : "高可信" }); continue; }
      const searchUrl = `https://zh.wiktionary.org/w/api.php?action=query&list=search&srsearch=${encodeURIComponent(item.word)}&srnamespace=0&srlimit=6&format=json&origin=*`;
      const search = await fetch(searchUrl).then(response => response.ok ? response.json() : Promise.reject(new Error("dictionary search unavailable")));
      const suggestion = (search?.query?.search || []).map((x: { title?: string }) => x.title?.replace(/[^㐀-鿿]/g, "") || "").filter((x: string) => x.length === item.word.length).sort((a: string, b: string) => characterDistance(item.word, a) - characterDistance(item.word, b))[0];
      if (suggestion && characterDistance(item.word, suggestion) <= 1 && item.example.includes(suggestion)) checked.push({ ...item, word: suggestion, candidates: Array.from(new Set([suggestion, ...(item.candidates || [])])).slice(0, 3), confidence: "待确认", ocrNote: `联网词典与左侧例句共同校正：${item.word} → ${suggestion}` });
      else checked.push({ ...item, confidence: item.confidence || "疑似错误", ocrNote: `${item.ocrNote || ""}；联网词典未找到完全一致条目，请重点校对`.replace(/^；/, "") });
    } catch { checked.push({ ...item, ocrNote: `${item.ocrNote || "本地校验完成"}；联网词典暂不可用`.replace(/^；/, "") }); }
  }
  onProgress("识别与智能校验完成，请确认后导入", 100); return checked;
}

const seedTasks: Task[] = [
  { id: 1, title: "资料分析｜超大杯第 12 组", subject: "资料分析", minutes: 50, done: true, date: localISO() },
  { id: 2, title: "言语理解｜逻辑填空 30 题", subject: "言语理解", minutes: 45, done: true, date: localISO() },
  { id: 3, title: "判断推理｜夸夸刷图推", subject: "判断推理", minutes: 40, done: false, date: localISO() },
  { id: 4, title: "复习昨日错题 12 道", subject: "错题复盘", minutes: 35, done: false, date: addDays(1) },
];
const seedPractices: Practice[] = [
  { id: 1, date: addDays(-12), source: "夸夸刷", module: "言语理解", correct: 35, total: 50, minutes: 44 },
  { id: 2, date: addDays(-9), source: "超大杯", module: "资料分析", correct: 15, total: 20, minutes: 25 },
  { id: 3, date: addDays(-7), source: "套卷", module: "判断推理", correct: 30, total: 40, minutes: 36 },
  { id: 4, date: addDays(-5), source: "夸夸刷", module: "判断推理", correct: 34, total: 40, minutes: 33 },
  { id: 5, date: addDays(-3), source: "超大杯", module: "数量关系", correct: 9, total: 15, minutes: 25 },
  { id: 6, date: addDays(-1), source: "套卷", module: "言语理解", correct: 32, total: 40, minutes: 35 },
  { id: 7, date: localISO(), source: "超大杯", module: "资料分析", correct: 18, total: 20, minutes: 22 },
];
const seedSessions: StudySession[] = [-6, -5, -4, -3, -2, -1].map((day, i) => ({ id: 100 + i, date: addDays(day), module: modules[i % modules.length], seconds: [7200, 10800, 9000, 12600, 10200, 13800][i] }));
const seedMistakes: Mistake[] = [
  { id: 1, module: "资料分析", source: "超大杯", title: "基期量与现期量混淆", answer: "B", mine: "C", reason: "审题错误", mastered: false },
  { id: 2, module: "判断推理", source: "夸夸刷", title: "黑白块位置规律判断", answer: "D", mine: "A", reason: "规律遗漏", mastered: false },
  { id: 3, module: "言语理解", source: "套卷", title: "成语语境对应关系", answer: "C", mine: "B", reason: "词义模糊", mastered: true },
];
const seedIdioms: Idiom[] = [
  { id: 1, word: "筚路蓝缕", meaning: "形容创业的艰苦。", example: "先辈筚路蓝缕，奠定了事业基础。", level: "模糊" },
  { id: 2, word: "缘木求鱼", meaning: "方向或办法不对，不可能达到目的。", example: "脱离实际寻找答案，无异于缘木求鱼。", level: "已掌握" },
  { id: 3, word: "不孚众望", meaning: "不能使大家信服，未符合众人的期望。", example: "这次结果不孚众望，引发了不少质疑。", level: "未掌握" },
  { id: 4, word: "擘肌分理", meaning: "比喻分析事理十分细致。", example: "报告擘肌分理地梳理了问题成因。", level: "模糊" },
];
const quotes = ["今天不必惊艳，只要比昨天更扎实一点。", "把会做的做对，把不会的逐个变成会做。", "稳定不是慢，稳定是最可靠的加速。", "每一道认真复盘的错题，都在替未来的你加分。", "别等状态，先完成今天的第一道题。"];
const dailyKnowledge = [
  { category: "政治", question: "我国的根本政治制度是什么？", answer: "人民代表大会制度。" },
  { category: "法律", question: "我国的根本大法是什么？", answer: "《中华人民共和国宪法》。" },
  { category: "人文", question: "“史家之绝唱，无韵之离骚”评价的是哪部作品？", answer: "司马迁的《史记》。" },
  { category: "地理", question: "我国地势的总体特征是什么？", answer: "西高东低，呈三级阶梯状分布。" },
  { category: "科技", question: "声音能否在真空中传播？", answer: "不能，声音传播需要介质。" },
  { category: "经济", question: "居民消费价格指数通常简称什么？", answer: "CPI。" },
  { category: "历史", question: "我国历史上第一个统一的中央集权封建国家是？", answer: "秦朝。" },
  { category: "公文", question: "适用于表彰先进、批评错误、传达重要精神的公文文种是？", answer: "通报。" },
  { category: "生态", question: "我国的基本国策中，与资源环境直接相关的是？", answer: "节约资源和保护环境。" },
  { category: "哲学", question: "唯物辩证法的实质和核心是什么？", answer: "对立统一规律。" },
];

function useStoredState<T>(key: string, initial: T) {
  const [value, setValue] = useState<T>(initial);
  const [ready, setReady] = useState(false);
  useEffect(() => {
    try { const saved = localStorage.getItem(key); if (saved) setValue(JSON.parse(saved)); } catch { /* keep defaults */ }
    setReady(true);
  }, [key]);
  useEffect(() => { if (ready) localStorage.setItem(key, JSON.stringify(value)); }, [key, ready, value]);
  return [value, setValue] as const;
}

type SiteBackup = { app: "上岸手账"; version: 9; exportedAt: string; storage: Record<string, string> };
function createSiteBackup(): SiteBackup {
  const storage: Record<string, string> = {};
  for (let i = 0; i < localStorage.length; i += 1) { const key = localStorage.key(i); if (key?.startsWith("shore-") && !key.startsWith("shore-news-") && !["shore-auto-backup-snapshot", "shore-last-auto-backup-at"].includes(key)) storage[key] = localStorage.getItem(key) || ""; }
  return { app: "上岸手账", version: 9, exportedAt: new Date().toISOString(), storage };
}
function downloadBackupFile(backup: SiteBackup, prefix = "全站备份") {
  const blob = new Blob([JSON.stringify(backup, null, 2)], { type: "application/json;charset=utf-8" }), link = document.createElement("a");
  link.href = URL.createObjectURL(blob); link.download = `上岸手账_${prefix}_${localISO()}.json`; link.click(); window.setTimeout(() => URL.revokeObjectURL(link.href), 500);
}

export default function App() {
  const [tab, setTab] = useState<Tab>("home");
  const [tasks, setTasks] = useStoredState<Task[]>("shore-tasks", seedTasks);
  const [routines, setRoutines] = useStoredState<DailyRoutine[]>("shore-daily-routines", []);
  const [practices, setPractices] = useStoredState<Practice[]>("shore-practices", seedPractices);
  const [sessions, setSessions] = useStoredState<StudySession[]>("shore-study-sessions", seedSessions);
  const [examDates, setExamDates] = useStoredState<ExamDates>("shore-exam-dates", { national: "", provincial: "" });
  const [mistakes, setMistakes] = useStoredState<Mistake[]>("shore-mistakes", seedMistakes);
  const [idioms, setIdioms] = useStoredState<Idiom[]>("shore-idioms", seedIdioms);
  const [timerOn, setTimerOn] = useState(false);
  const [activeSeconds, setActiveSeconds] = useState(0);
  const [timerModule, setTimerModule] = useState("资料分析");
  const [timerTitle, setTimerTitle] = useStoredState<string>("shore-timer-title", "今日重点任务");
  const [navOrder, setNavOrder] = useStoredState<Tab[]>("shore-nav-order", nav.map(item => item.id));
  const [dragTab, setDragTab] = useState<Tab | null>(null);
  const [timerStartedAt, setTimerStartedAt] = useState<Date | null>(null);
  const [mobileNav, setMobileNav] = useState(false);
  const [toast, setToast] = useState("");
  const timerRef = useRef<ReturnType<typeof setInterval> | null>(null);

  useEffect(() => {
    if (timerOn) timerRef.current = setInterval(() => setActiveSeconds(s => s + 1), 1000);
    else if (timerRef.current) clearInterval(timerRef.current);
    return () => { if (timerRef.current) clearInterval(timerRef.current); };
  }, [timerOn]);
  useEffect(() => {
    const saved = localStorage.getItem("shore-theme");
    if (!saved) return;
    try { const t = JSON.parse(saved); const r = document.documentElement; r.style.setProperty("--primary", t.primary); r.style.setProperty("--accent", t.accent); r.style.setProperty("--page", t.bg); r.style.setProperty("--ink", t.ink); } catch { /* keep defaults */ }
  }, []);
  useEffect(() => {
    if (localStorage.getItem("shore-weekly-backup-enabled") !== "true") return;
    const last = localStorage.getItem("shore-last-auto-backup-at"), due = !last || Date.now() - new Date(last).getTime() >= 7 * 86400000; if (!due) return;
    try { const now = new Date().toISOString(); localStorage.setItem("shore-auto-backup-snapshot", JSON.stringify(createSiteBackup())); localStorage.setItem("shore-last-auto-backup-at", now); localStorage.setItem("shore-last-backup-at", now); window.setTimeout(() => flash("已自动生成本周全站备份，可在考试与主题中下载"), 1400); } catch { window.setTimeout(() => flash("自动备份空间不足，请手动导出全站数据"), 1400); }
  }, []);

  const flash = (message: string) => { setToast(message); window.setTimeout(() => setToast(""), 2200); };
  useEffect(() => {
    if (sessionStorage.getItem("shore-backup-reminded")) return;
    const saved = localStorage.getItem("shore-last-backup-at"), days = saved ? Math.floor((Date.now() - new Date(saved).getTime()) / 86400000) : null;
    if (days === null || days >= 7) { window.setTimeout(() => flash(days === null ? "还没有备份过数据，建议到考试与主题中创建备份" : `距离上次备份已经 ${days} 天，建议及时备份`), 900); sessionStorage.setItem("shore-backup-reminded", "1"); }
  }, []);
  const toggleTimer = () => {
    if (!timerOn) { setTimerStartedAt(new Date()); setTimerOn(true); return; }
    if (activeSeconds > 12 * 3600 && !window.confirm(`本次学习记录达到 ${formatHours(activeSeconds)}，超过12小时。确认记录无误并保存吗？`)) return;
    const endedAt = new Date();
    if (activeSeconds > 0) setSessions([...sessions, { id: Date.now(), date: localISO(timerStartedAt || endedAt), module: timerModule, title: timerTitle.trim() || timerModule, seconds: activeSeconds, startTime: timerStartedAt ? `${pad(timerStartedAt.getHours())}:${pad(timerStartedAt.getMinutes())}` : undefined, endTime: `${pad(endedAt.getHours())}:${pad(endedAt.getMinutes())}` }]);
    setTimerOn(false); setTimerStartedAt(null); setActiveSeconds(0); flash("本次学习时长已保存");
  };
  const todayTasks = tasks.filter(t => normalizedDate(t.date) === localISO());
  const todayRoutines = routines.filter(r => r.month === localISO().slice(0, 7) && routineApplies(r, localISO()));
  const todayDone = todayTasks.filter(t => t.done).length + todayRoutines.filter(r => r.completedDates.includes(localISO())).length;
  const todayTotal = todayTasks.length + todayRoutines.length;
  const todaySeconds = sessions.filter(s => normalizedDate(s.date) === localISO()).reduce((sum, s) => sum + s.seconds, 0) + activeSeconds;
  const nextTask = todayRoutines.find(r => !r.completedDates.includes(localISO()) && !r.supervision?.[localISO()]?.canceled)?.title || todayTasks.find(t => !t.done && !t.supervision?.canceled)?.title || "";
  const orderedNav = [...navOrder.map(id => nav.find(item => item.id === id)).filter((item): item is typeof nav[number] => Boolean(item)), ...nav.filter(item => !navOrder.includes(item.id))];
  const moveNav = (target: Tab) => {
    if (!dragTab || dragTab === target) return;
    const current = orderedNav.map(item => item.id), from = current.indexOf(dragTab), to = current.indexOf(target);
    current.splice(from, 1); current.splice(to, 0, dragTab); setNavOrder(current); setDragTab(null);
  };

  const sharedTimer = { timerOn, activeSeconds, timerModule, setTimerModule, timerTitle, setTimerTitle, toggleTimer, todaySeconds, timerStartedAt };
  let page;
  if (tab === "home") page = <Dashboard tasks={tasks} setTasks={setTasks} routines={routines} setRoutines={setRoutines} practices={practices} mistakes={mistakes} idioms={idioms} examDates={examDates} setTab={setTab} timer={sharedTimer} />;
  else if (tab === "daily") page = <DailyPlan tasks={tasks} setTasks={setTasks} routines={routines} setRoutines={setRoutines} sessions={sessions} timer={sharedTimer} />;
  else if (tab === "plan") page = <Plan tasks={tasks} setTasks={setTasks} routines={routines} setRoutines={setRoutines} flash={flash} />;
  else if (tab === "study") page = <StudyTime sessions={sessions} setSessions={setSessions} flash={flash} timer={sharedTimer} />;
  else if (tab === "practice") page = <PracticeView practices={practices} setPractices={setPractices} flash={flash} />;
  else if (tab === "idiom") page = <IdiomView idioms={idioms} setIdioms={setIdioms} flash={flash} />;
  else if (tab === "mistakes") page = <MistakeView mistakes={mistakes} setMistakes={setMistakes} flash={flash} />;
  else if (tab === "review") page = <ReviewView mistakes={mistakes} setMistakes={setMistakes} idioms={idioms} setIdioms={setIdioms} />;
  else if (tab === "reports") page = <ReportView tasks={tasks} routines={routines} sessions={sessions} practices={practices} idioms={idioms} />;
  else page = <Settings examDates={examDates} setExamDates={setExamDates} flash={flash} />;

  return <main className="app-shell">
    <aside className={`sidebar ${mobileNav ? "open" : ""}`}>
      <div className="brand"><span className="brand-seal">岸</span><div><strong>上岸手账</strong><small>GONGKAO JOURNAL</small></div></div>
      <nav>{orderedNav.map(item => <button draggable key={item.id} className={`${tab === item.id ? "active" : ""} ${dragTab === item.id ? "dragging" : ""}`} title="按住拖动可调整栏目顺序" onDragStart={() => setDragTab(item.id)} onDragEnd={() => setDragTab(null)} onDragOver={e => e.preventDefault()} onDrop={() => moveNav(item.id)} onClick={() => { setTab(item.id); setMobileNav(false); }}><span>{item.mark}</span><b>{item.label}</b><i className="nav-grip" aria-hidden="true">⋮⋮</i></button>)}</nav>
      <div className="sidebar-foot"><div className="streak"><span>今日任务</span><strong>{todayDone}<small>/{todayTotal}</small></strong></div><p>距离目标，再近一点点。</p></div>
    </aside>
    {mobileNav && <button className="scrim" aria-label="关闭菜单" onClick={() => setMobileNav(false)} />}
    <section className="content"><header className="topbar"><button className="menu-button" onClick={() => setMobileNav(true)}>☰</button><div><p>{new Date().toLocaleDateString("zh-CN", { month: "long", day: "numeric", weekday: "long" })}</p><h1>{orderedNav.find(item => item.id === tab)?.label}</h1></div><button className="top-study-status" onClick={() => setTab("study")}><span>{timerOn ? "专注进行中" : "今日学习"}</span><strong>{timerOn ? timerTitle : formatHours(todaySeconds)}</strong><small>{todayDone}/{todayTotal} 项任务完成</small></button></header>{page}</section>
    {toast && <div className="toast">✓ {toast}</div>}
    <StudySupervisor nextTask={nextTask} timer={sharedTimer} done={todayDone} total={todayTotal} tasks={tasks} setTasks={setTasks} routines={routines} setRoutines={setRoutines} onOpen={() => setTab(nextTask ? "daily" : "study")} />
  </main>;
}

type TimerShared = { timerOn: boolean; activeSeconds: number; timerModule: string; setTimerModule: Dispatch<SetStateAction<string>>; timerTitle: string; setTimerTitle: Dispatch<SetStateAction<string>>; toggleTimer: () => void; todaySeconds: number; timerStartedAt: Date | null };

function Dashboard({ tasks, setTasks, routines, setRoutines, practices, mistakes, idioms, examDates, setTab, timer }: { tasks: Task[]; setTasks: Dispatch<SetStateAction<Task[]>>; routines: DailyRoutine[]; setRoutines: Dispatch<SetStateAction<DailyRoutine[]>>; practices: Practice[]; mistakes: Mistake[]; idioms: Idiom[]; examDates: ExamDates; setTab: Dispatch<SetStateAction<Tab>>; timer: TimerShared }) {
  const todayTasks = tasks.filter(t => normalizedDate(t.date) === localISO());
  const todayRoutines = routines.filter(r => r.month === localISO().slice(0, 7) && routineApplies(r, localISO()));
  const done = todayTasks.filter(t => t.done).length + todayRoutines.filter(r => r.completedDates.includes(localISO())).length;
  const total = todayTasks.length + todayRoutines.length;
  const completion = total ? Math.round(done / total * 100) : 0;
  const q = practices.reduce((s, x) => s + x.total, 0), c = practices.reduce((s, x) => s + x.correct, 0);
  const accuracy = q ? Math.round(c / q * 1000) / 10 : 0;
  const pendingM = mistakes.filter(x => !x.mastered).length, pendingI = idioms.filter(x => x.level !== "已掌握").length;
  const quote = quotes[new Date().getDate() % quotes.length];
  const examItems = [{ label: "国考", date: examDates.national }, { label: "省考", date: examDates.provincial }];
  return <div className="page-stack">
    <section className="hero-card"><div className="hero-copy"><p className="eyebrow">DAILY NOTE · 今日寄语</p><blockquote>“{quote}”</blockquote><p>把今天安排清楚，然后一项一项完成。</p></div><div className="exam-countdowns">{examItems.map(x => <button key={x.label} onClick={() => setTab("settings")}><span>{x.label}</span><strong>{daysUntil(x.date) ?? "—"}</strong><small>{x.date ? "天" : "设置日期"}</small></button>)}</div></section>
    <section className="metric-grid"><Metric label="今日完成" value={`${completion}%`} note={`${done}/${total} 项任务`} color="sage" /><Metric label="今日学习" value={formatHours(timer.todaySeconds)} note={timer.timerOn ? "正在计时" : "点击开始专注"} color="peach" /><Metric label="综合正确率" value={`${accuracy}%`} note={`累计 ${q} 题`} color="lilac" /><Metric label="等待复习" value={pendingM + pendingI} note={`${pendingM} 错题 · ${pendingI} 成语`} color="rose" /></section>
    <DailyBriefing />
    <section className="dashboard-grid"><div className="panel task-panel"><PanelTitle title="今日学习清单" action="打开时间轴" onClick={() => setTab("daily")} /><div className="progress-track"><i style={{ width: `${completion}%` }} /></div><div className="task-list">{total ? <>{todayRoutines.map(routine => { const checked = routine.completedDates.includes(localISO()); return <label className={`task-row ${checked ? "done" : ""}`} key={`r${routine.id}`}><input type="checkbox" checked={checked} onChange={() => setRoutines(routines.map(x => x.id === routine.id ? { ...x, completedDates: checked ? x.completedDates.filter(d => d !== localISO()) : [...x.completedDates, localISO()] } : x))} /><span className="custom-check">✓</span><span className="task-copy"><strong>{routine.title}</strong><small>每日任务 · {routine.subject}</small></span><em>{routine.minutes} 分钟</em></label>; })}{todayTasks.map(task => <label className={`task-row ${task.done ? "done" : ""}`} key={task.id}><input type="checkbox" checked={task.done} onChange={() => setTasks(tasks.map(x => x.id === task.id ? { ...x, done: !x.done } : x))} /><span className="custom-check">✓</span><span className="task-copy"><strong>{task.title}</strong><small>{task.subject}</small></span><em>{task.minutes} 分钟</em></label>)}</> : <Empty text="今天还没有安排任务" />}</div></div>
      <div className="panel timer-panel"><PanelTitle title="专注计时" action="时长统计" onClick={() => setTab("study")} /><div className={`timer-orb ${timer.timerOn ? "running" : ""}`}><small>{timer.timerOn ? "本次专注" : "今日累计"}</small><strong>{formatClock(timer.timerOn ? timer.activeSeconds : timer.todaySeconds)}</strong><span>{timer.timerModule}</span></div><select value={timer.timerModule} onChange={e => timer.setTimerModule(e.target.value)}>{modules.map(m => <option key={m}>{m}</option>)}</select><button className="primary-button wide" onClick={timer.toggleTimer}>{timer.timerOn ? "结束并保存" : "开始学习"}</button></div></section>
  </div>;
}

function DailyBriefing() {
  type SourceKey = "xuexi" | "people" | "gov";
  const sourceMeta: Record<SourceKey, { label: string }> = { xuexi: { label: "学习强国" }, people: { label: "人民日报" }, gov: { label: "中国政府网" } };
  const [source, setSource] = useState<SourceKey>("people"), [allNews, setAllNews] = useState<Record<SourceKey, NewsItem[]>>({ xuexi: [], people: [], gov: [] }), [updatedAt, setUpdatedAt] = useState(""), [feedDate, setFeedDate] = useState(""), [loading, setLoading] = useState(true), [showAnswer, setShowAnswer] = useState<number | null>(null), [refreshKey, setRefreshKey] = useState(0);
  const knowledgeIndex = Math.floor(new Date(`${localISO()}T12:00:00`).getTime() / 86400000);
  const facts = [dailyKnowledge[knowledgeIndex % dailyKnowledge.length], dailyKnowledge[(knowledgeIndex * 3 + 2) % dailyKnowledge.length], dailyKnowledge[(knowledgeIndex * 5 + 4) % dailyKnowledge.length]];
  useEffect(() => {
    setLoading(true);
    fetch(`${new URL("./daily-news.json", window.location.href)}?t=${Date.now()}`).then(response => response.ok ? response.json() : Promise.reject()).then(data => { setFeedDate(data.date || ""); setUpdatedAt(data.updatedAt || ""); setAllNews({ xuexi: data.sources?.xuexi || [], people: data.sources?.people || [], gov: data.sources?.gov || [] }); }).catch(() => { setFeedDate(""); setAllNews({ xuexi: [], people: [], gov: [] }); }).finally(() => setLoading(false));
  }, [refreshKey]);
  const visibleNews = (key: SourceKey) => allNews[key].filter(item => (!item.date || item.date === localISO()) && isDirectNewsArticle(item, key));
  const news = feedDate === localISO() ? visibleNews(source) : [], current = sourceMeta[source];
  return <section className="daily-briefing"><div className="panel news-panel"><div className="briefing-head"><div><p className="eyebrow">DAILY NEWS</p><h2>今日要闻</h2><span>仅展示 {localISO()} 当天发布内容 · {updatedAt ? `最近更新 ${new Date(updatedAt).toLocaleTimeString("zh-CN", { hour: "2-digit", minute: "2-digit" })}` : "等待更新"}</span></div><button onClick={() => setRefreshKey(x => x + 1)}>刷新</button></div><div className="news-source-tabs">{(Object.keys(sourceMeta) as SourceKey[]).map(key => <button className={source === key ? "active" : ""} key={key} onClick={() => setSource(key)}>{sourceMeta[key].label}<small>{feedDate === localISO() ? visibleNews(key).length : 0}</small></button>)}</div>{loading ? <div className="briefing-loading">正在获取今天最新要闻…</div> : news.length ? <div className="news-list">{news.map((item, index) => <a href={item.link} target="_blank" rel="noreferrer" key={`${item.link}-${index}`}><i>{pad(index + 1)}</i><span><strong>{item.title}</strong><small>{item.source} · {item.date || localISO()}</small></span><b>↗</b></a>)}</div> : <div className="news-fallback"><strong>{feedDate !== localISO() ? "今天的数据尚未完成更新" : `${current.label}今天暂未检索到可直接打开的文章`}</strong><p>系统已过滤官网首页、栏目页和“更多”入口，不再用这些页面冒充新闻文章。</p></div>}<div className="news-status"><span>仅保留可直接打开的文章链接</span><span>后台每小时检查更新</span></div></div><div className="panel knowledge-card"><p className="eyebrow">DAILY KNOWLEDGE</p>{facts.map((fact, index) => <article className="knowledge-item" key={index}><span className="knowledge-tag">{fact.category}常识 · 第{index + 1}题</span><h2>{fact.question}</h2><div className={`knowledge-answer ${showAnswer === index ? "show" : ""}`}>{showAnswer === index ? fact.answer : "先在心里作答，再查看答案"}</div><button className={showAnswer === index ? "soft-button" : "primary-button"} onClick={() => setShowAnswer(showAnswer === index ? null : index)}>{showAnswer === index ? "收起答案" : "查看答案"}</button></article>)}<small>每天自动更换三道，覆盖不同类型常识。</small></div></section>;
}

function Metric({ label, value, note, color }: { label: string; value: string | number; note: string; color: string }) { return <article className={`metric-card ${color}`}><p>{label}</p><strong>{value}</strong><small>{note}</small></article>; }
function PanelTitle({ title, action, onClick }: { title: string; action?: string; onClick?: () => void }) { return <div className="panel-title"><h2>{title}</h2>{action && <button onClick={onClick}>{action} →</button>}</div>; }
function Empty({ text }: { text: string }) { return <div className="empty"><span>✓</span><p>{text}</p></div>; }

function StudySupervisor({ nextTask, timer, done, total, tasks, setTasks, routines, setRoutines, onOpen }: { nextTask: string; timer: TimerShared; done: number; total: number; tasks: Task[]; setTasks: Dispatch<SetStateAction<Task[]>>; routines: DailyRoutine[]; setRoutines: Dispatch<SetStateAction<DailyRoutine[]>>; onOpen: () => void }) {
  const [expanded, setExpanded] = useState(false), [awayCount, setAwayCount] = useState(0), [now, setNow] = useState(Date.now()), [reasonMode, setReasonMode] = useState<"delay" | "cancel" | "review" | null>(null), [reason, setReason] = useState("");
  const today = localISO(), yesterday = addDays(-1), nowMinutes = new Date(now).getHours() * 60 + new Date(now).getMinutes();
  type Candidate = { kind: "task" | "routine"; id: number; date: string; title: string; subject: string; start?: string; state: SupervisionState; review?: boolean };
  const todayCandidates: Candidate[] = [
    ...tasks.filter(t => normalizedDate(t.date) === today && !t.done && t.plannedStart).map(t => ({ kind: "task" as const, id: t.id, date: today, title: t.title, subject: t.subject, start: t.plannedStart, state: t.supervision || {} })),
    ...routines.filter(r => r.month === today.slice(0, 7) && routineApplies(r, today) && !r.completedDates.includes(today) && r.plannedStart).map(r => ({ kind: "routine" as const, id: r.id, date: today, title: r.title, subject: r.subject, start: r.plannedStart, state: r.supervision?.[today] || {} })),
  ].filter(x => !x.state.canceled && (timeMinutes(x.start) || 0) <= nowMinutes && (!x.state.postponedUntil || new Date(x.state.postponedUntil).getTime() <= now)).sort((a, b) => (timeMinutes(a.start) || 0) - (timeMinutes(b.start) || 0));
  const reviewCandidates: Candidate[] = [
    ...tasks.filter(t => normalizedDate(t.date) === yesterday && !t.done && !t.supervision?.incompleteReason).map(t => ({ kind: "task" as const, id: t.id, date: yesterday, title: t.title, subject: t.subject, start: t.plannedStart, state: t.supervision || {}, review: true })),
    ...routines.filter(r => r.month === yesterday.slice(0, 7) && routineApplies(r, yesterday) && !r.completedDates.includes(yesterday) && !r.supervision?.[yesterday]?.incompleteReason).map(r => ({ kind: "routine" as const, id: r.id, date: yesterday, title: r.title, subject: r.subject, start: r.plannedStart, state: r.supervision?.[yesterday] || {}, review: true })),
  ];
  const candidate = todayCandidates[0] || reviewCandidates[0];
  const updateState = (item: Candidate, patch: SupervisionState) => { if (item.kind === "task") setTasks(rows => rows.map(t => t.id === item.id ? { ...t, supervision: { ...(t.supervision || {}), ...patch } } : t)); else setRoutines(rows => rows.map(r => r.id === item.id ? { ...r, supervision: { ...(r.supervision || {}), [item.date]: { ...(r.supervision?.[item.date] || {}), ...patch } } } : r)); };
  const todayLost = tasks.filter(t => normalizedDate(t.date) === today).reduce((sum, t) => sum + (t.supervision?.lostMinutes || 0), 0) + routines.reduce((sum, r) => sum + (r.supervision?.[today]?.lostMinutes || 0), 0);
  const strongReminder = (candidate?.state.delayCount || 0) >= 2;
  useEffect(() => { const timerId = window.setInterval(() => setNow(Date.now()), 30000); return () => window.clearInterval(timerId); }, []);
  useEffect(() => { if (candidate && !timer.timerOn) { setExpanded(true); setReasonMode(candidate.review ? "review" : null); if ("Notification" in window && Notification.permission === "granted") new Notification(candidate.review ? "补录未完成原因" : "计划开始提醒", { body: candidate.review ? `昨天的“${candidate.title}”未完成，请补充原因。` : `“${candidate.title}”已经到计划开始时间。` }); } }, [candidate?.kind, candidate?.id, candidate?.date, candidate?.state.postponedUntil, timer.timerOn]);
  useEffect(() => { if (!timer.timerOn) return; setAwayCount(0); const watch = () => { if (document.visibilityState !== "hidden") return; setAwayCount(x => x + 1); if ("Notification" in window && Notification.permission === "granted") new Notification("专注契约提醒", { body: `你正在学习“${timer.timerTitle}”，先完成这一段再离开。` }); }; document.addEventListener("visibilitychange", watch); return () => document.removeEventListener("visibilitychange", watch); }, [timer.timerOn]);
  const beginNow = () => { if (!candidate) return; const overdue = Math.max(0, nowMinutes - (timeMinutes(candidate.start) || nowMinutes)); updateState(candidate, { promptedAt: new Date().toISOString(), postponedUntil: undefined, lostMinutes: Math.max(candidate.state.lostMinutes || 0, overdue) }); timer.setTimerTitle(candidate.title); timer.setTimerModule(candidate.subject); if (!timer.timerOn) timer.toggleTimer(); setExpanded(false); };
  const saveReason = () => { if (!candidate || !reason.trim()) return; if (reasonMode === "delay") { const delayCount = (candidate.state.delayCount || 0) + 1; updateState(candidate, { delayCount, postponedUntil: new Date(Date.now() + 10 * 60000).toISOString(), reasons: [...(candidate.state.reasons || []), reason.trim()], lostMinutes: (candidate.state.lostMinutes || 0) + 10, promptedAt: new Date().toISOString() }); } else { updateState(candidate, { canceled: reasonMode === "cancel" ? true : candidate.state.canceled, incompleteReason: reason.trim(), reasons: [...(candidate.state.reasons || []), reason.trim()], lostMinutes: Math.max(candidate.state.lostMinutes || 0, candidate.start ? Math.max(0, nowMinutes - (timeMinutes(candidate.start) || nowMinutes)) : 0), promptedAt: new Date().toISOString() }); } setReason(""); setReasonMode(null); setExpanded(false); };
  if (!nextTask && !timer.timerOn && !candidate) return null;
  return <aside className={`focus-coach supervisor ${expanded ? "expanded" : ""} ${strongReminder ? "strong-reminder" : ""}`}><button className="coach-orb" onClick={() => setExpanded(!expanded)}>{timer.timerOn ? awayCount : strongReminder ? "!" : "督"}</button>{expanded && <div>{timer.timerOn ? <><small>专注契约进行中</small><strong>{timer.timerTitle}</strong><div className="supervisor-stats"><span><b>{formatClock(timer.activeSeconds)}</b>已专注</span><span><b>{awayCount}</b>离开页面</span><span><b>{done}/{total}</b>任务完成</span></div><div className="delay-loss">今日因拖延损失 <b>{todayLost} 分钟</b></div><span><button className="primary-button" onClick={() => { timer.toggleTimer(); setExpanded(false); }}>结束并保存</button></span></> : candidate ? <>{strongReminder && <div className="strong-alert">已经连续推迟 {candidate.state.delayCount} 次，请不要再把任务留给稍后的自己。</div>}<small>{candidate.review ? "昨日未完成原因补录" : "计划开始提醒"}</small><strong>{candidate.title}</strong><p>{candidate.review ? "这项任务昨天没有完成，请选择或填写真实原因，报告会据此给出调整建议。" : `原计划 ${candidate.start} 开始 · ${candidate.subject}`}</p><div className="delay-loss">今日因拖延损失 <b>{todayLost} 分钟</b></div>{reasonMode ? <div className="reason-form"><div>{["太难", "时间不足", "临时有事", "单纯拖延"].map(x => <button className={reason === x ? "active" : ""} key={x} onClick={() => setReason(x)}>{x}</button>)}</div><input autoFocus value={reason} onChange={e => setReason(e.target.value)} placeholder="必须填写原因，也可以自定义…" /><span><button className="primary-button" disabled={!reason.trim()} onClick={saveReason}>{reasonMode === "delay" ? "确认推迟10分钟" : "保存原因"}</button><button className="soft-button" onClick={() => { setReasonMode(null); setReason(""); }}>返回</button></span></div> : candidate.review ? <button className="primary-button wide" onClick={() => setReasonMode("review")}>填写未完成原因</button> : <div className="supervision-actions"><button className="primary-button" onClick={beginNow}>立即开始</button><button className="soft-button" onClick={() => setReasonMode("delay")}>推迟10分钟</button><button className="soft-button danger" onClick={() => setReasonMode("cancel")}>今日取消</button></div>}</> : <><small>今日执行监督</small><strong>{nextTask}</strong><p>先明确要完成的内容，再开始计时。</p><span><button className="primary-button" onClick={() => { onOpen(); setExpanded(false); }}>设置专注内容</button><button className="soft-button" onClick={() => setExpanded(false)}>稍后再说</button></span></>}{"Notification" in window && Notification.permission === "default" && <button className="coach-notify" onClick={() => Notification.requestPermission()}>开启计划到时与离开页面提醒</button>}</div>}</aside>;
}

function DailyPlan({ tasks, setTasks, routines, setRoutines, sessions, timer }: { tasks: Task[]; setTasks: Dispatch<SetStateAction<Task[]>>; routines: DailyRoutine[]; setRoutines: Dispatch<SetStateAction<DailyRoutine[]>>; sessions: StudySession[]; timer: TimerShared }) {
  const [date, setDate] = useState(localISO());
  const [timelineZoom, setTimelineZoom] = useStoredState<number>("shore-timeline-zoom", 100);
  const [editing, setEditing] = useState<{ kind: "task" | "routine"; id: number } | null>(null);
  const [editTitle, setEditTitle] = useState(""), [editStart, setEditStart] = useState("08:00"), [editEnd, setEditEnd] = useState("09:00");
  const dateTasks = tasks.filter(t => normalizedDate(t.date) === date);
  const dateRoutines = routines.filter(r => r.month === date.slice(0, 7) && routineApplies(r, date));
  const planned = [...dateRoutines.map(r => ({ kind: "routine" as const, id: r.id, title: r.title, subject: r.subject, start: r.plannedStart, end: r.plannedEnd, done: r.completedDates.includes(date), state: r.supervision?.[date] })), ...dateTasks.map(t => ({ kind: "task" as const, id: t.id, title: t.title, subject: t.subject, start: t.plannedStart, end: t.plannedEnd, done: t.done, state: t.supervision }))];
  const actual = sessions.filter(s => normalizedDate(s.date) === date).map(s => ({ ...s, running: false }));
  if (date === localISO() && timer.timerOn && timer.timerStartedAt) actual.push({ id: -1, date, module: timer.timerModule, seconds: timer.activeSeconds, startTime: `${pad(timer.timerStartedAt.getHours())}:${pad(timer.timerStartedAt.getMinutes())}`, endTime: `${pad(new Date().getHours())}:${pad(new Date().getMinutes())}`, running: true });
  const plannedLayout = layoutTimelineEntries(planned.filter(x => x.start), x => x.start, x => x.end);
  const actualLayout = layoutTimelineEntries(actual.filter(x => x.startTime), x => x.startTime, x => x.endTime);
  const laneStyle = (lane: number, laneCount: number) => laneCount > 1 ? { left: `${lane * 100 / laneCount}%`, right: "auto", width: `calc(${100 / laneCount}% - 5px)` } : { left: "0", right: "auto", width: "100%" };
  const beginEdit = (item: typeof planned[number]) => { setEditing({ kind: item.kind, id: item.id }); setEditTitle(item.title); setEditStart(item.start || "08:00"); setEditEnd(item.end || "09:00"); };
  const saveEdit = () => {
    if (!editing) return;
    if (!validTimeRange(editStart, editEnd)) { window.alert("结束时间必须晚于开始时间，请重新设置。"); return; }
    if (editing.kind === "task") setTasks(tasks.map(t => t.id === editing.id ? { ...t, title: editTitle, plannedStart: editStart, plannedEnd: editEnd, minutes: Math.max(5, (timeMinutes(editEnd) || 0) - (timeMinutes(editStart) || 0)) } : t));
    else setRoutines(routines.map(r => r.id === editing.id ? { ...r, title: editTitle, plannedStart: editStart, plannedEnd: editEnd, minutes: Math.max(5, (timeMinutes(editEnd) || 0) - (timeMinutes(editStart) || 0)) } : r));
    setEditing(null);
  };
  const togglePlanned = (item: typeof planned[number]) => {
    if (item.kind === "task") setTasks(tasks.map(t => t.id === item.id ? { ...t, done: !t.done } : t));
    else setRoutines(routines.map(r => r.id === item.id ? { ...r, completedDates: item.done ? r.completedDates.filter(d => d !== date) : [...r.completedDates, date] } : r));
  };
  const timelineTop = (time?: string) => Math.max(0, Math.min(100, (((timeMinutes(time) ?? 360) - 360) / 1080) * 100));
  const timelineHeight = (start?: string, end?: string) => Math.max(4, Math.min(18, (((timeMinutes(end) ?? 420) - (timeMinutes(start) ?? 360)) / 1080) * 100));
  const doneCount = planned.filter(x => x.done).length;
  const lostToday = planned.reduce((sum, x) => sum + (x.state?.lostMinutes || 0), 0);
  return <div className="page-stack"><section className="page-intro daily-intro"><div><p className="eyebrow">DAILY TIMELINE</p><h2>计划时间与真实投入并排看</h2><p>左边是预计安排，右边由学习计时自动生成实际完成时段。</p><span className="sync-badge">✓ 已与月度计划实时同步</span></div><input type="date" value={date} onChange={e => setDate(e.target.value)} /></section>
    <section className="daily-summary"><div><strong>{doneCount}/{planned.length}</strong><span>任务完成</span></div><div><strong>{formatHours(actual.reduce((sum, s) => sum + s.seconds, 0))}</strong><span>实际学习</span></div><div><strong>{planned.reduce((sum, x) => sum + Math.max(0, (timeMinutes(x.end) || 0) - (timeMinutes(x.start) || 0)), 0)}m</strong><span>计划投入</span></div><div className={lostToday ? "delay-summary" : ""}><strong>{lostToday}m</strong><span>拖延损失</span></div></section>
    <section className="panel schedule-dock"><div className="schedule-dock-head"><div><h2>待安排与时间设置</h2><p>未设时间的任务在这里直接安排，不必穿过时间轴来回翻页。</p></div></div>
      {(planned.some(x => !x.start) || actual.some(x => !x.startTime)) ? <div className="unscheduled"><div><strong>未设时间的任务</strong>{planned.filter(x => !x.start).map(item => <button className={editing?.kind === item.kind && editing.id === item.id ? "active" : ""} key={`${item.kind}${item.id}`} onClick={() => beginEdit(item)}>{item.title}<small>点击后就在下方安排时间</small></button>)}</div><div><strong>未记录具体时间的学习</strong>{actual.filter(x => !x.startTime).map(item => <span key={item.id}>{item.module} · {formatHours(item.seconds)}</span>)}</div></div> : <div className="all-scheduled">今天的任务都已设置时间，可点击时间轴中的任务继续修改。</div>}
      {editing && <div className="timeline-editor inline"><input value={editTitle} onChange={e => setEditTitle(e.target.value)} /><label>开始<input type="time" value={editStart} onChange={e => setEditStart(e.target.value)} /></label><label>结束<input type="time" value={editEnd} onChange={e => setEditEnd(e.target.value)} /></label><button className="primary-button" onClick={saveEdit}>保存安排</button><button className="soft-button" onClick={() => setEditing(null)}>取消</button></div>}</section>
    <section className="panel daily-timeline-panel"><div className="timeline-panel-tools"><span>时间轴缩放 <b>{timelineZoom}%</b></span><input type="range" min="65" max="160" step="5" value={timelineZoom} onChange={e => setTimelineZoom(Number(e.target.value))} /><button className="timeline-reset" onClick={() => setTimelineZoom(100)}>↺ 恢复标准</button></div><div className="timeline-head"><span>预计任务安排</span><b>时间</b><span>实际学习记录</span></div><div className="daily-timeline" style={{ height: `${Math.round(1080 * timelineZoom / 100)}px` }}><div className="timeline-side planned-side">{plannedLayout.map(({ item, lane, laneCount }) => <article key={`${item.kind}${item.id}`} className={`${item.done ? "done" : ""} ${item.state?.canceled ? "canceled" : ""} ${(item.state?.delayCount || 0) >= 2 ? "repeated-delay" : ""}`} style={{ top: `${timelineTop(item.start)}%`, minHeight: `${timelineHeight(item.start, item.end)}%`, ...laneStyle(lane, laneCount) }}><button className="timeline-check" onClick={() => togglePlanned(item)}>{item.done ? "✓" : ""}</button><div onClick={() => beginEdit(item)}><strong>{item.title}</strong><small>{item.start}–{item.end} · {item.subject}{item.state?.canceled ? " · 今日取消" : item.state?.delayCount ? ` · 已推迟${item.state.delayCount}次` : ""}</small></div></article>)}</div><div className="timeline-axis">{Array.from({ length: 19 }, (_, i) => i + 6).map(h => <div key={h}><span>{pad(h)}:00</span><i /></div>)}</div><div className="timeline-side actual-side">{actualLayout.map(({ item, lane, laneCount }) => <article key={item.id} className={item.running ? "running" : ""} style={{ top: `${timelineTop(item.startTime)}%`, minHeight: `${timelineHeight(item.startTime, item.endTime)}%`, ...laneStyle(lane, laneCount) }}><div><strong>{item.title || item.module}</strong><small>{item.startTime}–{item.endTime || "进行中"} · {formatHours(item.seconds)}</small></div></article>)}</div></div></section>
  </div>;
}

function Plan({ tasks, setTasks, routines, setRoutines, flash }: { tasks: Task[]; setTasks: Dispatch<SetStateAction<Task[]>>; routines: DailyRoutine[]; setRoutines: Dispatch<SetStateAction<DailyRoutine[]>>; flash: (x: string) => void }) {
  const [month, setMonth] = useState(localISO().slice(0, 7));
  const [date, setDate] = useState(localISO()), [title, setTitle] = useState(""), [subject, setSubject] = useState("资料分析"), [minutes, setMinutes] = useState(45);
  const [plannedStart, setPlannedStart] = useState("08:00"), [plannedEnd, setPlannedEnd] = useState("08:45");
  const [routineTitle, setRoutineTitle] = useState(""), [routineSubject, setRoutineSubject] = useState("资料分析"), [routineMinutes, setRoutineMinutes] = useState(30), [routineStart, setRoutineStart] = useState("07:30"), [routineEnd, setRoutineEnd] = useState("08:00");
  const [importing, setImporting] = useState(false), [importNote, setImportNote] = useState("");
  const [importedPlanFiles, setImportedPlanFiles] = useStoredState<string[]>("shore-imported-plan-files", []);
  const planFileRef = useRef<HTMLInputElement>(null);
  const monthTasks = tasks.filter(t => normalizedDate(t.date).startsWith(month));
  const monthRoutines = routines.filter(r => r.month === month);
  const finished = monthTasks.filter(t => t.done).length;
  const [year, monthNumber] = month.split("-").map(Number);
  const days = new Date(year, monthNumber, 0).getDate();
  const leading = (new Date(year, monthNumber - 1, 1).getDay() + 6) % 7;
  const cells = [...Array(leading).fill(null), ...Array.from({ length: days }, (_, i) => i + 1)];
  const routineTarget = monthRoutines.reduce((sum, r) => sum + Array.from({ length: days }, (_, i) => `${month}-${pad(i + 1)}`).filter(d => routineApplies(r, d)).length, 0);
  const routineFinished = monthRoutines.reduce((sum, r) => sum + r.completedDates.filter(d => d.startsWith(month) && routineApplies(r, d)).length, 0);
  const overallTotal = monthTasks.length + routineTarget, overallFinished = finished + routineFinished, overallPercent = overallTotal ? Math.round(overallFinished / overallTotal * 100) : 0;
  const plannedTotalMinutes = monthTasks.reduce((s, t) => s + t.minutes, 0) + monthRoutines.reduce((sum, r) => sum + r.minutes * Array.from({ length: days }, (_, i) => `${month}-${pad(i + 1)}`).filter(d => routineApplies(r, d)).length, 0);
  const add = (e: FormEvent) => { e.preventDefault(); if (!title.trim()) return; if (!validTimeRange(plannedStart, plannedEnd)) { flash("结束时间必须晚于开始时间"); return; } const duplicate = tasks.some(t => normalizedDate(t.date) === date && t.title.trim() === title.trim()); if (duplicate && !window.confirm("同一天已经存在完全相同的任务，仍要重复添加吗？")) return; setTasks([...tasks, { id: Date.now(), title: title.trim(), subject, minutes, date, done: false, plannedStart, plannedEnd }]); setTitle(""); setMonth(date.slice(0, 7)); flash("任务已加入对应日期"); };
  const addRoutine = (e: FormEvent) => { e.preventDefault(); if (!routineTitle.trim()) return; if (!validTimeRange(routineStart, routineEnd)) { flash("每日任务结束时间必须晚于开始时间"); return; } const duplicate = routines.some(r => r.month === month && r.title.trim() === routineTitle.trim()); if (duplicate && !window.confirm("本月已经存在同名每日任务，仍要重复添加吗？")) return; setRoutines([...routines, { id: Date.now(), title: routineTitle.trim(), subject: routineSubject, minutes: routineMinutes, month, plannedStart: routineStart, plannedEnd: routineEnd, completedDates: [] }]); setRoutineTitle(""); flash("每日任务已加入本月"); };
  const clearMonth = () => { const count = monthTasks.length + monthRoutines.length; if (!count) { flash("本月没有可清除的计划"); return; } if (!window.confirm(`确定清除 ${month} 的全部计划吗？\n将删除 ${monthTasks.length} 条单日任务和 ${monthRoutines.length} 条每日任务，此操作不可撤销。`)) return; setTasks(tasks.filter(t => !normalizedDate(t.date).startsWith(month))); setRoutines(routines.filter(r => r.month !== month)); flash(`已清除 ${month} 的全部计划`); };
  const clearWeek = () => { const anchor = date.startsWith(month) ? date : `${month}-01`, weekStart = startOfWeek(new Date(`${anchor}T12:00:00`)), weekEnd = endOfWeek(new Date(`${anchor}T12:00:00`)), weekDates = datesBetween(weekStart, weekEnd).filter(x => x.startsWith(month)), weekSet = new Set(weekDates), singleCount = tasks.filter(t => weekSet.has(normalizedDate(t.date))).length, routineCount = routines.filter(r => r.month === month).reduce((sum, r) => sum + weekDates.filter(d => routineApplies(r, d)).length, 0); if (!singleCount && !routineCount) { flash("所选周没有可清除的安排"); return; } if (!window.confirm(`确定清除 ${weekStart} 至 ${weekEnd} 的安排吗？\n将删除 ${singleCount} 条单日任务，并暂停 ${routineCount} 次每日任务；本月其他周不受影响。`)) return; setTasks(tasks.filter(t => !weekSet.has(normalizedDate(t.date)))); setRoutines(routines.map(r => r.month !== month ? r : { ...r, skippedDates: Array.from(new Set([...(r.skippedDates || []), ...weekDates.filter(d => routineApplies(r, d))])) })); flash("所选周安排已清除，每日计划已同步"); };
  const importPlan = (e: ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0]; if (!file) return;
    const fingerprint = `${file.name}|${file.size}|${file.lastModified}`; if (importedPlanFiles.includes(fingerprint) && !window.confirm("这个计划表已经导入过，继续可能产生重复任务。是否仍要重新识别？")) { e.target.value = ""; return; }
    setImporting(true); setImportNote("正在识别日期与任务…");
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const result = recognizePlanWorkbook(reader.result as ArrayBuffer);
        const existing = new Set(tasks.map(t => `${normalizedDate(t.date)}|${t.title}`));
        const fresh = result.tasks.filter(t => !existing.has(`${t.date}|${t.title}`)).map((t, i) => ({ ...t, id: Date.now() + i, done: false }));
        const existingRoutines = new Set(routines.map(r => `${r.month}|${r.title}`));
        const freshRoutines = result.routines.filter(r => !existingRoutines.has(`${r.month}|${r.title}`)).map((r, i) => ({ ...r, id: Date.now() + 10000 + i, completedDates: [] }));
        if (!fresh.length && !freshRoutines.length) { setImportNote("没有发现新的任务，可能已经导入过了"); flash("未发现可新增的计划任务"); }
        else {
          setTasks([...tasks, ...fresh]);
          setRoutines([...routines, ...freshRoutines]);
          setImportedPlanFiles(Array.from(new Set([...importedPlanFiles, fingerprint])));
          const firstMonth = fresh[0]?.date?.slice(0, 7) || freshRoutines[0]?.month; if (firstMonth) setMonth(firstMonth);
          setImportNote(`已从“${result.sheet}”识别 ${fresh.length} 条日期任务、${freshRoutines.length} 条每日任务；每日任务已单独归档，不会挤进日期格。`);
          flash(`成功导入 ${fresh.length + freshRoutines.length} 项计划`);
        }
      } catch { setImportNote("识别失败，请确认文件为 .xlsx 或 .xls 格式"); flash("表格识别失败"); }
      setImporting(false); e.target.value = "";
    };
    reader.onerror = () => { setImporting(false); setImportNote("文件读取失败，请重新选择"); };
    reader.readAsArrayBuffer(file);
  };
  return <div className="page-stack">
    <section className="page-intro"><div><p className="eyebrow">MONTHLY PLAN</p><h2>把月目标拆到每一天</h2><p>与每日计划实时同步；点击月历日期可指定“一键清除一周”的所在周。</p></div><div className="intro-actions plan-actions"><input hidden ref={planFileRef} type="file" accept=".xlsx,.xls" onChange={importPlan} /><button className="import-plan-button" onClick={() => planFileRef.current?.click()} disabled={importing}><span>表</span><div><strong>{importing ? "正在识别…" : "一键导入计划表"}</strong><small>自动识别日期与每日任务</small></div></button><input className="month-picker" type="month" value={month} onChange={e => { setMonth(e.target.value); if (!date.startsWith(e.target.value)) setDate(`${e.target.value}-01`); }} /></div></section>
    {importNote && <section className="import-result"><span>✓</span><p>{importNote}</p><button onClick={() => setImportNote("")}>关闭</button></section>}
    <section className="metric-grid compact"><Metric label="本月任务总量" value={overallTotal} note={`${monthTasks.length} 单日 · ${routineTarget} 每日`} color="sage" /><Metric label="本月已完成" value={overallFinished} note={`剩余 ${overallTotal - overallFinished} 项`} color="peach" /><Metric label="完成百分比" value={`${overallPercent}%`} note="含每日任务完成次数" color="lilac" /><Metric label="计划总时长" value={`${Math.round(plannedTotalMinutes / 60 * 10) / 10}h`} note="预计投入" color="rose" /></section>
    <section className="panel routine-panel"><div className="panel-title"><div><h2>本月每日任务</h2><p>这些任务每天都要完成，单独统计，不显示在月历日期格中。</p></div><span>{monthRoutines.length} 项</span></div><div className="routine-grid">{monthRoutines.map(r => { const monthDays = new Date(year, monthNumber, 0).getDate(), doneDays = r.completedDates.filter(d => d.startsWith(month)).length, todayChecked = r.completedDates.includes(localISO()), canCheck = r.month === localISO().slice(0, 7) && routineApplies(r, localISO()); return <article key={r.id}><div><span>{r.plannedStart || "待定"}–{r.plannedEnd || "待定"}</span><button onClick={() => setRoutines(routines.filter(x => x.id !== r.id))}>删除</button></div><strong>{r.title}</strong><small>{r.subject} · 本月 {doneDays}/{monthDays} 天完成</small><div className="progress-track"><i style={{ width: `${doneDays / monthDays * 100}%` }} /></div>{canCheck && <button className={`routine-check-button ${todayChecked ? "checked" : ""}`} onClick={() => setRoutines(routines.map(x => x.id === r.id ? { ...x, completedDates: todayChecked ? x.completedDates.filter(d => d !== localISO()) : [...x.completedDates, localISO()] } : x))}>{todayChecked ? "✓ 今日已打卡" : "今日打卡"}</button>}</article>; })}</div><form className="routine-form" onSubmit={addRoutine}><input value={routineTitle} onChange={e => setRoutineTitle(e.target.value)} placeholder="新增每天都要完成的任务" /><select value={routineSubject} onChange={e => setRoutineSubject(e.target.value)}>{[...modules, "错题复盘", "成语积累", "其他计划"].map(x => <option key={x}>{x}</option>)}</select><input type="time" value={routineStart} onChange={e => setRoutineStart(e.target.value)} /><input type="time" value={routineEnd} onChange={e => setRoutineEnd(e.target.value)} /><input type="number" min="5" value={routineMinutes} onChange={e => setRoutineMinutes(Number(e.target.value))} /><button className="primary-button">添加每日任务</button></form></section>
    <section className="month-layout"><div className="panel calendar-panel"><div className="calendar-week">{["一", "二", "三", "四", "五", "六", "日"].map(x => <span key={x}>周{x}</span>)}</div><div className="calendar-grid">{cells.map((day, i) => day === null ? <div className="calendar-cell blank" key={`b${i}`} /> : (() => { const key = `${month}-${pad(day)}`; const dayTasks = tasks.filter(t => normalizedDate(t.date) === key); return <div className={`calendar-cell ${key === localISO() ? "today" : ""}`} key={key}><button className="day-number" onClick={() => setDate(key)}>{day}</button><div className="day-tasks">{dayTasks.map(task => <div className={`calendar-task ${task.done ? "done" : ""}`} key={task.id}><label><input type="checkbox" checked={task.done} onChange={() => setTasks(tasks.map(x => x.id === task.id ? { ...x, done: !x.done } : x))} /><i /><span>{task.title}</span></label><button title="删除此任务" onClick={() => { setTasks(tasks.filter(x => x.id !== task.id)); flash("任务已删除，每日计划已同步更新"); }}>×</button></div>)}</div></div>; })())}</div></div><div className="month-side"><form className="panel form-card sticky-form" onSubmit={add}><PanelTitle title="添加单日任务" /><label>任务日期<input type="date" value={date} onChange={e => setDate(e.target.value)} /></label><label>任务内容<input value={title} onChange={e => setTitle(e.target.value)} placeholder="例如：完成资料分析第12组" /></label><div className="form-grid"><label>学习模块<select value={subject} onChange={e => setSubject(e.target.value)}>{[...modules, "错题复盘", "成语积累"].map(x => <option key={x}>{x}</option>)}</select></label><label>预计时长（分钟）<input type="number" min="5" step="5" value={minutes} onChange={e => setMinutes(Number(e.target.value))} /></label><label>计划开始<input type="time" value={plannedStart} onChange={e => setPlannedStart(e.target.value)} /></label><label>计划结束<input type="time" value={plannedEnd} onChange={e => setPlannedEnd(e.target.value)} /></label></div><button className="primary-button wide">添加到月计划</button></form><section className="panel clear-plan-panel"><div><strong>批量清除计划</strong><small>所选日期：{date}</small></div><div className="clear-plan-group"><button className="clear-week-button" onClick={clearWeek}>清除所选周</button><button className="clear-month-button" onClick={clearMonth}>清除本月</button></div><p>清除一周时，每日重复任务只暂停该周，不影响本月其他日期。</p></section></div></section>
  </div>;
}

function WhiteNoise() {
  const [playing, setPlaying] = useState(false), [kind, setKind] = useStoredState<"雨声" | "溪流" | "林间风">("shore-noise-kind-v2", "雨声"), [volume, setVolume] = useStoredState<number>("shore-noise-volume", 24);
  const contextRef = useRef<AudioContext | null>(null), sourceRef = useRef<AudioBufferSourceNode | null>(null), gainRef = useRef<GainNode | null>(null);
  const moods = { 雨声: { className: "rain", icon: "☂", note: "深蓝雨幕 · 适合长时间阅读" }, 溪流: { className: "stream", icon: "≈", note: "青绿水波 · 适合刷题复盘" }, 林间风: { className: "forest", icon: "叶", note: "柔和低频 · 适合安静记忆" } } as const;
  const stop = () => { sourceRef.current?.stop(); sourceRef.current?.disconnect(); gainRef.current?.disconnect(); contextRef.current?.close(); sourceRef.current = null; gainRef.current = null; contextRef.current = null; setPlaying(false); };
  const start = () => {
    const AudioCtor = window.AudioContext || (window as typeof window & { webkitAudioContext?: typeof AudioContext }).webkitAudioContext; if (!AudioCtor) return;
    const context = new AudioCtor(), source = context.createBufferSource(), filter = context.createBiquadFilter(), gain = context.createGain(), seconds = 4, buffer = context.createBuffer(1, context.sampleRate * seconds, context.sampleRate), data = buffer.getChannelData(0);
    let smooth = 0; for (let i = 0; i < data.length; i += 1) { const white = Math.random() * 2 - 1; smooth = smooth * .985 + white * .015; data[i] = kind === "雨声" ? white * .24 + smooth * .7 : kind === "溪流" ? white * .12 + smooth * 1.15 : smooth * 1.7; }
    filter.type = kind === "溪流" ? "bandpass" : "lowpass"; filter.frequency.value = kind === "雨声" ? 1450 : kind === "溪流" ? 680 : 340; filter.Q.value = kind === "溪流" ? .45 : .2;
    source.buffer = buffer; source.loop = true; gain.gain.value = volume / 100 * .22; source.connect(filter); filter.connect(gain); gain.connect(context.destination); source.start(); contextRef.current = context; sourceRef.current = source; gainRef.current = gain; setPlaying(true);
  };
  useEffect(() => { if (gainRef.current) gainRef.current.gain.value = volume / 100 * .22; }, [volume]);
  useEffect(() => () => { sourceRef.current?.disconnect(); contextRef.current?.close(); }, []);
  const mood = moods[kind];
  return <section className={`panel white-noise noise-${mood.className} ${playing ? "playing" : ""}`}><div className="noise-scene" aria-hidden="true"><span>{mood.icon}</span><i /><i /><i /></div><div className="noise-copy"><p className="eyebrow">AMBIENT SOUND</p><h2>{kind}</h2><span>{mood.note}</span></div><div className="noise-controls"><div className="noise-kinds">{(Object.keys(moods) as Array<keyof typeof moods>).map(x => <button className={kind === x ? "active" : ""} key={x} onClick={() => { if (playing) stop(); setKind(x); }}>{moods[x].icon}<span>{x}</span></button>)}</div><label>音量 <b>{volume}%</b><input type="range" min="0" max="70" value={volume} onChange={e => setVolume(Number(e.target.value))} /></label><button className="noise-play" onClick={playing ? stop : start}>{playing ? "■ 停止" : "▶ 播放"}</button></div></section>;
}

function StudyTime({ sessions, setSessions, flash, timer }: { sessions: StudySession[]; setSessions: Dispatch<SetStateAction<StudySession[]>>; flash: (x: string) => void; timer: TimerShared }) {
  const [date, setDate] = useState(localISO()), [module, setModule] = useState("资料分析");
  const [startTime, setStartTime] = useState("19:00"), [endTime, setEndTime] = useState("20:00"), [editingId, setEditingId] = useState<number | null>(null);
  const [heatMonth, setHeatMonth] = useState(localISO().slice(0, 7));
  const [selectedHeatDay, setSelectedHeatDay] = useState(localISO());
  const days = Array.from({ length: 7 }, (_, i) => addDays(i - 6));
  const daily = days.map(day => ({ day, seconds: sessions.filter(s => normalizedDate(s.date) === day).reduce((sum, s) => sum + s.seconds, 0) + (day === localISO() ? timer.activeSeconds : 0) }));
  const moduleStats = modules.map(m => ({ module: m, seconds: sessions.filter(s => s.module === m).reduce((sum, s) => sum + s.seconds, 0) + (timer.timerOn && timer.timerModule === m ? timer.activeSeconds : 0) }));
  const maxDay = Math.max(...daily.map(x => x.seconds), 1), maxModule = Math.max(...moduleStats.map(x => x.seconds), 1);
  const [heatYear, heatMonthNumber] = heatMonth.split("-").map(Number);
  const heatDays = new Date(heatYear, heatMonthNumber, 0).getDate();
  const heatLeading = (new Date(heatYear, heatMonthNumber - 1, 1).getDay() + 6) % 7;
  const heatCells = [...Array(heatLeading).fill(null), ...Array.from({ length: heatDays }, (_, i) => i + 1)];
  const heatData = Array.from({ length: heatDays }, (_, i) => {
    const day = `${heatMonth}-${pad(i + 1)}`;
    const seconds = sessions.filter(s => normalizedDate(s.date) === day).reduce((sum, s) => sum + s.seconds, 0) + (day === localISO() ? timer.activeSeconds : 0);
    const level = seconds === 0 ? 0 : seconds < 1800 ? 1 : seconds < 3600 ? 2 : seconds < 7200 ? 3 : seconds < 14400 ? 4 : 5;
    return { day, seconds, level };
  });
  const heatTotal = heatData.reduce((sum, x) => sum + x.seconds, 0), studiedDays = heatData.filter(x => x.seconds > 0).length;
  const selectedSessions = sessions.filter(s => normalizedDate(s.date) === selectedHeatDay);
  const addManual = (e: FormEvent) => { e.preventDefault(); if (!validTimeRange(startTime, endTime)) { flash("学习记录的结束时间必须晚于开始时间"); return; } const finalMinutes = (timeMinutes(endTime) || 0) - (timeMinutes(startTime) || 0); if (finalMinutes > 720 && !window.confirm(`这条学习记录达到 ${Math.round(finalMinutes / 60 * 10) / 10} 小时，超过12小时。确认记录无误吗？`)) return; const row = { date, module, seconds: finalMinutes * 60, startTime, endTime }; if (editingId !== null) { setSessions(sessions.map(s => s.id === editingId ? { ...s, ...row } : s)); setEditingId(null); flash("学习记录已修改"); } else { setSessions([...sessions, { id: Date.now(), ...row }]); flash("学习时长已补录"); } };
  const editSession = (s: StudySession) => { setEditingId(s.id); setDate(normalizedDate(s.date)); setModule(s.module); setStartTime(s.startTime || "19:00"); setEndTime(s.endTime || "20:00"); };
  return <div className="page-stack"><section className="page-intro"><div><p className="eyebrow">STUDY TIMER</p><h2>记录真正投入的时间</h2><p>实时计时与线下学习补录会统一进入统计。</p></div><div className="study-today"><strong>{formatHours(timer.todaySeconds)}</strong><span>今日累计</span></div></section>
    <section className="study-grid"><div className="panel focus-card"><PanelTitle title="当前专注" /><div className={`focus-clock ${timer.timerOn ? "running" : ""}`}><span>{timer.timerOn ? timer.timerTitle : "准备开始"}</span><strong>{formatClock(timer.activeSeconds)}</strong></div><label className="focus-field">专注内容<input value={timer.timerTitle} disabled={timer.timerOn} onChange={e => timer.setTimerTitle(e.target.value)} placeholder="例如：超大杯资料分析第12组" /></label><label className="focus-field">学习分类<select value={timer.timerModule} disabled={timer.timerOn} onChange={e => timer.setTimerModule(e.target.value)}>{modules.map(m => <option key={m}>{m}</option>)}</select></label><button className="primary-button wide" onClick={timer.toggleTimer}>{timer.timerOn ? "结束并计入统计" : "开始计时"}</button></div>
      <div className="panel"><PanelTitle title="近7天学习时长" /><div className="study-bars">{daily.map(x => <div key={x.day}><span>{formatHours(x.seconds)}</span><i style={{ height: `${Math.max(4, x.seconds / maxDay * 125)}px` }} /><small>{formatShortDate(x.day)}</small></div>)}</div></div>
      <div className="panel"><PanelTitle title="模块投入分布" /><div className="horizontal-bars">{moduleStats.map(x => <div key={x.module}><span>{x.module}</span><div><i style={{ width: `${x.seconds / maxModule * 100}%` }} /></div><strong>{Math.round(x.seconds / 360) / 10}h</strong></div>)}</div></div></section>
    <WhiteNoise />
    <section className="panel study-heatmap-panel"><div className="heatmap-head"><div><p className="eyebrow">MONTHLY FOCUS</p><h2>月历学习时长</h2><span>颜色越深，代表当天投入时间越长；点击日期查看当天记录。</span></div><input className="month-picker" type="month" value={heatMonth} onChange={e => setHeatMonth(e.target.value)} /></div><div className="heatmap-summary"><div><strong>{formatHours(heatTotal)}</strong><span>本月累计</span></div><div><strong>{studiedDays}</strong><span>学习天数</span></div><div><strong>{studiedDays ? formatHours(Math.round(heatTotal / studiedDays)) : "0h 0m"}</strong><span>日均时长</span></div></div><div className="heatmap-week">{["一", "二", "三", "四", "五", "六", "日"].map(x => <span key={x}>周{x}</span>)}</div><div className="study-heatmap">{heatCells.map((day, i) => day === null ? <div className="heat-day blank" key={`hb${i}`} /> : (() => { const item = heatData[day - 1]; return <button className={`heat-day level-${item.level} ${item.day === localISO() ? "today" : ""} ${item.day === selectedHeatDay ? "selected" : ""}`} key={item.day} onClick={() => { setDate(item.day); setSelectedHeatDay(item.day); }} title={`${item.day} · ${formatHours(item.seconds)}`}><span>{day}</span><strong>{item.seconds ? formatHours(item.seconds) : "—"}</strong></button>; })())}</div><div className="heat-legend"><span>少</span>{[0, 1, 2, 3, 4, 5].map(x => <i className={`level-${x}`} key={x} />)}<span>多</span><small>0 · &lt;30m · &lt;1h · &lt;2h · &lt;4h · ≥4h</small></div><div className="day-history"><div><h3>{selectedHeatDay} 学习历史</h3><strong>{formatHours(selectedSessions.reduce((sum, x) => sum + x.seconds, 0))}</strong></div>{selectedSessions.length ? selectedSessions.map(item => <article key={item.id}><span>{item.startTime || "未记录"}–{item.endTime || "未记录"}</span><div><strong>{item.title || item.module}</strong><small>{item.module}</small></div><em>{formatHours(item.seconds)}</em></article>) : <p>当天还没有学习记录。</p>}</div></section>
    <section className="two-col"><form className="panel form-card" onSubmit={addManual}><PanelTitle title={editingId === null ? "补录学习时长" : "修改学习记录"} /><div className="form-grid"><label>学习日期<input type="date" value={date} onChange={e => setDate(e.target.value)} /></label><label>学习模块<select value={module} onChange={e => setModule(e.target.value)}>{modules.map(m => <option key={m}>{m}</option>)}</select></label><label>开始时间<input type="time" value={startTime} onChange={e => setStartTime(e.target.value)} /></label><label>结束时间<input type="time" value={endTime} onChange={e => setEndTime(e.target.value)} /></label><label>自动计算时长<input readOnly value={validTimeRange(startTime, endTime) ? `${(timeMinutes(endTime) || 0) - (timeMinutes(startTime) || 0)} 分钟` : "时间范围无效"} /><small>结束时间必须晚于开始时间</small></label></div><button className="primary-button">{editingId === null ? "保存时长" : "保存修改"}</button>{editingId !== null && <button type="button" className="soft-button cancel-edit" onClick={() => setEditingId(null)}>取消修改</button>}</form>
      <div className="panel"><PanelTitle title="最近记录" /><div className="session-list">{sessions.slice().reverse().slice(0, 8).map(s => <div key={s.id}><span>{formatShortDate(s.date)}<small>{s.startTime && `${s.startTime}–${s.endTime || ""}`}</small></span><strong>{s.module}</strong><em>{formatHours(s.seconds)}</em><span className="row-actions"><button onClick={() => editSession(s)}>修改</button><button onClick={() => { setSessions(sessions.filter(x => x.id !== s.id)); if (editingId === s.id) setEditingId(null); }}>删除</button></span></div>)}</div></div></section>
  </div>;
}

function PracticeView({ practices, setPractices, flash }: { practices: Practice[]; setPractices: Dispatch<SetStateAction<Practice[]>>; flash: (x: string) => void }) {
  const [range, setRange] = useState<7 | 30>(7), [date, setDate] = useState(localISO()), [source, setSource] = useState("超大杯"), [module, setModule] = useState("资料分析"), [correct, setCorrect] = useState(18), [total, setTotal] = useState(20), [minutes, setMinutes] = useState(24);
  const [editingId, setEditingId] = useState<number | null>(null);
  const [sourceOptions, setSourceOptions] = useStoredState<string[]>("shore-practice-sources", sources), [moduleOptions, setModuleOptions] = useStoredState<string[]>("shore-practice-modules", modules);
  const [newSource, setNewSource] = useState(""), [newModule, setNewModule] = useState("");
  const allModules = Array.from(new Set([...moduleOptions, ...practices.map(x => x.module)]));
  const renameOption = (kind: "source" | "module", index: number, value: string) => { const clean = value.trim(); if (!clean) return; const list = kind === "source" ? sourceOptions : moduleOptions, old = list[index]; if (list.some((x, i) => i !== index && x === clean)) return; if (kind === "source") { setSourceOptions(list.map((x, i) => i === index ? clean : x)); setPractices(practices.map(x => x.source === old ? { ...x, source: clean } : x)); if (source === old) setSource(clean); } else { setModuleOptions(list.map((x, i) => i === index ? clean : x)); setPractices(practices.map(x => x.module === old ? { ...x, module: clean } : x)); if (module === old) setModule(clean); } };
  const deleteOption = (kind: "source" | "module", index: number) => { const list = kind === "source" ? sourceOptions : moduleOptions; if (list.length <= 1) return; const old = list[index], next = list.filter((_, i) => i !== index); if (kind === "source") { setSourceOptions(next); if (source === old) setSource(next[0]); } else { setModuleOptions(next); if (module === old) setModule(next[0]); } flash(`已删除选项“${old}”，历史记录仍保留`); };
  const addOption = (kind: "source" | "module") => { const value = (kind === "source" ? newSource : newModule).trim(), list = kind === "source" ? sourceOptions : moduleOptions; if (!value || list.includes(value)) return; if (kind === "source") { setSourceOptions([...list, value]); setNewSource(""); setSource(value); } else { setModuleOptions([...list, value]); setNewModule(""); setModule(value); } };
  const sorted = useMemo(() => practices.slice().sort((a, b) => normalizedDate(a.date).localeCompare(normalizedDate(b.date))), [practices]);
  const cutoff = new Date(); cutoff.setDate(cutoff.getDate() - range + 1);
  const previousCutoff = new Date(); previousCutoff.setDate(previousCutoff.getDate() - range * 2 + 1);
  const currentRows = sorted.filter(p => new Date(`${normalizedDate(p.date)}T12:00:00`) >= cutoff);
  const previousRows = sorted.filter(p => { const d = new Date(`${normalizedDate(p.date)}T12:00:00`); return d >= previousCutoff && d < cutoff; });
  const rate = (rows: Practice[]) => { const q = rows.reduce((s, x) => s + x.total, 0), c = rows.reduce((s, x) => s + x.correct, 0); return q ? c / q * 100 : 0; };
  const change = Math.round((rate(currentRows) - rate(previousRows)) * 10) / 10;
  const chartRows = currentRows.slice(-12);
  const points = chartRows.map((p, i) => ({ x: chartRows.length === 1 ? 350 : 50 + i * 600 / (chartRows.length - 1), y: 178 - p.correct / p.total * 138, pct: Math.round(p.correct / p.total * 100), date: formatShortDate(p.date) }));
  const stats = allModules.map(m => { const rows = practices.filter(p => p.module === m); return { module: m, accuracy: Math.round(rate(rows)), count: rows.reduce((s, x) => s + x.total, 0) }; });
  const add = (e: FormEvent) => { e.preventDefault(); if (correct > total) { flash("正确题数不能大于总题数，请检查后再保存"); return; } if (correct < 0 || total <= 0) { flash("题数必须是有效的正数"); return; } const row = { date, source, module, correct, total, minutes }; if (editingId !== null) { setPractices(practices.map(p => p.id === editingId ? { ...p, ...row } : p)); setEditingId(null); flash("练习记录已修改"); } else { setPractices([...practices, { id: Date.now(), ...row }]); flash("练习成绩已记录"); } };
  const editPractice = (p: Practice) => { setEditingId(p.id); setDate(normalizedDate(p.date)); setSource(p.source); setModule(p.module); setCorrect(p.correct); setTotal(p.total); setMinutes(p.minutes); };
  return <div className="page-stack"><section className="page-intro"><div><p className="eyebrow">ACCURACY TRACKER</p><h2>看见每一次进步的轨迹</h2><p>按模块与来源拆分，周变化和月变化统一呈现。</p></div></section>
    <section className="panel option-manager"><div className="option-column"><div><h2>练习来源管理</h2><p>可新增、改名或删除；改名会同步更新历史记录。</p></div><div className="option-list">{sourceOptions.map((x, i) => <div key={`${x}-${i}`}><input defaultValue={x} onBlur={e => renameOption("source", i, e.target.value)} /><button onClick={() => deleteOption("source", i)}>删除</button></div>)}</div><div className="option-add"><input value={newSource} onChange={e => setNewSource(e.target.value)} placeholder="新增来源，如：小程序模考" /><button onClick={() => addOption("source")}>添加</button></div></div><div className="option-column"><div><h2>学习模块管理</h2><p>自定义模块会立即出现在记录表单和正确率统计中。</p></div><div className="option-list">{moduleOptions.map((x, i) => <div key={`${x}-${i}`}><input defaultValue={x} onBlur={e => renameOption("module", i, e.target.value)} /><button onClick={() => deleteOption("module", i)}>删除</button></div>)}</div><div className="option-add"><input value={newModule} onChange={e => setNewModule(e.target.value)} placeholder="新增模块，如：政治理论" /><button onClick={() => addOption("module")}>添加</button></div></div></section>
    <section className="metric-grid compact">{stats.slice(0, 4).map((s, i) => <Metric key={s.module} label={s.module} value={`${s.accuracy}%`} note={`累计 ${s.count} 题`} color={["sage", "peach", "lilac", "rose"][i]} />)}</section>
    <section className="analysis-grid"><div className="panel"><PanelTitle title="各模块正确率" /><div className="horizontal-bars">{stats.map(s => <div key={s.module}><span>{s.module}</span><div><i style={{ width: `${s.accuracy}%` }} /></div><strong>{s.accuracy}%</strong></div>)}</div></div>
      <div className="panel trend-panel"><div className="panel-title"><h2>最近练习变化</h2><div className="range-tabs"><button className={range === 7 ? "active" : ""} onClick={() => setRange(7)}>周变化</button><button className={range === 30 ? "active" : ""} onClick={() => setRange(30)}>月变化</button></div></div><div className={`change-badge ${change < 0 ? "down" : ""}`}>{change >= 0 ? "↑" : "↓"} 较前周期 {Math.abs(change)} 个百分点</div>{points.length ? <svg className="trend-svg" viewBox="0 0 700 225" role="img" aria-label="正确率变化折线图"><line x1="45" y1="40" x2="660" y2="40" /><line x1="45" y1="109" x2="660" y2="109" /><line x1="45" y1="178" x2="660" y2="178" /><line className="axis" x1="45" y1="195" x2="660" y2="195" /><polyline points={points.map(p => `${p.x},${p.y}`).join(" ")} />{points.map((p, i) => <g key={i}><circle cx={p.x} cy={p.y} r="5" /><text className="value" x={p.x} y={p.y - 11}>{p.pct}%</text><text className="date" x={p.x} y="216">{p.date}</text></g>)}</svg> : <Empty text="当前周期还没有练习记录" />}</div></section>
    <section className="two-col practice-bottom"><div className="panel"><PanelTitle title="练习明细" /><div className="data-table editable-table"><div className="table-head"><span>日期</span><span>来源</span><span>模块</span><span>正确率</span><span>操作</span></div>{sorted.slice().reverse().slice(0, 10).map(p => <div key={p.id}><span>{formatShortDate(p.date)}</span><span>{p.source}</span><span>{p.module}</span><strong>{Math.round(p.correct / p.total * 100)}%</strong><span className="row-actions"><button onClick={() => editPractice(p)}>修改</button><button onClick={() => { setPractices(practices.filter(x => x.id !== p.id)); if (editingId === p.id) setEditingId(null); }}>删除</button></span></div>)}</div></div>
      <form className="panel form-card practice-form" onSubmit={add}><PanelTitle title={editingId === null ? "记录一组练习" : "修改练习记录"} /><div className="form-grid"><label className="full-field">练习日期<input type="date" value={date} onChange={e => setDate(e.target.value)} /></label><label>练习来源<select value={source} onChange={e => setSource(e.target.value)}>{sourceOptions.map(x => <option key={x}>{x}</option>)}</select></label><label>学习模块<select value={module} onChange={e => setModule(e.target.value)}>{moduleOptions.map(x => <option key={x}>{x}</option>)}</select></label><label>正确题数<input type="number" min="0" value={correct} onChange={e => setCorrect(Number(e.target.value))} /></label><label>总题数<input type="number" min="1" value={total} onChange={e => setTotal(Number(e.target.value))} /></label><label>实际用时（分钟）<input type="number" min="1" value={minutes} onChange={e => setMinutes(Number(e.target.value))} /></label></div><button className="primary-button wide">{editingId === null ? "保存练习记录" : "保存修改"}</button>{editingId !== null && <button type="button" className="soft-button wide cancel-edit" onClick={() => setEditingId(null)}>取消修改</button>}</form></section>
  </div>;
}

function MeaningText({ item }: { item: Idiom }) {
  const highlights = (item.meaningHighlights || []).filter(Boolean); if (!highlights.length) return <>{item.meaning}</>;
  let parts: ReactNode[] = [item.meaning];
  highlights.forEach((key, keyIndex) => { const next: ReactNode[] = []; parts.forEach((part, partIndex) => { if (typeof part !== "string") { next.push(part); return; } const cleanKey = key.replace(/[，。；、]/g, ""); const at = cleanKey ? part.indexOf(cleanKey) : -1; if (at < 0) next.push(part); else next.push(part.slice(0, at), <mark key={`${keyIndex}-${partIndex}`}>{part.slice(at, at + cleanKey.length)}</mark>, part.slice(at + cleanKey.length)); }); parts = next; });
  return <>{parts}</>;
}

function ExampleText({ item }: { item: Idiom }) { const example = item.example || "待补充"; const at = example.indexOf(item.word); return at < 0 ? <>{example}</> : <>{example.slice(0, at)}<mark className="example-word">{item.word}</mark>{example.slice(at + item.word.length)}</>; }

function IdiomView({ idioms, setIdioms, flash }: { idioms: Idiom[]; setIdioms: Dispatch<SetStateAction<Idiom[]>>; flash: (x: string) => void }) {
  const [query, setQuery] = useState(""), [word, setWord] = useState(""), [meaning, setMeaning] = useState("");
  const [ocrBusy, setOcrBusy] = useState(false), [ocrText, setOcrText] = useState(""), [ocrProgress, setOcrProgress] = useState(0), [preview, setPreview] = useState<Idiom[]>([]);
  const [kindFilter, setKindFilter] = useState<"全部" | "成语" | "词语">("全部"), [levelFilter, setLevelFilter] = useState<"全部" | Idiom["level"]>("全部"), [editItem, setEditItem] = useState<Idiom | null>(null);
  const [highOnly, setHighOnly] = useState(false), [importedOcrFiles, setImportedOcrFiles] = useStoredState<string[]>("shore-imported-ocr-files", []);
  const pdfRef = useRef<HTMLInputElement>(null);
  useEffect(() => { if (!ocrBusy) return; const warn = (event: BeforeUnloadEvent) => { event.preventDefault(); event.returnValue = ""; }; window.addEventListener("beforeunload", warn); return () => window.removeEventListener("beforeunload", warn); }, [ocrBusy]);
  const filtered = idioms.filter(x => `${x.word}${x.meaning}${x.source || ""}`.includes(query) && (kindFilter === "全部" || (x.kind || (x.word.length === 4 ? "成语" : "词语")) === kindFilter) && (levelFilter === "全部" || x.level === levelFilter));
  const rotate = (id: number) => setIdioms(idioms.map(x => { if (x.id !== id) return x; const level = x.level === "未掌握" ? "模糊" : x.level === "模糊" ? "已掌握" : "未掌握"; return { ...x, level, masteredAt: level === "已掌握" ? new Date().toISOString() : undefined }; }));
  const saveIdiomEdit = () => { if (!editItem || !editItem.word.trim() || !editItem.meaning.trim()) return; setIdioms(idioms.map(x => x.id === editItem.id ? { ...editItem, word: editItem.word.trim(), meaning: editItem.meaning.trim() } : x)); setEditItem(null); flash("词条已修改"); };
  const importPdf = async (e: ChangeEvent<HTMLInputElement>) => { const file = e.target.files?.[0]; if (!file) return; const fingerprint = `${file.name}|${file.size}|${file.lastModified}`; if (importedOcrFiles.includes(fingerprint) && !window.confirm("这个 PDF 已经识别过，继续可能产生重复词条。是否仍要重新识别？")) { e.target.value = ""; return; } setOcrBusy(true); setPreview([]); try { const rows = await recognizePeopleDailyPdf(file, (text, value) => { setOcrText(text); setOcrProgress(value); }); const verified = rows.length ? await verifyIdiomsOnline(rows, (text, value) => { setOcrText(text); setOcrProgress(value); }) : rows; setPreview(verified); setImportedOcrFiles(Array.from(new Set([...importedOcrFiles, fingerprint]))); if (!rows.length) setOcrText("没有识别到绿色框内容，请确认版式与示例一致"); } catch (error) { console.error(error); setOcrText("识别失败。请检查网络后重试，首次使用需要加载中文识别模型。"); } finally { setOcrBusy(false); e.target.value = ""; } };
  const confirmPreview = (onlyHigh = false) => { const existing = new Set(idioms.map(x => `${x.source || ""}|${x.word}`)), selected = onlyHigh ? preview.filter(x => x.confidence === "高可信") : preview; const fresh = selected.filter(x => !existing.has(`${x.source || ""}|${x.word}`)).map((x, i) => { const { ocrCrop: _ocrCrop, exampleCrop: _exampleCrop, candidates: _candidates, uncertainPositions: _uncertain, rawWord: _raw, ...clean } = x; return { ...clean, id: Date.now() + i, createdAt: new Date().toISOString() }; }); setIdioms([...idioms, ...fresh]); setPreview([]); flash(`已导入 ${fresh.length} 条人民日报成语`); };
  const cancelPreview = () => { setPreview([]); setOcrText(""); setOcrProgress(0); flash("已取消导入并清除本次识别结果"); };
  const highCount = preview.filter(x => x.confidence === "高可信").length, visiblePreview = highOnly ? preview.filter(x => x.confidence === "高可信") : preview;
  return <div className="page-stack">{ocrBusy && <div className="ocr-lock" role="alert"><div><span>识</span><h2>正在识别 PDF</h2><p>{ocrText || "正在准备中文识别模型…"}</p><div><i style={{ width: `${ocrProgress}%` }} /></div><strong>{ocrProgress}%</strong><small>请保持当前页面打开，不要切换栏目、刷新或关闭窗口，否则本次识别可能中断。</small></div></div>}<section className="page-intro"><div><p className="eyebrow">IDIOM NOTEBOOK</p><h2>成语不是背过，是能辨清</h2><p>支持识别人民日报图片型 PDF：绿色框选成语、左侧例句与蓝色重点词义。</p></div><div className="stat-pills"><span>总数 <b>{idioms.length}</b></span><span>未掌握 <b>{idioms.filter(x => x.level !== "已掌握").length}</b></span></div></section>
    <section className="panel pdf-import-card"><div><p className="eyebrow">PEOPLE'S DAILY OCR</p><h2>导入人民日报笔记</h2><p>文件只在当前浏览器本地识别。首次使用需加载中文模型，四页通常需要等待一段时间。</p></div><input hidden ref={pdfRef} type="file" accept="application/pdf,.pdf" onChange={importPdf} /><button className="primary-button" disabled={ocrBusy} onClick={() => pdfRef.current?.click()}>{ocrBusy ? "正在识别…" : "选择 PDF 并识别"}</button>{(ocrBusy || ocrText) && <div className="ocr-progress"><div><i style={{ width: `${ocrProgress}%` }} /></div><span>{ocrText}</span><strong>{ocrProgress}%</strong></div>}</section>
    {preview.length > 0 && <section className="panel ocr-preview second-review"><div className="panel-title"><div><h2>识别结果二次校对</h2><p>原图、例句、词义和词库候选同时展示；橙色字符代表识别不确定。</p></div><div className="ocr-preview-actions"><button className="soft-button" onClick={cancelPreview}>取消导入并清除</button><button className="soft-button" disabled={!highCount} onClick={() => confirmPreview(true)}>仅导入高可信 {highCount} 条</button><button className="primary-button" onClick={() => confirmPreview(false)}>确认导入全部 {preview.length} 条</button></div></div><div className="ocr-confidence-toolbar"><div>{(["高可信", "待确认", "疑似错误"] as OcrConfidence[]).map(level => <span className={`confidence-badge ${level}`} key={level}>{level} {preview.filter(x => x.confidence === level).length}</span>)}</div><label><input type="checkbox" checked={highOnly} onChange={e => setHighOnly(e.target.checked)} />只显示高可信</label></div><div className="ocr-review-list">{visiblePreview.map(item => { const index = preview.findIndex(x => x.id === item.id); return <article className={`ocr-review-card confidence-${item.confidence || "待确认"}`} key={item.id}><div className="ocr-proof" id={`ocr-proof-${item.id}`}><div><span>右侧词义原图</span>{item.ocrCrop ? <img src={item.ocrCrop} alt={`${item.word}词义原图`} /> : <p>未保留到裁剪图</p>}</div><div><span>左侧例句原图</span>{item.exampleCrop ? <img src={item.exampleCrop} alt={`${item.word}例句原图`} /> : <p>未保留到裁剪图</p>}</div></div><div className="ocr-fields"><div className="ocr-card-head"><span className={`confidence-badge ${item.confidence || "待确认"}`}>{item.confidence || "待确认"}</span><button onClick={() => document.getElementById(`ocr-proof-${item.id}`)?.scrollIntoView({ behavior: "smooth", block: "center" })}>定位原文</button><button className="remove" onClick={() => setPreview(preview.filter(x => x.id !== item.id))}>移除此条</button></div><div className="ocr-word-proof">{Array.from(item.word).map((char, charIndex) => <mark className={(item.uncertainPositions || []).includes(charIndex) ? "uncertain" : ""} key={charIndex}>{char}</mark>)}</div>{item.candidates?.length ? <div className="candidate-row"><span>候选：</span>{item.candidates.map(candidate => <button className={candidate === item.word ? "active" : ""} key={candidate} onClick={() => setPreview(preview.map((x, j) => j === index ? { ...x, word: candidate, uncertainPositions: [], confidence: commonIdioms.includes(candidate) ? "高可信" : "待确认", ocrNote: `已人工选择候选：${candidate}` } : x))}>{candidate}</button>)}</div> : <div className="candidate-row empty-candidate">暂无可靠候选，请对照原图手工修改</div>}<div className={`ocr-check-note ${item.confidence === "疑似错误" ? "warning" : "ok"}`}>{item.ocrNote || "等待校验"}</div><label>来源<input value={item.source || ""} onChange={e => setPreview(preview.map((x, j) => j === index ? { ...x, source: e.target.value } : x))} /></label><label>识别词条<input className="idiom-word-input" value={item.word} onChange={e => setPreview(preview.map((x, j) => j === index ? { ...x, word: e.target.value, uncertainPositions: [], confidence: "待确认", ocrNote: "已人工修改" } : x))} /></label><label>词义（与右侧绿色框对照）<textarea value={item.meaning} onChange={e => setPreview(preview.map((x, j) => j === index ? { ...x, meaning: e.target.value } : x))} /></label><label>重点高亮<input value={(item.meaningHighlights || []).join("；")} onChange={e => setPreview(preview.map((x, j) => j === index ? { ...x, meaningHighlights: e.target.value.split("；").filter(Boolean) } : x))} /></label><label>例句（与左侧原图对照）<textarea value={item.example} onChange={e => setPreview(preview.map((x, j) => j === index ? { ...x, example: e.target.value } : x))} /></label></div></article>; })}</div></section>}
    {editItem && <section className="panel idiom-edit-panel"><div className="panel-title"><h2>编辑已导入词条</h2><button onClick={() => setEditItem(null)}>取消</button></div><div className="idiom-edit-grid"><label>类型<select value={editItem.kind || (editItem.word.length === 4 ? "成语" : "词语")} onChange={e => setEditItem({ ...editItem, kind: e.target.value as "成语" | "词语" })}><option>成语</option><option>词语</option></select></label><label>词条<input value={editItem.word} onChange={e => setEditItem({ ...editItem, word: e.target.value })} /></label><label>来源<input value={editItem.source || ""} onChange={e => setEditItem({ ...editItem, source: e.target.value })} /></label><label className="wide-field">词义<textarea value={editItem.meaning} onChange={e => setEditItem({ ...editItem, meaning: e.target.value })} /></label><label className="wide-field">例句<textarea value={editItem.example} onChange={e => setEditItem({ ...editItem, example: e.target.value })} /></label><label className="wide-field">重点高亮<input value={(editItem.meaningHighlights || []).join("；")} onChange={e => setEditItem({ ...editItem, meaningHighlights: e.target.value.split("；").filter(Boolean) })} /></label></div><button className="primary-button" onClick={saveIdiomEdit}>保存修改</button></section>}
    <section className="panel"><div className="toolbar idiom-toolbar"><input className="search" value={query} onChange={e => setQuery(e.target.value)} placeholder="搜索成语、词语、词义或来源…" /><div className="idiom-filters">{["全部", "成语", "词语"].map(x => <button className={kindFilter === x ? "active" : ""} key={x} onClick={() => setKindFilter(x as typeof kindFilter)}>{x}</button>)}</div><select value={levelFilter} onChange={e => setLevelFilter(e.target.value as typeof levelFilter)}><option>全部</option><option>未掌握</option><option>模糊</option><option>已掌握</option></select><button className="soft-button" onClick={() => { const blob = new Blob(["来源,成语,词义,重点,例句,掌握状态\n" + idioms.map(x => `${x.source || ""},${x.word},${x.meaning},${(x.meaningHighlights || []).join("；")},${x.example},${x.level}`).join("\n")], { type: "text/csv" }); const a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = "成语积累.csv"; a.click(); }}>导出表格</button></div><div className="idiom-grid">{filtered.map(x => <article className="idiom-card" key={x.id}><div><span><i className="word-kind">{x.kind || (x.word.length === 4 ? "成语" : "词语")}</i><i className={`level ${x.level}`}>{x.level}</i></span><span className="idiom-card-actions"><button onClick={() => rotate(x.id)}>切换状态</button><button onClick={() => setEditItem({ ...x, kind: x.kind || (x.word.length === 4 ? "成语" : "词语") })}>编辑</button><button onClick={() => { setIdioms(idioms.filter(y => y.id !== x.id)); if (editItem?.id === x.id) setEditItem(null); }}>删除</button></span></div>{x.source && <b className="idiom-source">{x.source}</b>}<h3>{x.word}</h3><p><MeaningText item={x} /></p><small>例：<ExampleText item={x} /></small></article>)}</div></section>
    <form className="panel inline-form" onSubmit={e => { e.preventDefault(); if (!word || !meaning) return; setIdioms([...idioms, { id: Date.now(), word, meaning, example: "", level: "未掌握", kind: word.length === 4 ? "成语" : "词语", createdAt: new Date().toISOString() }]); setWord(""); setMeaning(""); flash("词条已加入积累本"); }}><h3>快速添加</h3><input value={word} onChange={e => setWord(e.target.value)} placeholder="成语或词语" /><input value={meaning} onChange={e => setMeaning(e.target.value)} placeholder="词义或易错点" /><button className="primary-button">添加</button></form></div>;
}

function MistakeView({ mistakes, setMistakes, flash }: { mistakes: Mistake[]; setMistakes: Dispatch<SetStateAction<Mistake[]>>; flash: (x: string) => void }) {
  const [filter, setFilter] = useState("全部"); const fileRef = useRef<HTMLInputElement>(null), imageRef = useRef<HTMLInputElement>(null);
  const importCsv = (e: ChangeEvent<HTMLInputElement>) => { const file = e.target.files?.[0]; if (!file) return; const reader = new FileReader(); reader.onload = () => { const rows: Mistake[] = String(reader.result).split(/\r?\n/).slice(1).filter(Boolean).map((line, i) => { const [module = "未分类", source = "批量导入", title = "导入错题", mine = "", answer = "", reason = "待复盘"] = line.split(","); return { id: Date.now() + i, module, source, title, mine, answer, reason, mastered: false }; }); setMistakes([...mistakes, ...rows]); flash(`成功导入 ${rows.length} 道错题`); }; reader.readAsText(file); e.target.value = ""; };
  const importImage = (e: ChangeEvent<HTMLInputElement>) => { const file = e.target.files?.[0]; if (!file) return; const reader = new FileReader(); reader.onload = () => { setMistakes([{ id: Date.now(), module: "待分类", source: "图片导入", title: file.name.replace(/\.[^.]+$/, ""), answer: "", mine: "", reason: "待复盘", image: String(reader.result), mastered: false }, ...mistakes]); flash("图片错题已导入"); }; reader.readAsDataURL(file); e.target.value = ""; };
  const shown = filter === "全部" ? mistakes : mistakes.filter(x => x.module === filter);
  return <div className="page-stack"><section className="page-intro"><div><p className="eyebrow">MISTAKE LIBRARY</p><h2>错题的终点，是不再错</h2><p>支持图片与 CSV 批量导入，按模块、来源和错因整理。</p></div><div className="import-actions"><input hidden ref={fileRef} type="file" accept=".csv" onChange={importCsv} /><input hidden ref={imageRef} type="file" accept="image/*" onChange={importImage} /><button className="soft-button" onClick={() => fileRef.current?.click()}>导入 CSV</button><button className="primary-button" onClick={() => imageRef.current?.click()}>导入题目图片</button></div></section>
    <section className="panel"><div className="toolbar"><div className="filter-tabs">{["全部", ...modules].map(x => <button key={x} className={filter === x ? "active" : ""} onClick={() => setFilter(x)}>{x}</button>)}</div><span>{shown.length} 道</span></div><div className="mistake-list">{shown.map(x => <article key={x.id} className={x.mastered ? "mastered" : ""}>{x.image && <img src={x.image} alt="导入的错题" />}<div className="mistake-main"><p><span>{x.module}</span><span>{x.source}</span></p><h3>{x.title}</h3><small>我的答案：{x.mine || "待补充"}　正确答案：{x.answer || "待补充"}　·　{x.reason}</small></div><div className="mistake-actions"><button className="mastery-button" onClick={() => setMistakes(mistakes.map(y => y.id === x.id ? { ...y, mastered: !y.mastered } : y))}>{x.mastered ? "已掌握" : "标记掌握"}</button><button className="mistake-delete-button" title="删除这道错题" aria-label="删除这道错题" onClick={() => { if (window.confirm("确定删除这道错题吗？")) { setMistakes(mistakes.filter(y => y.id !== x.id)); flash("错题已删除"); } }}><span aria-hidden="true">×</span>删除</button></div></article>)}</div></section><div className="notice-card"><strong>CSV 导入格式</strong><p>首行为表头，字段顺序：模块、来源、题目、我的答案、正确答案、错因。</p></div></div>;
}

function ReviewView({ mistakes, setMistakes, idioms, setIdioms }: { mistakes: Mistake[]; setMistakes: Dispatch<SetStateAction<Mistake[]>>; idioms: Idiom[]; setIdioms: Dispatch<SetStateAction<Idiom[]>> }) {
  const pendingM = mistakes.filter(x => !x.mastered), pendingI = idioms.filter(x => x.level !== "已掌握");
  return <div className="page-stack"><section className="page-intro"><div><p className="eyebrow">DAILY REVIEW</p><h2>今日复习队列</h2><p>先清错题，再过成语。少量、多次、稳定重复。</p></div><div className="review-count"><strong>{pendingM.length + pendingI.length}</strong><span>项待复习</span></div></section><section className="two-col"><div className="panel"><PanelTitle title={`错题复盘 · ${pendingM.length}`} />{pendingM.length ? pendingM.slice(0, 6).map(x => <div className="review-row" key={x.id}><span>{x.module.slice(0, 1)}</span><div><strong>{x.title}</strong><small>{x.reason}</small></div><button onClick={() => setMistakes(mistakes.map(y => y.id === x.id ? { ...y, mastered: true } : y))}>掌握</button></div>) : <Empty text="今日错题已清空" />}</div><div className="panel"><PanelTitle title={`成语复习 · ${pendingI.length}`} />{pendingI.length ? pendingI.slice(0, 6).map(x => <div className="review-row idiom" key={x.id}><span>{x.word.slice(0, 1)}</span><div><strong>{x.word}</strong><small>{x.meaning}</small></div><button onClick={() => setIdioms(idioms.map(y => y.id === x.id ? { ...y, level: "已掌握", masteredAt: new Date().toISOString() } : y))}>记住了</button></div>) : <Empty text="今日成语已清空" />}</div></section></div>;
}

const datesBetween = (start: string, end: string) => { const rows: string[] = [], cursor = new Date(`${start}T12:00:00`), last = new Date(`${end}T12:00:00`); while (cursor <= last) { rows.push(localISO(cursor)); cursor.setDate(cursor.getDate() + 1); } return rows; };
const startOfWeek = (date = new Date()) => { const d = new Date(date), day = (d.getDay() + 6) % 7; d.setDate(d.getDate() - day); return localISO(d); };
const endOfWeek = (date = new Date()) => { const d = new Date(`${startOfWeek(date)}T12:00:00`); d.setDate(d.getDate() + 6); return localISO(d); };

function ReportView({ tasks, routines, sessions, practices, idioms }: { tasks: Task[]; routines: DailyRoutine[]; sessions: StudySession[]; practices: Practice[]; idioms: Idiom[] }) {
  const [period, setPeriod] = useState<"week" | "month">("week"), today = new Date(), currentMonth = localISO(today).slice(0, 7);
  const start = period === "week" ? startOfWeek(today) : `${currentMonth}-01`, end = period === "week" ? endOfWeek(today) : `${currentMonth}-${pad(new Date(today.getFullYear(), today.getMonth() + 1, 0).getDate())}`, rangeDates = datesBetween(start, end), rangeSet = new Set(rangeDates);
  const duration = rangeDates.length, previousEndDate = new Date(`${start}T12:00:00`); previousEndDate.setDate(previousEndDate.getDate() - 1); const previousStartDate = new Date(previousEndDate); previousStartDate.setDate(previousStartDate.getDate() - duration + 1); const previousStart = localISO(previousStartDate), previousEnd = localISO(previousEndDate);
  const planned: Array<{ title: string; subject: string; minutes: number; done: boolean; start?: string; state?: SupervisionState }> = [];
  tasks.filter(t => rangeSet.has(normalizedDate(t.date))).forEach(t => planned.push({ title: t.title, subject: t.subject, minutes: t.minutes, done: t.done, start: t.plannedStart, state: t.supervision }));
  routines.forEach(r => rangeDates.filter(date => date.startsWith(r.month) && routineApplies(r, date)).forEach(date => planned.push({ title: r.title, subject: r.subject, minutes: r.minutes, done: r.completedDates.includes(date), start: r.plannedStart, state: r.supervision?.[date] })));
  const completed = planned.filter(x => x.done).length, completion = planned.length ? Math.round(completed / planned.length * 100) : 0, plannedMinutes = planned.reduce((sum, x) => sum + x.minutes, 0), periodSessions = sessions.filter(s => rangeSet.has(normalizedDate(s.date))), actualMinutes = Math.round(periodSessions.reduce((sum, x) => sum + x.seconds, 0) / 60), gap = actualMinutes - plannedMinutes;
  const moduleTime = Array.from(new Set([...modules, ...periodSessions.map(s => s.module)])).map(module => ({ module, minutes: Math.round(periodSessions.filter(s => s.module === module).reduce((sum, s) => sum + s.seconds, 0) / 60) })).filter(x => x.minutes > 0).sort((a, b) => b.minutes - a.minutes);
  const accuracy = (rows: Practice[]) => { const total = rows.reduce((sum, x) => sum + x.total, 0); return total ? Math.round(rows.reduce((sum, x) => sum + x.correct, 0) / total * 1000) / 10 : 0; };
  const currentPractice = practices.filter(p => normalizedDate(p.date) >= start && normalizedDate(p.date) <= end), previousPractice = practices.filter(p => normalizedDate(p.date) >= previousStart && normalizedDate(p.date) <= previousEnd), practiceModules = Array.from(new Set([...currentPractice, ...previousPractice].map(x => x.module)));
  const accuracyRows = practiceModules.map(module => { const current = accuracy(currentPractice.filter(x => x.module === module)), previous = accuracy(previousPractice.filter(x => x.module === module)); return { module, current, change: Math.round((current - previous) * 10) / 10 }; });
  const delayed = planned.filter(x => (x.state?.delayCount || 0) > 0 || x.state?.canceled), lostMinutes = planned.reduce((sum, x) => sum + (x.state?.lostMinutes || 0), 0), timeGroups = delayed.reduce<Record<string, number>>((map, x) => { const hour = Number(x.start?.slice(0, 2) || 0), key = hour < 9 ? "早晨 6–9点" : hour < 12 ? "上午 9–12点" : hour < 14 ? "午间 12–14点" : hour < 18 ? "下午 14–18点" : "晚间 18点后"; map[key] = (map[key] || 0) + 1; return map; }, {}), worstTime = Object.entries(timeGroups).sort((a, b) => b[1] - a[1])[0];
  const unfinishedGroups = planned.filter(x => !x.done).reduce<Record<string, number>>((map, x) => { map[x.subject] = (map[x.subject] || 0) + 1; return map; }, {}), worstSubject = Object.entries(unfinishedGroups).sort((a, b) => b[1] - a[1])[0];
  const reasonGroups = planned.reduce<Record<string, number>>((map, x) => { const key = x.state?.incompleteReason || x.state?.reasons?.[x.state.reasons.length - 1]; if (key) map[key] = (map[key] || 0) + 1; return map; }, {}), topReason = Object.entries(reasonGroups).sort((a, b) => b[1] - a[1])[0];
  const createdIdioms = idioms.filter(x => x.createdAt && normalizedDate(x.createdAt.slice(0, 10)) >= start && normalizedDate(x.createdAt.slice(0, 10)) <= end).length, masteredIdioms = idioms.filter(x => x.masteredAt && normalizedDate(x.masteredAt.slice(0, 10)) >= start && normalizedDate(x.masteredAt.slice(0, 10)) <= end).length;
  const reasonAdvice = !topReason ? "继续如实记录推迟与未完成原因，报告会逐渐形成更准确的建议。" : topReason[0] === "太难" ? "“太难”是主要原因：把同类任务拆成20–30分钟的小步骤，并先安排基础题。" : topReason[0] === "时间不足" ? "“时间不足”是主要原因：减少每日任务数量，并给高频超时任务预留20%缓冲。" : topReason[0] === "临时有事" ? "“临时有事”是主要原因：每天保留一个机动时段，不要把计划排满。" : topReason[0] === "单纯拖延" ? "“单纯拖延”是主要原因：把最重要任务放在第一个学习时段，到点直接启动10分钟专注。" : `主要记录原因为“${topReason[0]}”，下个周期安排时应提前规避这一因素。`;
  const suggestions = [completion < 70 ? "下个周期先减少并行任务，把每日核心任务控制在3项以内。" : "计划完成率较稳定，可以保持当前任务密度。", gap < -Math.max(60, plannedMinutes * .2) ? `实际投入比计划少 ${Math.abs(gap)} 分钟，建议减少空泛安排并为任务设置明确开始时间。` : gap > Math.max(60, plannedMinutes * .2) ? `实际投入比计划多 ${gap} 分钟，后续应适当上调同类任务预计时长。` : "计划时长与实际投入接近，当前估时较合理。", worstTime ? `${worstTime[0]}最容易发生推迟，下个周期可把最难任务移出该时段。` : "暂未记录明显拖延时段，继续按计划开始计时。", worstSubject ? `${worstSubject[0]}未完成最多，建议拆小任务或降低单次任务量。` : "本周期任务均已完成。", reasonAdvice];
  return <div className="page-stack"><section className="page-intro report-intro"><div><p className="eyebrow">STUDY REPORT</p><h2>自动学习复盘</h2><p>{start} 至 {end} · 根据计划、计时、练习和成语记录自动生成</p></div><div className="report-tabs"><button className={period === "week" ? "active" : ""} onClick={() => setPeriod("week")}>本周周报</button><button className={period === "month" ? "active" : ""} onClick={() => setPeriod("month")}>本月月报</button></div></section><section className="metric-grid"><Metric label="计划完成率" value={`${completion}%`} note={`${completed}/${planned.length} 项完成`} color="sage" /><Metric label="计划 / 实际" value={`${Math.round(plannedMinutes / 60 * 10) / 10}h`} note={`实际 ${Math.round(actualMinutes / 6) / 10}h · ${gap >= 0 ? "+" : ""}${gap}m`} color="peach" /><Metric label="拖延损失" value={`${lostMinutes}m`} note={`${delayed.length} 项发生推迟或取消`} color="rose" /><Metric label="成语变化" value={`+${createdIdioms}`} note={`本期掌握 ${masteredIdioms} 条`} color="lilac" /></section><section className="report-grid"><div className="panel"><PanelTitle title="各模块投入时间" /><div className="report-bars">{moduleTime.length ? moduleTime.map(x => <div key={x.module}><span>{x.module}</span><div><i style={{ width: `${x.minutes / Math.max(...moduleTime.map(y => y.minutes), 1) * 100}%` }} /></div><strong>{Math.round(x.minutes / 6) / 10}h</strong></div>) : <Empty text="本周期还没有学习计时" />}</div></div><div className="panel"><PanelTitle title="各模块正确率变化" /><div className="accuracy-change-list">{accuracyRows.length ? accuracyRows.map(x => <div key={x.module}><span>{x.module}</span><strong>{x.current}%</strong><em className={x.change < 0 ? "down" : ""}>{x.change >= 0 ? "+" : ""}{x.change}%</em></div>) : <Empty text="本周期还没有练习记录" />}</div></div><div className="panel procrastination-card"><PanelTitle title="拖延与未完成分析" /><div className="report-insights"><div><span>最容易拖延的时间</span><strong>{worstTime ? worstTime[0] : "暂无明显时段"}</strong><small>{worstTime ? `${worstTime[1]} 次记录` : "完成更多监督记录后生成"}</small></div><div><span>经常未完成的任务类型</span><strong>{worstSubject ? worstSubject[0] : "暂无"}</strong><small>{worstSubject ? `${worstSubject[1]} 项未完成` : "本周期执行良好"}</small></div><div><span>主要推迟或未完成原因</span><strong>{topReason ? topReason[0] : "暂无"}</strong><small>{topReason ? `${topReason[1]} 次记录` : "请在监督提醒中如实选择"}</small></div></div></div><div className="panel suggestions-card"><PanelTitle title={period === "week" ? "下周调整建议" : "下月调整建议"} /><ol>{suggestions.map((x, i) => <li key={i}>{x}</li>)}</ol></div></section></div>;
}

function BackupCenter({ flash }: { flash: (x: string) => void }) {
  const [weekly, setWeekly] = useStoredState<boolean>("shore-weekly-backup-enabled", false), [lastAt, setLastAt] = useState(localStorage.getItem("shore-last-backup-at") || "");
  const importRef = useRef<HTMLInputElement>(null), autoSnapshot = localStorage.getItem("shore-auto-backup-snapshot"), days = lastAt ? Math.max(0, Math.floor((Date.now() - new Date(lastAt).getTime()) / 86400000)) : null;
  const exportAll = () => { const backup = createSiteBackup(), now = new Date().toISOString(); downloadBackupFile(backup); localStorage.setItem("shore-last-backup-at", now); setLastAt(now); flash("全站数据备份已导出"); };
  const downloadAuto = () => { if (!autoSnapshot) return; try { downloadBackupFile(JSON.parse(autoSnapshot), "每周自动备份"); flash("自动备份已下载，可传到手机或其他电脑恢复"); } catch { flash("自动备份文件已损坏，请重新手动导出"); } };
  const restore = (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0]; if (!file) return; const reader = new FileReader();
    reader.onload = () => { try { const backup = JSON.parse(String(reader.result)) as SiteBackup; if (backup.app !== "上岸手账" || !backup.storage || typeof backup.storage !== "object") throw new Error("invalid"); if (!window.confirm(`备份时间：${new Date(backup.exportedAt).toLocaleString("zh-CN")}\n恢复后将覆盖当前浏览器中的任务、记录和设置，确定继续吗？`)) return; Object.keys(localStorage).filter(key => key.startsWith("shore-")).forEach(key => localStorage.removeItem(key)); Object.entries(backup.storage).forEach(([key, value]) => { if (key.startsWith("shore-") && typeof value === "string") localStorage.setItem(key, value); }); localStorage.setItem("shore-last-backup-at", new Date().toISOString()); window.alert("数据恢复成功，网站将重新加载。"); window.location.reload(); } catch { flash("无法恢复：请选择由上岸手账导出的 JSON 备份文件"); } finally { event.target.value = ""; } };
    reader.readAsText(file);
  };
  return <section className="panel backup-center"><div className="backup-heading"><div><p className="eyebrow">DATA BACKUP</p><h2>全站数据备份与恢复</h2><p>包含计划、打卡、学习时长、练习、错题、成语、考试日期、主题和栏目顺序。</p></div><div className={`backup-age ${days === null || days >= 7 ? "warning" : ""}`}><strong>{days === null ? "未备份" : days === 0 ? "今天" : `${days} 天`}</strong><span>{days === null ? "建议立即创建" : "距离上次备份"}</span></div></div><div className="backup-actions"><input hidden ref={importRef} type="file" accept="application/json,.json" onChange={restore} /><button className="primary-button" onClick={exportAll}><span>↓</span><b>一键导出全部数据</b><small>生成可迁移的 JSON 文件</small></button><button className="soft-button" onClick={() => importRef.current?.click()}><span>↑</span><b>一键导入恢复</b><small>覆盖当前浏览器数据</small></button>{autoSnapshot && <button className="soft-button" onClick={downloadAuto}><span>周</span><b>下载最近自动备份</b><small>用于电脑与手机迁移</small></button>}</div><label className="weekly-backup-toggle"><input type="checkbox" checked={weekly} onChange={e => setWeekly(e.target.checked)} /><i /><span><strong>每周自动生成浏览器内备份</strong><small>开启后每隔7天自动创建快照，并提醒你下载保存。</small></span></label><div className="migration-note"><b>电脑与手机迁移方法</b><span>在旧设备导出备份 → 通过微信、网盘或数据线传到新设备 → 在新设备打开网站并选择“一键导入恢复”。</span></div></section>;
}

function Settings({ examDates, setExamDates, flash }: { examDates: ExamDates; setExamDates: Dispatch<SetStateAction<ExamDates>>; flash: (x: string) => void }) {
  const themes = [{ name: "奶油鼠尾草", primary: "#6f8271", accent: "#e5a48d", bg: "#f4f0e8", ink: "#3e433e" }, { name: "蜜桃燕麦", primary: "#a76f61", accent: "#d7a85b", bg: "#f7eee6", ink: "#4d3d38" }, { name: "雾紫可可", primary: "#7c708d", accent: "#c98f98", bg: "#f1edf3", ink: "#403b47" }, { name: "抹茶红豆", primary: "#71805f", accent: "#a96565", bg: "#f1f0e6", ink: "#3e4339" }];
  const apply = (t: typeof themes[0]) => { const r = document.documentElement; r.style.setProperty("--primary", t.primary); r.style.setProperty("--accent", t.accent); r.style.setProperty("--page", t.bg); r.style.setProperty("--ink", t.ink); localStorage.setItem("shore-theme", JSON.stringify(t)); flash(`已切换为${t.name}`); };
  const custom = (key: string, value: string) => { document.documentElement.style.setProperty(key, value); const r = getComputedStyle(document.documentElement); localStorage.setItem("shore-theme", JSON.stringify({ primary: r.getPropertyValue("--primary"), accent: r.getPropertyValue("--accent"), bg: r.getPropertyValue("--page"), ink: r.getPropertyValue("--ink") })); };
  return <div className="page-stack"><section className="page-intro"><div><p className="eyebrow">EXAM & COLOR</p><h2>设置你的双考试目标</h2><p>国考与省考分别倒计时，日期可随时修改。</p></div></section>
    <BackupCenter flash={flash} />
    <section className="panel exam-settings"><PanelTitle title="考试日期" /><div className="exam-date-grid"><label><span>国考日期</span><input type="date" value={examDates.national} onChange={e => setExamDates({ ...examDates, national: e.target.value })} /><small>{daysUntil(examDates.national) === null ? "尚未设置" : `还有 ${daysUntil(examDates.national)} 天`}</small></label><label><span>省考日期</span><input type="date" value={examDates.provincial} onChange={e => setExamDates({ ...examDates, provincial: e.target.value })} /><small>{daysUntil(examDates.provincial) === null ? "尚未设置" : `还有 ${daysUntil(examDates.provincial)} 天`}</small></label></div><button className="primary-button" onClick={() => flash("考试日期已保存")}>保存考试日期</button></section>
    <section className="theme-grid">{themes.map(t => <button key={t.name} className="theme-card" onClick={() => apply(t)}><div className="swatches"><i style={{ background: t.bg }} /><i style={{ background: t.primary }} /><i style={{ background: t.accent }} /><i style={{ background: t.ink }} /></div><strong>{t.name}</strong><small>点击应用主题</small></button>)}</section>
    <section className="panel custom-theme"><PanelTitle title="自定义色彩" /><div className="color-controls"><label><input type="color" defaultValue="#6f8271" onChange={e => custom("--primary", e.target.value)} /><span>主色</span></label><label><input type="color" defaultValue="#e5a48d" onChange={e => custom("--accent", e.target.value)} /><span>强调色</span></label><label><input type="color" defaultValue="#f4f0e8" onChange={e => custom("--page", e.target.value)} /><span>背景色</span></label><label><input type="color" defaultValue="#3e433e" onChange={e => custom("--ink", e.target.value)} /><span>文字色</span></label></div></section><section className="notice-card"><strong>数据说明</strong><p>任务、练习、时长和考试日期均保存在当前浏览器中。重新部署不会删除数据，清理浏览器缓存前请先备份。</p></section></div>;
}
