import { ChangeEvent, Dispatch, FormEvent, ReactNode, SetStateAction, useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";

type Tab = "home" | "daily" | "plan" | "study" | "practice" | "idiom" | "mistakes" | "review" | "settings";
type Task = { id: number; title: string; subject: string; minutes: number; done: boolean; date?: string; plannedStart?: string; plannedEnd?: string };
type DailyRoutine = { id: number; title: string; subject: string; minutes: number; month: string; plannedStart?: string; plannedEnd?: string; completedDates: string[] };
type Practice = { id: number; date: string; source: string; module: string; correct: number; total: number; minutes: number };
type StudySession = { id: number; date: string; module: string; seconds: number; startTime?: string; endTime?: string };
type ExamDates = { national: string; provincial: string };
type Mistake = { id: number; module: string; source: string; title: string; answer: string; mine: string; reason: string; image?: string; mastered: boolean };
type Idiom = { id: number; word: string; meaning: string; example: string; level: "未掌握" | "模糊" | "已掌握"; source?: string; meaningHighlights?: string[] };

const modules = ["资料分析", "判断推理", "言语理解", "数量关系", "常识判断"];
const sources = ["夸夸刷", "超大杯", "套卷", "粉笔题库", "真题", "错题重做"];
const nav: { id: Tab; label: string; mark: string }[] = [
  { id: "home", label: "今日主页", mark: "今" },
  { id: "daily", label: "每日计划", mark: "日" },
  { id: "plan", label: "月度计划", mark: "计" },
  { id: "study", label: "学习计时", mark: "时" },
  { id: "practice", label: "刷题统计", mark: "练" },
  { id: "idiom", label: "成语积累", mark: "词" },
  { id: "mistakes", label: "错题本", mark: "错" },
  { id: "review", label: "复习中心", mark: "复" },
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
const routineApplies = (routine: DailyRoutine, date: string) => !(/除周日|周日除外/.test(routine.title) && new Date(`${date}T12:00:00`).getDay() === 0);
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
const excelDate = (value: unknown) => {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return localISO(value);
  if (typeof value === "number" && value > 25000 && value < 80000) {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed) return `${parsed.y}-${pad(parsed.m)}-${pad(parsed.d)}`;
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
  const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
  let best: { tasks: Omit<Task, "id" | "done">[]; routines: Omit<DailyRoutine, "id" | "completedDates">[]; sheet: string; dated: number } = { tasks: [], routines: [], sheet: "", dated: 0 };
  workbook.SheetNames.forEach(sheetName => {
    const rows = XLSX.utils.sheet_to_json<unknown[]>(workbook.Sheets[sheetName], { header: 1, raw: true, defval: null });
    const dateRows = rows.map((row, rowIndex) => ({ rowIndex, dates: row.map((cell, col) => ({ col, date: excelDate(cell) })).filter(x => x.date) })).filter(x => x.dates.length >= 2);
    const datedTasks: Omit<Task, "id" | "done">[] = [];
    dateRows.forEach(entry => {
      entry.dates.forEach(({ col, date }) => {
        const nextMatchingRow = dateRows.find(next => next.rowIndex > entry.rowIndex && next.dates.some(candidate => Math.abs(candidate.col - col) <= 1));
        const stop = nextMatchingRow?.rowIndex ?? Math.min(rows.length, entry.rowIndex + 7);
        for (let r = entry.rowIndex + 1; r < stop; r += 1) {
          const candidates = [cleanPlanText(rows[r]?.[col]), cleanPlanText(rows[r]?.[col + 1])].filter(Boolean);
          const title = candidates.find(x => !/^(每日计划|星期[一二三四五六日天]|周[一二三四五六日天])$/.test(x) && !/^\d+$/.test(x));
          if (title) datedTasks.push({ title, subject: inferSubject(title), minutes: inferMinutes(title), date });
        }
      });
    });
    const planHeader = rows.findIndex(row => row.some(cell => cleanPlanText(cell) === "每日计划"));
    const recurringTitles = planHeader < 0 ? [] : rows.slice(planHeader + 1, planHeader + 9).flatMap(row => row.map(cleanPlanText)).filter(x => x && !/^(每日计划|星期|周[一二三四五六日天]|\d+)$/.test(x));
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
const findExample = (text: string, word: string) => { const compact = text.replace(/\s+/g, ""); const index = compact.indexOf(word); if (index < 0) return "待校对例句"; const before = compact.slice(Math.max(0, index - 55), index), after = compact.slice(index, index + 75); const start = Math.max(before.lastIndexOf("。"), before.lastIndexOf("；"), before.lastIndexOf("？"), before.lastIndexOf("！")); const endCandidates = [after.indexOf("。"), after.indexOf("；"), after.indexOf("？"), after.indexOf("！")].filter(x => x >= 0); const end = endCandidates.length ? Math.min(...endCandidates) + 1 : after.length; return `${before.slice(start + 1)}${after.slice(0, end)}`.replace(/^\d+[、.．]?/, ""); };
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
    const page = await pdf.getPage(pageNumber), viewport = page.getViewport({ scale: 1.7 }), canvas = document.createElement("canvas"); canvas.width = Math.round(viewport.width); canvas.height = Math.round(viewport.height); const context = canvas.getContext("2d"); if (!context) continue; await page.render({ canvas, canvasContext: context, viewport }).promise;
    const left = cropCanvas(canvas, { x: 0, y: 0, width: canvas.width * .52, height: canvas.height });
    const leftText = (await worker.recognize(left)).data.text;
    for (const box of findGreenBoxes(canvas)) {
      const greenText = (await worker.recognize(cropCanvas(canvas, box))).data.text, entries = parseGreenText(greenText);
      const blueTexts: string[] = [];
      for (const blue of findBlueBands(canvas, box)) { const key = (await worker.recognize(cropCanvas(canvas, blue))).data.text.replace(/\s+/g, "").replace(/[^㐀-鿿，。；、]/g, ""); if (key.length >= 2) blueTexts.push(key); }
      entries.forEach((entry, index) => found.push({ id: Date.now() + found.length, word: entry.word, meaning: entry.meaning, example: findExample(leftText, entry.word), level: "未掌握", source, meaningHighlights: blueTexts.length ? [blueTexts[Math.min(blueTexts.length - 1, Math.floor(index * blueTexts.length / entries.length))]] : [] }));
    }
  }
  await worker.terminate(); onProgress("识别完成，请校对后导入", 100);
  return Array.from(new Map(found.map(item => [`${item.source}|${item.word}`, item])).values());
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

  const flash = (message: string) => { setToast(message); window.setTimeout(() => setToast(""), 2200); };
  const toggleTimer = () => {
    if (!timerOn) { setTimerStartedAt(new Date()); setTimerOn(true); return; }
    const endedAt = new Date();
    if (activeSeconds > 0) setSessions([...sessions, { id: Date.now(), date: localISO(timerStartedAt || endedAt), module: timerModule, seconds: activeSeconds, startTime: timerStartedAt ? `${pad(timerStartedAt.getHours())}:${pad(timerStartedAt.getMinutes())}` : undefined, endTime: `${pad(endedAt.getHours())}:${pad(endedAt.getMinutes())}` }]);
    setTimerOn(false); setTimerStartedAt(null); setActiveSeconds(0); flash("本次学习时长已保存");
  };
  const todayTasks = tasks.filter(t => normalizedDate(t.date) === localISO());
  const todayRoutines = routines.filter(r => r.month === localISO().slice(0, 7) && routineApplies(r, localISO()));
  const todayDone = todayTasks.filter(t => t.done).length + todayRoutines.filter(r => r.completedDates.includes(localISO())).length;
  const todayTotal = todayTasks.length + todayRoutines.length;
  const todaySeconds = sessions.filter(s => normalizedDate(s.date) === localISO()).reduce((sum, s) => sum + s.seconds, 0) + activeSeconds;
  const nextTask = todayRoutines.find(r => !r.completedDates.includes(localISO()))?.title || todayTasks.find(t => !t.done)?.title || "";

  const sharedTimer = { timerOn, activeSeconds, timerModule, setTimerModule, toggleTimer, todaySeconds, timerStartedAt };
  let page;
  if (tab === "home") page = <Dashboard tasks={tasks} setTasks={setTasks} routines={routines} setRoutines={setRoutines} practices={practices} mistakes={mistakes} idioms={idioms} examDates={examDates} setTab={setTab} timer={sharedTimer} />;
  else if (tab === "daily") page = <DailyPlan tasks={tasks} setTasks={setTasks} routines={routines} setRoutines={setRoutines} sessions={sessions} timer={sharedTimer} />;
  else if (tab === "plan") page = <Plan tasks={tasks} setTasks={setTasks} routines={routines} setRoutines={setRoutines} flash={flash} />;
  else if (tab === "study") page = <StudyTime sessions={sessions} setSessions={setSessions} flash={flash} timer={sharedTimer} />;
  else if (tab === "practice") page = <PracticeView practices={practices} setPractices={setPractices} flash={flash} />;
  else if (tab === "idiom") page = <IdiomView idioms={idioms} setIdioms={setIdioms} flash={flash} />;
  else if (tab === "mistakes") page = <MistakeView mistakes={mistakes} setMistakes={setMistakes} flash={flash} />;
  else if (tab === "review") page = <ReviewView mistakes={mistakes} setMistakes={setMistakes} idioms={idioms} setIdioms={setIdioms} />;
  else page = <Settings examDates={examDates} setExamDates={setExamDates} flash={flash} />;

  return <main className="app-shell">
    <aside className={`sidebar ${mobileNav ? "open" : ""}`}>
      <div className="brand"><span className="brand-seal">岸</span><div><strong>上岸手账</strong><small>GONGKAO JOURNAL</small></div></div>
      <nav>{nav.map(item => <button key={item.id} className={tab === item.id ? "active" : ""} onClick={() => { setTab(item.id); setMobileNav(false); }}><span>{item.mark}</span>{item.label}</button>)}</nav>
      <div className="sidebar-foot"><div className="streak"><span>今日任务</span><strong>{todayDone}<small>/{todayTotal}</small></strong></div><p>距离目标，再近一点点。</p></div>
    </aside>
    {mobileNav && <button className="scrim" aria-label="关闭菜单" onClick={() => setMobileNav(false)} />}
    <section className="content"><header className="topbar"><button className="menu-button" onClick={() => setMobileNav(true)}>☰</button><div><p>{new Date().toLocaleDateString("zh-CN", { month: "long", day: "numeric", weekday: "long" })}</p><h1>{nav.find(item => item.id === tab)?.label}</h1></div><div className="top-actions"><button className="soft-button" onClick={() => setTab("settings")}>◐ 设置</button><div className="avatar">W</div></div></header>{page}</section>
    {toast && <div className="toast">✓ {toast}</div>}
    <FocusCoach nextTask={nextTask} timer={sharedTimer} onOpen={() => setTab(nextTask ? "daily" : "plan")} />
  </main>;
}

type TimerShared = { timerOn: boolean; activeSeconds: number; timerModule: string; setTimerModule: Dispatch<SetStateAction<string>>; toggleTimer: () => void; todaySeconds: number; timerStartedAt: Date | null };

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
    <section className="dashboard-grid"><div className="panel task-panel"><PanelTitle title="今日学习清单" action="打开时间轴" onClick={() => setTab("daily")} /><div className="progress-track"><i style={{ width: `${completion}%` }} /></div><div className="task-list">{total ? <>{todayRoutines.map(routine => { const checked = routine.completedDates.includes(localISO()); return <label className={`task-row ${checked ? "done" : ""}`} key={`r${routine.id}`}><input type="checkbox" checked={checked} onChange={() => setRoutines(routines.map(x => x.id === routine.id ? { ...x, completedDates: checked ? x.completedDates.filter(d => d !== localISO()) : [...x.completedDates, localISO()] } : x))} /><span className="custom-check">✓</span><span className="task-copy"><strong>{routine.title}</strong><small>每日任务 · {routine.subject}</small></span><em>{routine.minutes} 分钟</em></label>; })}{todayTasks.map(task => <label className={`task-row ${task.done ? "done" : ""}`} key={task.id}><input type="checkbox" checked={task.done} onChange={() => setTasks(tasks.map(x => x.id === task.id ? { ...x, done: !x.done } : x))} /><span className="custom-check">✓</span><span className="task-copy"><strong>{task.title}</strong><small>{task.subject}</small></span><em>{task.minutes} 分钟</em></label>)}</> : <Empty text="今天还没有安排任务" />}</div></div>
      <div className="panel timer-panel"><PanelTitle title="专注计时" action="时长统计" onClick={() => setTab("study")} /><div className={`timer-orb ${timer.timerOn ? "running" : ""}`}><small>{timer.timerOn ? "本次专注" : "今日累计"}</small><strong>{formatClock(timer.timerOn ? timer.activeSeconds : timer.todaySeconds)}</strong><span>{timer.timerModule}</span></div><select value={timer.timerModule} onChange={e => timer.setTimerModule(e.target.value)}>{modules.map(m => <option key={m}>{m}</option>)}</select><button className="primary-button wide" onClick={timer.toggleTimer}>{timer.timerOn ? "结束并保存" : "开始学习"}</button></div></section>
  </div>;
}

function Metric({ label, value, note, color }: { label: string; value: string | number; note: string; color: string }) { return <article className={`metric-card ${color}`}><p>{label}</p><strong>{value}</strong><small>{note}</small></article>; }
function PanelTitle({ title, action, onClick }: { title: string; action?: string; onClick?: () => void }) { return <div className="panel-title"><h2>{title}</h2>{action && <button onClick={onClick}>{action} →</button>}</div>; }
function Empty({ text }: { text: string }) { return <div className="empty"><span>✓</span><p>{text}</p></div>; }

function FocusCoach({ nextTask, timer, onOpen }: { nextTask: string; timer: TimerShared; onOpen: () => void }) {
  const [expanded, setExpanded] = useState(false), [snoozeUntil, setSnoozeUntil] = useState(0);
  useEffect(() => { if (!nextTask || timer.timerOn) return; const remind = () => { if (Date.now() < snoozeUntil) return; setExpanded(true); if ("Notification" in window && Notification.permission === "granted") new Notification("上岸手账提醒你", { body: `别把今天交给明天：${nextTask}` }); }; const first = window.setTimeout(remind, 20 * 60 * 1000); const repeat = window.setInterval(remind, 30 * 60 * 1000); return () => { clearTimeout(first); clearInterval(repeat); }; }, [nextTask, snoozeUntil, timer.timerOn]);
  if (!nextTask && !timer.timerOn) return null;
  return <aside className={`focus-coach ${expanded ? "expanded" : ""}`}><button className="coach-orb" onClick={() => setExpanded(!expanded)}>{timer.timerOn ? "专" : "督"}</button>{expanded && <div><small>{timer.timerOn ? "正在专注" : "现在该做"}</small><strong>{timer.timerOn ? `${timer.timerModule} · ${formatClock(timer.activeSeconds)}` : nextTask}</strong><p>{timer.timerOn ? "保持这一段专注，结束后记得保存。" : "如果现在不开始，这项任务很容易又被搁置。"}</p><span><button className="primary-button" onClick={() => { if (timer.timerOn) timer.toggleTimer(); else onOpen(); setExpanded(false); }}>{timer.timerOn ? "结束并保存" : "打开今日计划"}</button>{!timer.timerOn && <button className="soft-button" onClick={() => { setSnoozeUntil(Date.now() + 10 * 60 * 1000); setExpanded(false); }}>10分钟后提醒</button>}</span>{"Notification" in window && Notification.permission === "default" && <button className="coach-notify" onClick={() => Notification.requestPermission()}>开启桌面提醒</button>}</div>}</aside>;
}

function DailyPlan({ tasks, setTasks, routines, setRoutines, sessions, timer }: { tasks: Task[]; setTasks: Dispatch<SetStateAction<Task[]>>; routines: DailyRoutine[]; setRoutines: Dispatch<SetStateAction<DailyRoutine[]>>; sessions: StudySession[]; timer: TimerShared }) {
  const [date, setDate] = useState(localISO());
  const [editing, setEditing] = useState<{ kind: "task" | "routine"; id: number } | null>(null);
  const [editTitle, setEditTitle] = useState(""), [editStart, setEditStart] = useState("08:00"), [editEnd, setEditEnd] = useState("09:00");
  const dateTasks = tasks.filter(t => normalizedDate(t.date) === date);
  const dateRoutines = routines.filter(r => r.month === date.slice(0, 7) && routineApplies(r, date));
  const planned = [...dateRoutines.map(r => ({ kind: "routine" as const, id: r.id, title: r.title, subject: r.subject, start: r.plannedStart, end: r.plannedEnd, done: r.completedDates.includes(date) })), ...dateTasks.map(t => ({ kind: "task" as const, id: t.id, title: t.title, subject: t.subject, start: t.plannedStart, end: t.plannedEnd, done: t.done }))];
  const actual = sessions.filter(s => normalizedDate(s.date) === date).map(s => ({ ...s, running: false }));
  if (date === localISO() && timer.timerOn && timer.timerStartedAt) actual.push({ id: -1, date, module: timer.timerModule, seconds: timer.activeSeconds, startTime: `${pad(timer.timerStartedAt.getHours())}:${pad(timer.timerStartedAt.getMinutes())}`, endTime: `${pad(new Date().getHours())}:${pad(new Date().getMinutes())}`, running: true });
  const beginEdit = (item: typeof planned[number]) => { setEditing({ kind: item.kind, id: item.id }); setEditTitle(item.title); setEditStart(item.start || "08:00"); setEditEnd(item.end || "09:00"); };
  const saveEdit = () => {
    if (!editing) return;
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
  return <div className="page-stack"><section className="page-intro daily-intro"><div><p className="eyebrow">DAILY TIMELINE</p><h2>计划时间与真实投入并排看</h2><p>左边是预计安排，右边由学习计时自动生成实际完成时段。</p></div><input type="date" value={date} onChange={e => setDate(e.target.value)} /></section>
    <section className="daily-summary"><div><strong>{doneCount}/{planned.length}</strong><span>任务完成</span></div><div><strong>{formatHours(actual.reduce((sum, s) => sum + s.seconds, 0))}</strong><span>实际学习</span></div><div><strong>{planned.reduce((sum, x) => sum + Math.max(0, (timeMinutes(x.end) || 0) - (timeMinutes(x.start) || 0)), 0)}m</strong><span>计划投入</span></div></section>
    {editing && <section className="panel timeline-editor"><input value={editTitle} onChange={e => setEditTitle(e.target.value)} /><label>开始<input type="time" value={editStart} onChange={e => setEditStart(e.target.value)} /></label><label>结束<input type="time" value={editEnd} onChange={e => setEditEnd(e.target.value)} /></label><button className="primary-button" onClick={saveEdit}>保存安排</button><button className="soft-button" onClick={() => setEditing(null)}>取消</button></section>}
    <section className="panel daily-timeline-panel"><div className="timeline-head"><span>预计任务安排</span><b>时间</b><span>实际学习记录</span></div><div className="daily-timeline"><div className="timeline-side planned-side">{planned.filter(x => x.start).map(item => <article key={`${item.kind}${item.id}`} className={item.done ? "done" : ""} style={{ top: `${timelineTop(item.start)}%`, minHeight: `${timelineHeight(item.start, item.end)}%` }}><button className="timeline-check" onClick={() => togglePlanned(item)}>{item.done ? "✓" : ""}</button><div onClick={() => beginEdit(item)}><strong>{item.title}</strong><small>{item.start}–{item.end} · {item.subject}</small></div></article>)}</div><div className="timeline-axis">{Array.from({ length: 19 }, (_, i) => i + 6).map(h => <div key={h}><span>{pad(h)}:00</span><i /></div>)}</div><div className="timeline-side actual-side">{actual.filter(x => x.startTime).map(item => <article key={item.id} className={item.running ? "running" : ""} style={{ top: `${timelineTop(item.startTime)}%`, minHeight: `${timelineHeight(item.startTime, item.endTime)}%` }}><div><strong>{item.module}</strong><small>{item.startTime}–{item.endTime || "进行中"} · {formatHours(item.seconds)}</small></div></article>)}</div></div>
      {(planned.some(x => !x.start) || actual.some(x => !x.startTime)) && <div className="unscheduled"><div><strong>未设时间的任务</strong>{planned.filter(x => !x.start).map(item => <button key={`${item.kind}${item.id}`} onClick={() => beginEdit(item)}>{item.title} · 点击安排</button>)}</div><div><strong>未记录具体时间的学习</strong>{actual.filter(x => !x.startTime).map(item => <span key={item.id}>{item.module} · {formatHours(item.seconds)}</span>)}</div></div>}</section>
  </div>;
}

function Plan({ tasks, setTasks, routines, setRoutines, flash }: { tasks: Task[]; setTasks: Dispatch<SetStateAction<Task[]>>; routines: DailyRoutine[]; setRoutines: Dispatch<SetStateAction<DailyRoutine[]>>; flash: (x: string) => void }) {
  const [month, setMonth] = useState(localISO().slice(0, 7));
  const [date, setDate] = useState(localISO()), [title, setTitle] = useState(""), [subject, setSubject] = useState("资料分析"), [minutes, setMinutes] = useState(45);
  const [plannedStart, setPlannedStart] = useState("08:00"), [plannedEnd, setPlannedEnd] = useState("08:45");
  const [routineTitle, setRoutineTitle] = useState(""), [routineSubject, setRoutineSubject] = useState("资料分析"), [routineMinutes, setRoutineMinutes] = useState(30), [routineStart, setRoutineStart] = useState("07:30"), [routineEnd, setRoutineEnd] = useState("08:00");
  const [importing, setImporting] = useState(false), [importNote, setImportNote] = useState("");
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
  const add = (e: FormEvent) => { e.preventDefault(); if (!title.trim()) return; setTasks([...tasks, { id: Date.now(), title: title.trim(), subject, minutes, date, done: false, plannedStart, plannedEnd }]); setTitle(""); setMonth(date.slice(0, 7)); flash("任务已加入对应日期"); };
  const addRoutine = (e: FormEvent) => { e.preventDefault(); if (!routineTitle.trim()) return; setRoutines([...routines, { id: Date.now(), title: routineTitle.trim(), subject: routineSubject, minutes: routineMinutes, month, plannedStart: routineStart, plannedEnd: routineEnd, completedDates: [] }]); setRoutineTitle(""); flash("每日任务已加入本月"); };
  const importPlan = (e: ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0]; if (!file) return;
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
    <section className="page-intro"><div><p className="eyebrow">MONTHLY PLAN</p><h2>把月目标拆到每一天</h2><p>每个日期都可以添加、查看和勾选任务。</p></div><div className="intro-actions"><input hidden ref={planFileRef} type="file" accept=".xlsx,.xls" onChange={importPlan} /><button className="import-plan-button" onClick={() => planFileRef.current?.click()} disabled={importing}><span>表</span><div><strong>{importing ? "正在识别…" : "一键导入计划表"}</strong><small>自动识别日期与每日任务</small></div></button><input className="month-picker" type="month" value={month} onChange={e => setMonth(e.target.value)} /></div></section>
    {importNote && <section className="import-result"><span>✓</span><p>{importNote}</p><button onClick={() => setImportNote("")}>关闭</button></section>}
    <section className="metric-grid compact"><Metric label="本月任务总量" value={overallTotal} note={`${monthTasks.length} 单日 · ${routineTarget} 每日`} color="sage" /><Metric label="本月已完成" value={overallFinished} note={`剩余 ${overallTotal - overallFinished} 项`} color="peach" /><Metric label="完成百分比" value={`${overallPercent}%`} note="含每日任务完成次数" color="lilac" /><Metric label="计划总时长" value={`${Math.round(plannedTotalMinutes / 60 * 10) / 10}h`} note="预计投入" color="rose" /></section>
    <section className="panel routine-panel"><div className="panel-title"><div><h2>本月每日任务</h2><p>这些任务每天都要完成，单独统计，不显示在月历日期格中。</p></div><span>{monthRoutines.length} 项</span></div><div className="routine-grid">{monthRoutines.map(r => { const monthDays = new Date(year, monthNumber, 0).getDate(); const doneDays = r.completedDates.filter(d => d.startsWith(month)).length; return <article key={r.id}><div><span>{r.plannedStart || "待定"}–{r.plannedEnd || "待定"}</span><button onClick={() => setRoutines(routines.filter(x => x.id !== r.id))}>删除</button></div><strong>{r.title}</strong><small>{r.subject} · 本月 {doneDays}/{monthDays} 天完成</small><div className="progress-track"><i style={{ width: `${doneDays / monthDays * 100}%` }} /></div></article>; })}</div><form className="routine-form" onSubmit={addRoutine}><input value={routineTitle} onChange={e => setRoutineTitle(e.target.value)} placeholder="新增每天都要完成的任务" /><select value={routineSubject} onChange={e => setRoutineSubject(e.target.value)}>{[...modules, "错题复盘", "成语积累", "其他计划"].map(x => <option key={x}>{x}</option>)}</select><input type="time" value={routineStart} onChange={e => setRoutineStart(e.target.value)} /><input type="time" value={routineEnd} onChange={e => setRoutineEnd(e.target.value)} /><input type="number" min="5" value={routineMinutes} onChange={e => setRoutineMinutes(Number(e.target.value))} /><button className="primary-button">添加每日任务</button></form></section>
    <section className="month-layout"><div className="panel calendar-panel"><div className="calendar-week">{["一", "二", "三", "四", "五", "六", "日"].map(x => <span key={x}>周{x}</span>)}</div><div className="calendar-grid">{cells.map((day, i) => day === null ? <div className="calendar-cell blank" key={`b${i}`} /> : (() => { const key = `${month}-${pad(day)}`; const dayTasks = tasks.filter(t => normalizedDate(t.date) === key); return <div className={`calendar-cell ${key === localISO() ? "today" : ""}`} key={key}><button className="day-number" onClick={() => setDate(key)}>{day}</button><div className="day-tasks">{dayTasks.map(task => <label className={task.done ? "done" : ""} key={task.id}><input type="checkbox" checked={task.done} onChange={() => setTasks(tasks.map(x => x.id === task.id ? { ...x, done: !x.done } : x))} /><i /> <span>{task.title}</span></label>)}</div></div>; })())}</div></div>
      <form className="panel form-card sticky-form" onSubmit={add}><PanelTitle title="添加单日任务" /><label>任务日期<input type="date" value={date} onChange={e => setDate(e.target.value)} /></label><label>任务内容<input value={title} onChange={e => setTitle(e.target.value)} placeholder="例如：完成资料分析第12组" /></label><div className="form-grid"><label>学习模块<select value={subject} onChange={e => setSubject(e.target.value)}>{[...modules, "错题复盘", "成语积累"].map(x => <option key={x}>{x}</option>)}</select></label><label>预计时长（分钟）<input type="number" min="5" step="5" value={minutes} onChange={e => setMinutes(Number(e.target.value))} /></label><label>计划开始<input type="time" value={plannedStart} onChange={e => setPlannedStart(e.target.value)} /></label><label>计划结束<input type="time" value={plannedEnd} onChange={e => setPlannedEnd(e.target.value)} /></label></div><button className="primary-button wide">添加到月计划</button></form></section>
  </div>;
}

function StudyTime({ sessions, setSessions, flash, timer }: { sessions: StudySession[]; setSessions: Dispatch<SetStateAction<StudySession[]>>; flash: (x: string) => void; timer: TimerShared }) {
  const [date, setDate] = useState(localISO()), [module, setModule] = useState("资料分析"), [minutes, setMinutes] = useState(60);
  const [startTime, setStartTime] = useState("19:00"), [endTime, setEndTime] = useState("20:00"), [editingId, setEditingId] = useState<number | null>(null);
  const [heatMonth, setHeatMonth] = useState(localISO().slice(0, 7));
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
  const addManual = (e: FormEvent) => { e.preventDefault(); const byRange = (timeMinutes(endTime) || 0) - (timeMinutes(startTime) || 0); const finalMinutes = byRange > 0 ? byRange : minutes; if (finalMinutes <= 0) return; const row = { date, module, seconds: finalMinutes * 60, startTime, endTime }; if (editingId !== null) { setSessions(sessions.map(s => s.id === editingId ? { ...s, ...row } : s)); setEditingId(null); flash("学习记录已修改"); } else { setSessions([...sessions, { id: Date.now(), ...row }]); flash("学习时长已补录"); } };
  const editSession = (s: StudySession) => { setEditingId(s.id); setDate(normalizedDate(s.date)); setModule(s.module); setMinutes(Math.round(s.seconds / 60)); setStartTime(s.startTime || "19:00"); setEndTime(s.endTime || "20:00"); };
  return <div className="page-stack"><section className="page-intro"><div><p className="eyebrow">STUDY TIMER</p><h2>记录真正投入的时间</h2><p>实时计时与线下学习补录会统一进入统计。</p></div><div className="study-today"><strong>{formatHours(timer.todaySeconds)}</strong><span>今日累计</span></div></section>
    <section className="study-grid"><div className="panel focus-card"><PanelTitle title="当前专注" /><div className={`focus-clock ${timer.timerOn ? "running" : ""}`}><span>{timer.timerOn ? "正在学习" : "准备开始"}</span><strong>{formatClock(timer.activeSeconds)}</strong></div><select value={timer.timerModule} onChange={e => timer.setTimerModule(e.target.value)}>{modules.map(m => <option key={m}>{m}</option>)}</select><button className="primary-button wide" onClick={timer.toggleTimer}>{timer.timerOn ? "结束并计入统计" : "开始计时"}</button></div>
      <div className="panel"><PanelTitle title="近7天学习时长" /><div className="study-bars">{daily.map(x => <div key={x.day}><span>{formatHours(x.seconds)}</span><i style={{ height: `${Math.max(4, x.seconds / maxDay * 125)}px` }} /><small>{formatShortDate(x.day)}</small></div>)}</div></div>
      <div className="panel"><PanelTitle title="模块投入分布" /><div className="horizontal-bars">{moduleStats.map(x => <div key={x.module}><span>{x.module}</span><div><i style={{ width: `${x.seconds / maxModule * 100}%` }} /></div><strong>{Math.round(x.seconds / 360) / 10}h</strong></div>)}</div></div></section>
    <section className="panel study-heatmap-panel"><div className="heatmap-head"><div><p className="eyebrow">MONTHLY FOCUS</p><h2>月历学习时长</h2><span>颜色越深，代表当天投入时间越长。</span></div><input className="month-picker" type="month" value={heatMonth} onChange={e => setHeatMonth(e.target.value)} /></div><div className="heatmap-summary"><div><strong>{formatHours(heatTotal)}</strong><span>本月累计</span></div><div><strong>{studiedDays}</strong><span>学习天数</span></div><div><strong>{studiedDays ? formatHours(Math.round(heatTotal / studiedDays)) : "0h 0m"}</strong><span>日均时长</span></div></div><div className="heatmap-week">{["一", "二", "三", "四", "五", "六", "日"].map(x => <span key={x}>周{x}</span>)}</div><div className="study-heatmap">{heatCells.map((day, i) => day === null ? <div className="heat-day blank" key={`hb${i}`} /> : (() => { const item = heatData[day - 1]; return <button className={`heat-day level-${item.level} ${item.day === localISO() ? "today" : ""}`} key={item.day} onClick={() => { setDate(item.day); if (item.seconds) flash(`${formatShortDate(item.day)} 学习 ${formatHours(item.seconds)}`); }} title={`${item.day} · ${formatHours(item.seconds)}`}><span>{day}</span><strong>{item.seconds ? formatHours(item.seconds) : "—"}</strong></button>; })())}</div><div className="heat-legend"><span>少</span>{[0, 1, 2, 3, 4, 5].map(x => <i className={`level-${x}`} key={x} />)}<span>多</span><small>0 · &lt;30m · &lt;1h · &lt;2h · &lt;4h · ≥4h</small></div></section>
    <section className="two-col"><form className="panel form-card" onSubmit={addManual}><PanelTitle title={editingId === null ? "补录学习时长" : "修改学习记录"} /><div className="form-grid"><label>学习日期<input type="date" value={date} onChange={e => setDate(e.target.value)} /></label><label>学习模块<select value={module} onChange={e => setModule(e.target.value)}>{modules.map(m => <option key={m}>{m}</option>)}</select></label><label>开始时间<input type="time" value={startTime} onChange={e => setStartTime(e.target.value)} /></label><label>结束时间<input type="time" value={endTime} onChange={e => setEndTime(e.target.value)} /></label><label>学习分钟数<input type="number" min="1" value={minutes} onChange={e => setMinutes(Number(e.target.value))} /><small>设置有效时间段后优先按时间段计算</small></label></div><button className="primary-button">{editingId === null ? "保存时长" : "保存修改"}</button>{editingId !== null && <button type="button" className="soft-button cancel-edit" onClick={() => setEditingId(null)}>取消修改</button>}</form>
      <div className="panel"><PanelTitle title="最近记录" /><div className="session-list">{sessions.slice().reverse().slice(0, 8).map(s => <div key={s.id}><span>{formatShortDate(s.date)}<small>{s.startTime && `${s.startTime}–${s.endTime || ""}`}</small></span><strong>{s.module}</strong><em>{formatHours(s.seconds)}</em><span className="row-actions"><button onClick={() => editSession(s)}>修改</button><button onClick={() => { setSessions(sessions.filter(x => x.id !== s.id)); if (editingId === s.id) setEditingId(null); }}>删除</button></span></div>)}</div></div></section>
  </div>;
}

function PracticeView({ practices, setPractices, flash }: { practices: Practice[]; setPractices: Dispatch<SetStateAction<Practice[]>>; flash: (x: string) => void }) {
  const [range, setRange] = useState<7 | 30>(7), [date, setDate] = useState(localISO()), [source, setSource] = useState("超大杯"), [module, setModule] = useState("资料分析"), [correct, setCorrect] = useState(18), [total, setTotal] = useState(20), [minutes, setMinutes] = useState(24);
  const [editingId, setEditingId] = useState<number | null>(null);
  const sorted = useMemo(() => practices.slice().sort((a, b) => normalizedDate(a.date).localeCompare(normalizedDate(b.date))), [practices]);
  const cutoff = new Date(); cutoff.setDate(cutoff.getDate() - range + 1);
  const previousCutoff = new Date(); previousCutoff.setDate(previousCutoff.getDate() - range * 2 + 1);
  const currentRows = sorted.filter(p => new Date(`${normalizedDate(p.date)}T12:00:00`) >= cutoff);
  const previousRows = sorted.filter(p => { const d = new Date(`${normalizedDate(p.date)}T12:00:00`); return d >= previousCutoff && d < cutoff; });
  const rate = (rows: Practice[]) => { const q = rows.reduce((s, x) => s + x.total, 0), c = rows.reduce((s, x) => s + x.correct, 0); return q ? c / q * 100 : 0; };
  const change = Math.round((rate(currentRows) - rate(previousRows)) * 10) / 10;
  const chartRows = currentRows.slice(-12);
  const points = chartRows.map((p, i) => ({ x: chartRows.length === 1 ? 350 : 50 + i * 600 / (chartRows.length - 1), y: 178 - p.correct / p.total * 138, pct: Math.round(p.correct / p.total * 100), date: formatShortDate(p.date) }));
  const stats = modules.map(m => { const rows = practices.filter(p => p.module === m); return { module: m, accuracy: Math.round(rate(rows)), count: rows.reduce((s, x) => s + x.total, 0) }; });
  const add = (e: FormEvent) => { e.preventDefault(); const row = { date, source, module, correct: Math.min(correct, total), total, minutes }; if (editingId !== null) { setPractices(practices.map(p => p.id === editingId ? { ...p, ...row } : p)); setEditingId(null); flash("练习记录已修改"); } else { setPractices([...practices, { id: Date.now(), ...row }]); flash("练习成绩已记录"); } };
  const editPractice = (p: Practice) => { setEditingId(p.id); setDate(normalizedDate(p.date)); setSource(p.source); setModule(p.module); setCorrect(p.correct); setTotal(p.total); setMinutes(p.minutes); };
  return <div className="page-stack"><section className="page-intro"><div><p className="eyebrow">ACCURACY TRACKER</p><h2>看见每一次进步的轨迹</h2><p>按模块与来源拆分，周变化和月变化统一呈现。</p></div></section>
    <section className="metric-grid compact">{stats.slice(0, 4).map((s, i) => <Metric key={s.module} label={s.module} value={`${s.accuracy}%`} note={`累计 ${s.count} 题`} color={["sage", "peach", "lilac", "rose"][i]} />)}</section>
    <section className="analysis-grid"><div className="panel"><PanelTitle title="各模块正确率" /><div className="horizontal-bars">{stats.map(s => <div key={s.module}><span>{s.module}</span><div><i style={{ width: `${s.accuracy}%` }} /></div><strong>{s.accuracy}%</strong></div>)}</div></div>
      <div className="panel trend-panel"><div className="panel-title"><h2>最近练习变化</h2><div className="range-tabs"><button className={range === 7 ? "active" : ""} onClick={() => setRange(7)}>周变化</button><button className={range === 30 ? "active" : ""} onClick={() => setRange(30)}>月变化</button></div></div><div className={`change-badge ${change < 0 ? "down" : ""}`}>{change >= 0 ? "↑" : "↓"} 较前周期 {Math.abs(change)} 个百分点</div>{points.length ? <svg className="trend-svg" viewBox="0 0 700 225" role="img" aria-label="正确率变化折线图"><line x1="45" y1="40" x2="660" y2="40" /><line x1="45" y1="109" x2="660" y2="109" /><line x1="45" y1="178" x2="660" y2="178" /><line className="axis" x1="45" y1="195" x2="660" y2="195" /><polyline points={points.map(p => `${p.x},${p.y}`).join(" ")} />{points.map((p, i) => <g key={i}><circle cx={p.x} cy={p.y} r="5" /><text className="value" x={p.x} y={p.y - 11}>{p.pct}%</text><text className="date" x={p.x} y="216">{p.date}</text></g>)}</svg> : <Empty text="当前周期还没有练习记录" />}</div></section>
    <section className="two-col practice-bottom"><div className="panel"><PanelTitle title="练习明细" /><div className="data-table editable-table"><div className="table-head"><span>日期</span><span>来源</span><span>模块</span><span>正确率</span><span>操作</span></div>{sorted.slice().reverse().slice(0, 10).map(p => <div key={p.id}><span>{formatShortDate(p.date)}</span><span>{p.source}</span><span>{p.module}</span><strong>{Math.round(p.correct / p.total * 100)}%</strong><span className="row-actions"><button onClick={() => editPractice(p)}>修改</button><button onClick={() => { setPractices(practices.filter(x => x.id !== p.id)); if (editingId === p.id) setEditingId(null); }}>删除</button></span></div>)}</div></div>
      <form className="panel form-card practice-form" onSubmit={add}><PanelTitle title={editingId === null ? "记录一组练习" : "修改练习记录"} /><div className="form-grid"><label className="full-field">练习日期<input type="date" value={date} onChange={e => setDate(e.target.value)} /></label><label>练习来源<select value={source} onChange={e => setSource(e.target.value)}>{sources.map(x => <option key={x}>{x}</option>)}</select></label><label>学习模块<select value={module} onChange={e => setModule(e.target.value)}>{modules.map(x => <option key={x}>{x}</option>)}</select></label><label>正确题数<input type="number" min="0" value={correct} onChange={e => setCorrect(Number(e.target.value))} /></label><label>总题数<input type="number" min="1" value={total} onChange={e => setTotal(Number(e.target.value))} /></label><label>实际用时（分钟）<input type="number" min="1" value={minutes} onChange={e => setMinutes(Number(e.target.value))} /></label></div><button className="primary-button wide">{editingId === null ? "保存练习记录" : "保存修改"}</button>{editingId !== null && <button type="button" className="soft-button wide cancel-edit" onClick={() => setEditingId(null)}>取消修改</button>}</form></section>
  </div>;
}

function MeaningText({ item }: { item: Idiom }) {
  const highlights = (item.meaningHighlights || []).filter(Boolean); if (!highlights.length) return <>{item.meaning}</>;
  let parts: ReactNode[] = [item.meaning];
  highlights.forEach((key, keyIndex) => { const next: ReactNode[] = []; parts.forEach((part, partIndex) => { if (typeof part !== "string") { next.push(part); return; } const cleanKey = key.replace(/[，。；、]/g, ""); const at = cleanKey ? part.indexOf(cleanKey) : -1; if (at < 0) next.push(part); else next.push(part.slice(0, at), <mark key={`${keyIndex}-${partIndex}`}>{part.slice(at, at + cleanKey.length)}</mark>, part.slice(at + cleanKey.length)); }); parts = next; });
  return <>{parts}<span className="meaning-key">重点：{highlights.join("；")}</span></>;
}

function IdiomView({ idioms, setIdioms, flash }: { idioms: Idiom[]; setIdioms: Dispatch<SetStateAction<Idiom[]>>; flash: (x: string) => void }) {
  const [query, setQuery] = useState(""), [word, setWord] = useState(""), [meaning, setMeaning] = useState("");
  const [ocrBusy, setOcrBusy] = useState(false), [ocrText, setOcrText] = useState(""), [ocrProgress, setOcrProgress] = useState(0), [preview, setPreview] = useState<Idiom[]>([]);
  const pdfRef = useRef<HTMLInputElement>(null);
  const filtered = idioms.filter(x => `${x.word}${x.meaning}${x.source || ""}`.includes(query));
  const rotate = (id: number) => setIdioms(idioms.map(x => x.id === id ? { ...x, level: x.level === "未掌握" ? "模糊" : x.level === "模糊" ? "已掌握" : "未掌握" } : x));
  const importPdf = async (e: ChangeEvent<HTMLInputElement>) => { const file = e.target.files?.[0]; if (!file) return; setOcrBusy(true); setPreview([]); try { const rows = await recognizePeopleDailyPdf(file, (text, value) => { setOcrText(text); setOcrProgress(value); }); setPreview(rows); if (!rows.length) setOcrText("没有识别到绿色框内容，请确认版式与示例一致"); } catch (error) { console.error(error); setOcrText("识别失败。请检查网络后重试，首次使用需要加载中文识别模型。"); } finally { setOcrBusy(false); e.target.value = ""; } };
  const confirmPreview = () => { const existing = new Set(idioms.map(x => `${x.source || ""}|${x.word}`)); const fresh = preview.filter(x => !existing.has(`${x.source || ""}|${x.word}`)).map((x, i) => ({ ...x, id: Date.now() + i })); setIdioms([...idioms, ...fresh]); setPreview([]); flash(`已导入 ${fresh.length} 条人民日报成语`); };
  return <div className="page-stack"><section className="page-intro"><div><p className="eyebrow">IDIOM NOTEBOOK</p><h2>成语不是背过，是能辨清</h2><p>支持识别人民日报图片型 PDF：绿色框选成语、左侧例句与蓝色重点词义。</p></div><div className="stat-pills"><span>总数 <b>{idioms.length}</b></span><span>未掌握 <b>{idioms.filter(x => x.level !== "已掌握").length}</b></span></div></section>
    <section className="panel pdf-import-card"><div><p className="eyebrow">PEOPLE'S DAILY OCR</p><h2>导入人民日报笔记</h2><p>文件只在当前浏览器本地识别。首次使用需加载中文模型，四页通常需要等待一段时间。</p></div><input hidden ref={pdfRef} type="file" accept="application/pdf,.pdf" onChange={importPdf} /><button className="primary-button" disabled={ocrBusy} onClick={() => pdfRef.current?.click()}>{ocrBusy ? "正在识别…" : "选择 PDF 并识别"}</button>{(ocrBusy || ocrText) && <div className="ocr-progress"><div><i style={{ width: `${ocrProgress}%` }} /></div><span>{ocrText}</span><strong>{ocrProgress}%</strong></div>}</section>
    {preview.length > 0 && <section className="panel ocr-preview"><div className="panel-title"><div><h2>识别结果校对</h2><p>可以直接修改错字，确认后再加入成语本。</p></div><button className="primary-button" onClick={confirmPreview}>确认导入 {preview.length} 条</button></div><div className="ocr-preview-grid">{preview.map((item, i) => <article key={item.id}><label>来源<input value={item.source || ""} onChange={e => setPreview(preview.map((x, j) => j === i ? { ...x, source: e.target.value } : x))} /></label><label>成语<input value={item.word} onChange={e => setPreview(preview.map((x, j) => j === i ? { ...x, word: e.target.value } : x))} /></label><label>词义<textarea value={item.meaning} onChange={e => setPreview(preview.map((x, j) => j === i ? { ...x, meaning: e.target.value } : x))} /></label><label>蓝色重点<input value={(item.meaningHighlights || []).join("；")} onChange={e => setPreview(preview.map((x, j) => j === i ? { ...x, meaningHighlights: e.target.value.split("；").filter(Boolean) } : x))} /></label><label>左侧例句<textarea value={item.example} onChange={e => setPreview(preview.map((x, j) => j === i ? { ...x, example: e.target.value } : x))} /></label><button onClick={() => setPreview(preview.filter((_, j) => j !== i))}>移除此条</button></article>)}</div></section>}
    <section className="panel"><div className="toolbar"><input className="search" value={query} onChange={e => setQuery(e.target.value)} placeholder="搜索成语、词义或来源…" /><button className="soft-button" onClick={() => { const blob = new Blob(["来源,成语,词义,重点,例句,掌握状态\n" + idioms.map(x => `${x.source || ""},${x.word},${x.meaning},${(x.meaningHighlights || []).join("；")},${x.example},${x.level}`).join("\n")], { type: "text/csv" }); const a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = "成语积累.csv"; a.click(); }}>导出表格</button></div><div className="idiom-grid">{filtered.map(x => <article className="idiom-card" key={x.id}><div><span className={`level ${x.level}`}>{x.level}</span><button onClick={() => rotate(x.id)}>切换状态</button></div>{x.source && <b className="idiom-source">{x.source}</b>}<h3>{x.word}</h3><p><MeaningText item={x} /></p><small>例：{x.example || "待补充"}</small></article>)}</div></section>
    <form className="panel inline-form" onSubmit={e => { e.preventDefault(); if (!word || !meaning) return; setIdioms([...idioms, { id: Date.now(), word, meaning, example: "", level: "未掌握" }]); setWord(""); setMeaning(""); flash("成语已加入积累本"); }}><h3>快速添加</h3><input value={word} onChange={e => setWord(e.target.value)} placeholder="成语" /><input value={meaning} onChange={e => setMeaning(e.target.value)} placeholder="词义或易错点" /><button className="primary-button">添加</button></form></div>;
}

function MistakeView({ mistakes, setMistakes, flash }: { mistakes: Mistake[]; setMistakes: Dispatch<SetStateAction<Mistake[]>>; flash: (x: string) => void }) {
  const [filter, setFilter] = useState("全部"); const fileRef = useRef<HTMLInputElement>(null), imageRef = useRef<HTMLInputElement>(null);
  const importCsv = (e: ChangeEvent<HTMLInputElement>) => { const file = e.target.files?.[0]; if (!file) return; const reader = new FileReader(); reader.onload = () => { const rows: Mistake[] = String(reader.result).split(/\r?\n/).slice(1).filter(Boolean).map((line, i) => { const [module = "未分类", source = "批量导入", title = "导入错题", mine = "", answer = "", reason = "待复盘"] = line.split(","); return { id: Date.now() + i, module, source, title, mine, answer, reason, mastered: false }; }); setMistakes([...mistakes, ...rows]); flash(`成功导入 ${rows.length} 道错题`); }; reader.readAsText(file); e.target.value = ""; };
  const importImage = (e: ChangeEvent<HTMLInputElement>) => { const file = e.target.files?.[0]; if (!file) return; const reader = new FileReader(); reader.onload = () => { setMistakes([{ id: Date.now(), module: "待分类", source: "图片导入", title: file.name.replace(/\.[^.]+$/, ""), answer: "", mine: "", reason: "待复盘", image: String(reader.result), mastered: false }, ...mistakes]); flash("图片错题已导入"); }; reader.readAsDataURL(file); e.target.value = ""; };
  const shown = filter === "全部" ? mistakes : mistakes.filter(x => x.module === filter);
  return <div className="page-stack"><section className="page-intro"><div><p className="eyebrow">MISTAKE LIBRARY</p><h2>错题的终点，是不再错</h2><p>支持图片与 CSV 批量导入，按模块、来源和错因整理。</p></div><div className="import-actions"><input hidden ref={fileRef} type="file" accept=".csv" onChange={importCsv} /><input hidden ref={imageRef} type="file" accept="image/*" onChange={importImage} /><button className="soft-button" onClick={() => fileRef.current?.click()}>导入 CSV</button><button className="primary-button" onClick={() => imageRef.current?.click()}>导入题目图片</button></div></section>
    <section className="panel"><div className="toolbar"><div className="filter-tabs">{["全部", ...modules].map(x => <button key={x} className={filter === x ? "active" : ""} onClick={() => setFilter(x)}>{x}</button>)}</div><span>{shown.length} 道</span></div><div className="mistake-list">{shown.map(x => <article key={x.id} className={x.mastered ? "mastered" : ""}>{x.image && <img src={x.image} alt="导入的错题" />}<div className="mistake-main"><p><span>{x.module}</span><span>{x.source}</span></p><h3>{x.title}</h3><small>我的答案：{x.mine || "待补充"}　正确答案：{x.answer || "待补充"}　·　{x.reason}</small></div><button onClick={() => setMistakes(mistakes.map(y => y.id === x.id ? { ...y, mastered: !y.mastered } : y))}>{x.mastered ? "已掌握" : "标记掌握"}</button></article>)}</div></section><div className="notice-card"><strong>CSV 导入格式</strong><p>首行为表头，字段顺序：模块、来源、题目、我的答案、正确答案、错因。</p></div></div>;
}

function ReviewView({ mistakes, setMistakes, idioms, setIdioms }: { mistakes: Mistake[]; setMistakes: Dispatch<SetStateAction<Mistake[]>>; idioms: Idiom[]; setIdioms: Dispatch<SetStateAction<Idiom[]>> }) {
  const pendingM = mistakes.filter(x => !x.mastered), pendingI = idioms.filter(x => x.level !== "已掌握");
  return <div className="page-stack"><section className="page-intro"><div><p className="eyebrow">DAILY REVIEW</p><h2>今日复习队列</h2><p>先清错题，再过成语。少量、多次、稳定重复。</p></div><div className="review-count"><strong>{pendingM.length + pendingI.length}</strong><span>项待复习</span></div></section><section className="two-col"><div className="panel"><PanelTitle title={`错题复盘 · ${pendingM.length}`} />{pendingM.length ? pendingM.slice(0, 6).map(x => <div className="review-row" key={x.id}><span>{x.module.slice(0, 1)}</span><div><strong>{x.title}</strong><small>{x.reason}</small></div><button onClick={() => setMistakes(mistakes.map(y => y.id === x.id ? { ...y, mastered: true } : y))}>掌握</button></div>) : <Empty text="今日错题已清空" />}</div><div className="panel"><PanelTitle title={`成语复习 · ${pendingI.length}`} />{pendingI.length ? pendingI.slice(0, 6).map(x => <div className="review-row idiom" key={x.id}><span>{x.word.slice(0, 1)}</span><div><strong>{x.word}</strong><small>{x.meaning}</small></div><button onClick={() => setIdioms(idioms.map(y => y.id === x.id ? { ...y, level: "已掌握" } : y))}>记住了</button></div>) : <Empty text="今日成语已清空" />}</div></section></div>;
}

function Settings({ examDates, setExamDates, flash }: { examDates: ExamDates; setExamDates: Dispatch<SetStateAction<ExamDates>>; flash: (x: string) => void }) {
  const themes = [{ name: "奶油鼠尾草", primary: "#6f8271", accent: "#e5a48d", bg: "#f4f0e8", ink: "#3e433e" }, { name: "蜜桃燕麦", primary: "#a76f61", accent: "#d7a85b", bg: "#f7eee6", ink: "#4d3d38" }, { name: "雾紫可可", primary: "#7c708d", accent: "#c98f98", bg: "#f1edf3", ink: "#403b47" }, { name: "抹茶红豆", primary: "#71805f", accent: "#a96565", bg: "#f1f0e6", ink: "#3e4339" }];
  const apply = (t: typeof themes[0]) => { const r = document.documentElement; r.style.setProperty("--primary", t.primary); r.style.setProperty("--accent", t.accent); r.style.setProperty("--page", t.bg); r.style.setProperty("--ink", t.ink); localStorage.setItem("shore-theme", JSON.stringify(t)); flash(`已切换为${t.name}`); };
  const custom = (key: string, value: string) => { document.documentElement.style.setProperty(key, value); const r = getComputedStyle(document.documentElement); localStorage.setItem("shore-theme", JSON.stringify({ primary: r.getPropertyValue("--primary"), accent: r.getPropertyValue("--accent"), bg: r.getPropertyValue("--page"), ink: r.getPropertyValue("--ink") })); };
  return <div className="page-stack"><section className="page-intro"><div><p className="eyebrow">EXAM & COLOR</p><h2>设置你的双考试目标</h2><p>国考与省考分别倒计时，日期可随时修改。</p></div></section>
    <section className="panel exam-settings"><PanelTitle title="考试日期" /><div className="exam-date-grid"><label><span>国考日期</span><input type="date" value={examDates.national} onChange={e => setExamDates({ ...examDates, national: e.target.value })} /><small>{daysUntil(examDates.national) === null ? "尚未设置" : `还有 ${daysUntil(examDates.national)} 天`}</small></label><label><span>省考日期</span><input type="date" value={examDates.provincial} onChange={e => setExamDates({ ...examDates, provincial: e.target.value })} /><small>{daysUntil(examDates.provincial) === null ? "尚未设置" : `还有 ${daysUntil(examDates.provincial)} 天`}</small></label></div><button className="primary-button" onClick={() => flash("考试日期已保存")}>保存考试日期</button></section>
    <section className="theme-grid">{themes.map(t => <button key={t.name} className="theme-card" onClick={() => apply(t)}><div className="swatches"><i style={{ background: t.bg }} /><i style={{ background: t.primary }} /><i style={{ background: t.accent }} /><i style={{ background: t.ink }} /></div><strong>{t.name}</strong><small>点击应用主题</small></button>)}</section>
    <section className="panel custom-theme"><PanelTitle title="自定义色彩" /><div className="color-controls"><label><input type="color" defaultValue="#6f8271" onChange={e => custom("--primary", e.target.value)} /><span>主色</span></label><label><input type="color" defaultValue="#e5a48d" onChange={e => custom("--accent", e.target.value)} /><span>强调色</span></label><label><input type="color" defaultValue="#f4f0e8" onChange={e => custom("--page", e.target.value)} /><span>背景色</span></label><label><input type="color" defaultValue="#3e433e" onChange={e => custom("--ink", e.target.value)} /><span>文字色</span></label></div></section><section className="notice-card"><strong>数据说明</strong><p>任务、练习、时长和考试日期均保存在当前浏览器中。重新部署不会删除数据，清理浏览器缓存前请先备份。</p></section></div>;
}
