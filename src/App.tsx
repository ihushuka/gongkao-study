import { ChangeEvent, Dispatch, FormEvent, SetStateAction, useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";

type Tab = "home" | "plan" | "study" | "practice" | "idiom" | "mistakes" | "review" | "settings";
type Task = { id: number; title: string; subject: string; minutes: number; done: boolean; date?: string };
type Practice = { id: number; date: string; source: string; module: string; correct: number; total: number; minutes: number };
type StudySession = { id: number; date: string; module: string; seconds: number };
type ExamDates = { national: string; provincial: string };
type Mistake = { id: number; module: string; source: string; title: string; answer: string; mine: string; reason: string; image?: string; mastered: boolean };
type Idiom = { id: number; word: string; meaning: string; example: string; level: "未掌握" | "模糊" | "已掌握" };

const modules = ["资料分析", "判断推理", "言语理解", "数量关系", "常识判断"];
const sources = ["夸夸刷", "超大杯", "套卷", "粉笔题库", "真题", "错题重做"];
const nav: { id: Tab; label: string; mark: string }[] = [
  { id: "home", label: "今日主页", mark: "今" },
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
  let best: { tasks: Omit<Task, "id" | "done">[]; sheet: string; dated: number; recurring: number } = { tasks: [], sheet: "", dated: 0, recurring: 0 };
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
    const recurringTasks: Omit<Task, "id" | "done">[] = [];
    if (targetMonth) {
      const [year, month] = targetMonth.split("-").map(Number);
      const count = new Date(year, month, 0).getDate();
      Array.from(new Set(recurringTitles)).forEach(title => {
        for (let day = 1; day <= count; day += 1) {
          const date = `${targetMonth}-${pad(day)}`;
          const weekday = new Date(`${date}T12:00:00`).getDay();
          if (/除周日|周日除外/.test(title) && weekday === 0) continue;
          recurringTasks.push({ title, subject: inferSubject(title), minutes: inferMinutes(title), date });
        }
      });
    }
    const unique = Array.from(new Map([...datedTasks, ...recurringTasks].map(t => [`${t.date}|${t.title}`, t])).values());
    if (unique.length > best.tasks.length) best = { tasks: unique, sheet: sheetName, dated: datedTasks.length, recurring: recurringTasks.length };
  });
  return best;
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
  const [practices, setPractices] = useStoredState<Practice[]>("shore-practices", seedPractices);
  const [sessions, setSessions] = useStoredState<StudySession[]>("shore-study-sessions", seedSessions);
  const [examDates, setExamDates] = useStoredState<ExamDates>("shore-exam-dates", { national: "", provincial: "" });
  const [mistakes, setMistakes] = useStoredState<Mistake[]>("shore-mistakes", seedMistakes);
  const [idioms, setIdioms] = useStoredState<Idiom[]>("shore-idioms", seedIdioms);
  const [timerOn, setTimerOn] = useState(false);
  const [activeSeconds, setActiveSeconds] = useState(0);
  const [timerModule, setTimerModule] = useState("资料分析");
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
    if (!timerOn) { setTimerOn(true); return; }
    if (activeSeconds > 0) setSessions([...sessions, { id: Date.now(), date: localISO(), module: timerModule, seconds: activeSeconds }]);
    setTimerOn(false); setActiveSeconds(0); flash("本次学习时长已保存");
  };
  const todayTasks = tasks.filter(t => normalizedDate(t.date) === localISO());
  const todayDone = todayTasks.filter(t => t.done).length;
  const todaySeconds = sessions.filter(s => normalizedDate(s.date) === localISO()).reduce((sum, s) => sum + s.seconds, 0) + activeSeconds;

  const sharedTimer = { timerOn, activeSeconds, timerModule, setTimerModule, toggleTimer, todaySeconds };
  let page;
  if (tab === "home") page = <Dashboard tasks={tasks} setTasks={setTasks} practices={practices} mistakes={mistakes} idioms={idioms} examDates={examDates} setTab={setTab} timer={sharedTimer} />;
  else if (tab === "plan") page = <Plan tasks={tasks} setTasks={setTasks} flash={flash} />;
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
      <div className="sidebar-foot"><div className="streak"><span>今日任务</span><strong>{todayDone}<small>/{todayTasks.length}</small></strong></div><p>距离目标，再近一点点。</p></div>
    </aside>
    {mobileNav && <button className="scrim" aria-label="关闭菜单" onClick={() => setMobileNav(false)} />}
    <section className="content"><header className="topbar"><button className="menu-button" onClick={() => setMobileNav(true)}>☰</button><div><p>{new Date().toLocaleDateString("zh-CN", { month: "long", day: "numeric", weekday: "long" })}</p><h1>{nav.find(item => item.id === tab)?.label}</h1></div><div className="top-actions"><button className="soft-button" onClick={() => setTab("settings")}>◐ 设置</button><div className="avatar">W</div></div></header>{page}</section>
    {toast && <div className="toast">✓ {toast}</div>}
  </main>;
}

type TimerShared = { timerOn: boolean; activeSeconds: number; timerModule: string; setTimerModule: Dispatch<SetStateAction<string>>; toggleTimer: () => void; todaySeconds: number };

function Dashboard({ tasks, setTasks, practices, mistakes, idioms, examDates, setTab, timer }: { tasks: Task[]; setTasks: Dispatch<SetStateAction<Task[]>>; practices: Practice[]; mistakes: Mistake[]; idioms: Idiom[]; examDates: ExamDates; setTab: Dispatch<SetStateAction<Tab>>; timer: TimerShared }) {
  const todayTasks = tasks.filter(t => normalizedDate(t.date) === localISO());
  const done = todayTasks.filter(t => t.done).length;
  const completion = todayTasks.length ? Math.round(done / todayTasks.length * 100) : 0;
  const q = practices.reduce((s, x) => s + x.total, 0), c = practices.reduce((s, x) => s + x.correct, 0);
  const accuracy = q ? Math.round(c / q * 1000) / 10 : 0;
  const pendingM = mistakes.filter(x => !x.mastered).length, pendingI = idioms.filter(x => x.level !== "已掌握").length;
  const quote = quotes[new Date().getDate() % quotes.length];
  const examItems = [{ label: "国考", date: examDates.national }, { label: "省考", date: examDates.provincial }];
  return <div className="page-stack">
    <section className="hero-card"><div className="hero-copy"><p className="eyebrow">DAILY NOTE · 今日寄语</p><blockquote>“{quote}”</blockquote><p>把今天安排清楚，然后一项一项完成。</p></div><div className="exam-countdowns">{examItems.map(x => <button key={x.label} onClick={() => setTab("settings")}><span>{x.label}</span><strong>{daysUntil(x.date) ?? "—"}</strong><small>{x.date ? "天" : "设置日期"}</small></button>)}</div></section>
    <section className="metric-grid"><Metric label="今日完成" value={`${completion}%`} note={`${done}/${todayTasks.length} 项任务`} color="sage" /><Metric label="今日学习" value={formatHours(timer.todaySeconds)} note={timer.timerOn ? "正在计时" : "点击开始专注"} color="peach" /><Metric label="综合正确率" value={`${accuracy}%`} note={`累计 ${q} 题`} color="lilac" /><Metric label="等待复习" value={pendingM + pendingI} note={`${pendingM} 错题 · ${pendingI} 成语`} color="rose" /></section>
    <section className="dashboard-grid"><div className="panel task-panel"><PanelTitle title="今日学习清单" action="查看月计划" onClick={() => setTab("plan")} /><div className="progress-track"><i style={{ width: `${completion}%` }} /></div><div className="task-list">{todayTasks.length ? todayTasks.map(task => <label className={`task-row ${task.done ? "done" : ""}`} key={task.id}><input type="checkbox" checked={task.done} onChange={() => setTasks(tasks.map(x => x.id === task.id ? { ...x, done: !x.done } : x))} /><span className="custom-check">✓</span><span className="task-copy"><strong>{task.title}</strong><small>{task.subject}</small></span><em>{task.minutes} 分钟</em></label>) : <Empty text="今天还没有安排任务" />}</div></div>
      <div className="panel timer-panel"><PanelTitle title="专注计时" action="时长统计" onClick={() => setTab("study")} /><div className={`timer-orb ${timer.timerOn ? "running" : ""}`}><small>{timer.timerOn ? "本次专注" : "今日累计"}</small><strong>{formatClock(timer.timerOn ? timer.activeSeconds : timer.todaySeconds)}</strong><span>{timer.timerModule}</span></div><select value={timer.timerModule} onChange={e => timer.setTimerModule(e.target.value)}>{modules.map(m => <option key={m}>{m}</option>)}</select><button className="primary-button wide" onClick={timer.toggleTimer}>{timer.timerOn ? "结束并保存" : "开始学习"}</button></div></section>
  </div>;
}

function Metric({ label, value, note, color }: { label: string; value: string | number; note: string; color: string }) { return <article className={`metric-card ${color}`}><p>{label}</p><strong>{value}</strong><small>{note}</small></article>; }
function PanelTitle({ title, action, onClick }: { title: string; action?: string; onClick?: () => void }) { return <div className="panel-title"><h2>{title}</h2>{action && <button onClick={onClick}>{action} →</button>}</div>; }
function Empty({ text }: { text: string }) { return <div className="empty"><span>✓</span><p>{text}</p></div>; }

function Plan({ tasks, setTasks, flash }: { tasks: Task[]; setTasks: Dispatch<SetStateAction<Task[]>>; flash: (x: string) => void }) {
  const [month, setMonth] = useState(localISO().slice(0, 7));
  const [date, setDate] = useState(localISO()), [title, setTitle] = useState(""), [subject, setSubject] = useState("资料分析"), [minutes, setMinutes] = useState(45);
  const [importing, setImporting] = useState(false), [importNote, setImportNote] = useState("");
  const planFileRef = useRef<HTMLInputElement>(null);
  const monthTasks = tasks.filter(t => normalizedDate(t.date).startsWith(month));
  const finished = monthTasks.filter(t => t.done).length;
  const percent = monthTasks.length ? Math.round(finished / monthTasks.length * 100) : 0;
  const [year, monthNumber] = month.split("-").map(Number);
  const days = new Date(year, monthNumber, 0).getDate();
  const leading = (new Date(year, monthNumber - 1, 1).getDay() + 6) % 7;
  const cells = [...Array(leading).fill(null), ...Array.from({ length: days }, (_, i) => i + 1)];
  const add = (e: FormEvent) => { e.preventDefault(); if (!title.trim()) return; setTasks([...tasks, { id: Date.now(), title: title.trim(), subject, minutes, date, done: false }]); setTitle(""); setMonth(date.slice(0, 7)); flash("任务已加入对应日期"); };
  const importPlan = (e: ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0]; if (!file) return;
    setImporting(true); setImportNote("正在识别日期与任务…");
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const result = recognizePlanWorkbook(reader.result as ArrayBuffer);
        const existing = new Set(tasks.map(t => `${normalizedDate(t.date)}|${t.title}`));
        const fresh = result.tasks.filter(t => !existing.has(`${t.date}|${t.title}`)).map((t, i) => ({ ...t, id: Date.now() + i, done: false }));
        if (!fresh.length) { setImportNote("没有发现新的任务，可能已经导入过了"); flash("未发现可新增的计划任务"); }
        else {
          setTasks([...tasks, ...fresh]);
          const firstMonth = fresh[0].date?.slice(0, 7); if (firstMonth) setMonth(firstMonth);
          setImportNote(`已从“${result.sheet}”识别 ${fresh.length} 项：日期任务 ${result.dated} 项，并自动展开每日计划`);
          flash(`成功导入 ${fresh.length} 项月计划`);
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
    <section className="metric-grid compact"><Metric label="本月任务总量" value={monthTasks.length} note={`${month.replace("-", "年")}月`} color="sage" /><Metric label="本月已完成" value={finished} note={`剩余 ${monthTasks.length - finished} 项`} color="peach" /><Metric label="完成百分比" value={`${percent}%`} note="按任务数量计算" color="lilac" /><Metric label="计划总时长" value={`${Math.round(monthTasks.reduce((s, t) => s + t.minutes, 0) / 60 * 10) / 10}h`} note="预计投入" color="rose" /></section>
    <section className="month-layout"><div className="panel calendar-panel"><div className="calendar-week">{["一", "二", "三", "四", "五", "六", "日"].map(x => <span key={x}>周{x}</span>)}</div><div className="calendar-grid">{cells.map((day, i) => day === null ? <div className="calendar-cell blank" key={`b${i}`} /> : (() => { const key = `${month}-${pad(day)}`; const dayTasks = tasks.filter(t => normalizedDate(t.date) === key); return <div className={`calendar-cell ${key === localISO() ? "today" : ""}`} key={key}><button className="day-number" onClick={() => setDate(key)}>{day}</button><div className="day-tasks">{dayTasks.map(task => <label className={task.done ? "done" : ""} key={task.id}><input type="checkbox" checked={task.done} onChange={() => setTasks(tasks.map(x => x.id === task.id ? { ...x, done: !x.done } : x))} /><i /> <span>{task.title}</span></label>)}</div></div>; })())}</div></div>
      <form className="panel form-card sticky-form" onSubmit={add}><PanelTitle title="添加任务" /><label>任务日期<input type="date" value={date} onChange={e => setDate(e.target.value)} /></label><label>任务内容<input value={title} onChange={e => setTitle(e.target.value)} placeholder="例如：完成资料分析第12组" /></label><div className="form-grid"><label>学习模块<select value={subject} onChange={e => setSubject(e.target.value)}>{[...modules, "错题复盘", "成语积累"].map(x => <option key={x}>{x}</option>)}</select></label><label>预计时长（分钟）<input type="number" min="5" step="5" value={minutes} onChange={e => setMinutes(Number(e.target.value))} /></label></div><button className="primary-button wide">添加到月计划</button></form></section>
  </div>;
}

function StudyTime({ sessions, setSessions, flash, timer }: { sessions: StudySession[]; setSessions: Dispatch<SetStateAction<StudySession[]>>; flash: (x: string) => void; timer: TimerShared }) {
  const [date, setDate] = useState(localISO()), [module, setModule] = useState("资料分析"), [minutes, setMinutes] = useState(60);
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
  const addManual = (e: FormEvent) => { e.preventDefault(); if (minutes <= 0) return; setSessions([...sessions, { id: Date.now(), date, module, seconds: minutes * 60 }]); flash("学习时长已补录"); };
  return <div className="page-stack"><section className="page-intro"><div><p className="eyebrow">STUDY TIMER</p><h2>记录真正投入的时间</h2><p>实时计时与线下学习补录会统一进入统计。</p></div><div className="study-today"><strong>{formatHours(timer.todaySeconds)}</strong><span>今日累计</span></div></section>
    <section className="study-grid"><div className="panel focus-card"><PanelTitle title="当前专注" /><div className={`focus-clock ${timer.timerOn ? "running" : ""}`}><span>{timer.timerOn ? "正在学习" : "准备开始"}</span><strong>{formatClock(timer.activeSeconds)}</strong></div><select value={timer.timerModule} onChange={e => timer.setTimerModule(e.target.value)}>{modules.map(m => <option key={m}>{m}</option>)}</select><button className="primary-button wide" onClick={timer.toggleTimer}>{timer.timerOn ? "结束并计入统计" : "开始计时"}</button></div>
      <div className="panel"><PanelTitle title="近7天学习时长" /><div className="study-bars">{daily.map(x => <div key={x.day}><span>{formatHours(x.seconds)}</span><i style={{ height: `${Math.max(4, x.seconds / maxDay * 125)}px` }} /><small>{formatShortDate(x.day)}</small></div>)}</div></div>
      <div className="panel"><PanelTitle title="模块投入分布" /><div className="horizontal-bars">{moduleStats.map(x => <div key={x.module}><span>{x.module}</span><div><i style={{ width: `${x.seconds / maxModule * 100}%` }} /></div><strong>{Math.round(x.seconds / 360) / 10}h</strong></div>)}</div></div></section>
    <section className="panel study-heatmap-panel"><div className="heatmap-head"><div><p className="eyebrow">MONTHLY FOCUS</p><h2>月历学习时长</h2><span>颜色越深，代表当天投入时间越长。</span></div><input className="month-picker" type="month" value={heatMonth} onChange={e => setHeatMonth(e.target.value)} /></div><div className="heatmap-summary"><div><strong>{formatHours(heatTotal)}</strong><span>本月累计</span></div><div><strong>{studiedDays}</strong><span>学习天数</span></div><div><strong>{studiedDays ? formatHours(Math.round(heatTotal / studiedDays)) : "0h 0m"}</strong><span>日均时长</span></div></div><div className="heatmap-week">{["一", "二", "三", "四", "五", "六", "日"].map(x => <span key={x}>周{x}</span>)}</div><div className="study-heatmap">{heatCells.map((day, i) => day === null ? <div className="heat-day blank" key={`hb${i}`} /> : (() => { const item = heatData[day - 1]; return <button className={`heat-day level-${item.level} ${item.day === localISO() ? "today" : ""}`} key={item.day} onClick={() => { setDate(item.day); if (item.seconds) flash(`${formatShortDate(item.day)} 学习 ${formatHours(item.seconds)}`); }} title={`${item.day} · ${formatHours(item.seconds)}`}><span>{day}</span><strong>{item.seconds ? formatHours(item.seconds) : "—"}</strong></button>; })())}</div><div className="heat-legend"><span>少</span>{[0, 1, 2, 3, 4, 5].map(x => <i className={`level-${x}`} key={x} />)}<span>多</span><small>0 · &lt;30m · &lt;1h · &lt;2h · &lt;4h · ≥4h</small></div></section>
    <section className="two-col"><form className="panel form-card" onSubmit={addManual}><PanelTitle title="补录学习时长" /><div className="form-grid"><label>学习日期<input type="date" value={date} onChange={e => setDate(e.target.value)} /></label><label>学习模块<select value={module} onChange={e => setModule(e.target.value)}>{modules.map(m => <option key={m}>{m}</option>)}</select></label><label>学习分钟数<input type="number" min="1" value={minutes} onChange={e => setMinutes(Number(e.target.value))} /></label></div><button className="primary-button">保存时长</button></form>
      <div className="panel"><PanelTitle title="最近记录" /><div className="session-list">{sessions.slice().reverse().slice(0, 8).map(s => <div key={s.id}><span>{formatShortDate(s.date)}</span><strong>{s.module}</strong><em>{formatHours(s.seconds)}</em><button onClick={() => setSessions(sessions.filter(x => x.id !== s.id))}>删除</button></div>)}</div></div></section>
  </div>;
}

function PracticeView({ practices, setPractices, flash }: { practices: Practice[]; setPractices: Dispatch<SetStateAction<Practice[]>>; flash: (x: string) => void }) {
  const [range, setRange] = useState<7 | 30>(7), [date, setDate] = useState(localISO()), [source, setSource] = useState("超大杯"), [module, setModule] = useState("资料分析"), [correct, setCorrect] = useState(18), [total, setTotal] = useState(20), [minutes, setMinutes] = useState(24);
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
  const add = (e: FormEvent) => { e.preventDefault(); setPractices([...practices, { id: Date.now(), date, source, module, correct: Math.min(correct, total), total, minutes }]); flash("练习成绩已记录"); };
  return <div className="page-stack"><section className="page-intro"><div><p className="eyebrow">ACCURACY TRACKER</p><h2>看见每一次进步的轨迹</h2><p>按模块与来源拆分，周变化和月变化统一呈现。</p></div></section>
    <section className="metric-grid compact">{stats.slice(0, 4).map((s, i) => <Metric key={s.module} label={s.module} value={`${s.accuracy}%`} note={`累计 ${s.count} 题`} color={["sage", "peach", "lilac", "rose"][i]} />)}</section>
    <section className="analysis-grid"><div className="panel"><PanelTitle title="各模块正确率" /><div className="horizontal-bars">{stats.map(s => <div key={s.module}><span>{s.module}</span><div><i style={{ width: `${s.accuracy}%` }} /></div><strong>{s.accuracy}%</strong></div>)}</div></div>
      <div className="panel trend-panel"><div className="panel-title"><h2>最近练习变化</h2><div className="range-tabs"><button className={range === 7 ? "active" : ""} onClick={() => setRange(7)}>周变化</button><button className={range === 30 ? "active" : ""} onClick={() => setRange(30)}>月变化</button></div></div><div className={`change-badge ${change < 0 ? "down" : ""}`}>{change >= 0 ? "↑" : "↓"} 较前周期 {Math.abs(change)} 个百分点</div>{points.length ? <svg className="trend-svg" viewBox="0 0 700 225" role="img" aria-label="正确率变化折线图"><line x1="45" y1="40" x2="660" y2="40" /><line x1="45" y1="109" x2="660" y2="109" /><line x1="45" y1="178" x2="660" y2="178" /><line className="axis" x1="45" y1="195" x2="660" y2="195" /><polyline points={points.map(p => `${p.x},${p.y}`).join(" ")} />{points.map((p, i) => <g key={i}><circle cx={p.x} cy={p.y} r="5" /><text className="value" x={p.x} y={p.y - 11}>{p.pct}%</text><text className="date" x={p.x} y="216">{p.date}</text></g>)}</svg> : <Empty text="当前周期还没有练习记录" />}</div></section>
    <section className="two-col practice-bottom"><div className="panel"><PanelTitle title="练习明细" /><div className="data-table"><div className="table-head"><span>日期</span><span>来源</span><span>模块</span><span>正确率</span></div>{sorted.slice().reverse().slice(0, 10).map(p => <div key={p.id}><span>{formatShortDate(p.date)}</span><span>{p.source}</span><span>{p.module}</span><strong>{Math.round(p.correct / p.total * 100)}%</strong></div>)}</div></div>
      <form className="panel form-card practice-form" onSubmit={add}><PanelTitle title="记录一组练习" /><div className="form-grid"><label className="full-field">练习日期<input type="date" value={date} onChange={e => setDate(e.target.value)} /></label><label>练习来源<select value={source} onChange={e => setSource(e.target.value)}>{sources.map(x => <option key={x}>{x}</option>)}</select></label><label>学习模块<select value={module} onChange={e => setModule(e.target.value)}>{modules.map(x => <option key={x}>{x}</option>)}</select></label><label>正确题数<input type="number" min="0" value={correct} onChange={e => setCorrect(Number(e.target.value))} /></label><label>总题数<input type="number" min="1" value={total} onChange={e => setTotal(Number(e.target.value))} /></label><label>实际用时（分钟）<input type="number" min="1" value={minutes} onChange={e => setMinutes(Number(e.target.value))} /></label></div><button className="primary-button wide">保存练习记录</button></form></section>
  </div>;
}

function IdiomView({ idioms, setIdioms, flash }: { idioms: Idiom[]; setIdioms: Dispatch<SetStateAction<Idiom[]>>; flash: (x: string) => void }) {
  const [query, setQuery] = useState(""), [word, setWord] = useState(""), [meaning, setMeaning] = useState("");
  const filtered = idioms.filter(x => `${x.word}${x.meaning}`.includes(query));
  const rotate = (id: number) => setIdioms(idioms.map(x => x.id === id ? { ...x, level: x.level === "未掌握" ? "模糊" : x.level === "模糊" ? "已掌握" : "未掌握" } : x));
  return <div className="page-stack"><section className="page-intro"><div><p className="eyebrow">IDIOM NOTEBOOK</p><h2>成语不是背过，是能辨清</h2><p>点击掌握状态即可切换，复习中心会自动收集未掌握内容。</p></div><div className="stat-pills"><span>总数 <b>{idioms.length}</b></span><span>未掌握 <b>{idioms.filter(x => x.level !== "已掌握").length}</b></span></div></section>
    <section className="panel"><div className="toolbar"><input className="search" value={query} onChange={e => setQuery(e.target.value)} placeholder="搜索成语或词义…" /><button className="soft-button" onClick={() => { const blob = new Blob(["成语,词义,例句,掌握状态\n" + idioms.map(x => `${x.word},${x.meaning},${x.example},${x.level}`).join("\n")], { type: "text/csv" }); const a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = "成语积累.csv"; a.click(); }}>导出表格</button></div><div className="idiom-grid">{filtered.map(x => <article className="idiom-card" key={x.id}><div><span className={`level ${x.level}`}>{x.level}</span><button onClick={() => rotate(x.id)}>切换状态</button></div><h3>{x.word}</h3><p>{x.meaning}</p><small>例：{x.example}</small></article>)}</div></section>
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
