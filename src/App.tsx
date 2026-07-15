"use client";

import { ChangeEvent, Dispatch, FormEvent, SetStateAction, useEffect, useMemo, useRef, useState } from "react";

type Tab = "home" | "plan" | "practice" | "idiom" | "mistakes" | "review" | "settings";
type Task = { id: number; title: string; subject: string; minutes: number; done: boolean };
type Practice = { id: number; date: string; source: string; module: string; correct: number; total: number; minutes: number };
type Mistake = { id: number; module: string; source: string; title: string; answer: string; mine: string; reason: string; image?: string; mastered: boolean };
type Idiom = { id: number; word: string; meaning: string; example: string; level: "未掌握" | "模糊" | "已掌握" };

const nav: { id: Tab; label: string; mark: string }[] = [
  { id: "home", label: "今日主页", mark: "今" },
  { id: "plan", label: "学习计划", mark: "计" },
  { id: "practice", label: "刷题统计", mark: "练" },
  { id: "idiom", label: "成语积累", mark: "词" },
  { id: "mistakes", label: "错题本", mark: "错" },
  { id: "review", label: "复习中心", mark: "复" },
  { id: "settings", label: "主题设置", mark: "色" },
];

const seedTasks: Task[] = [
  { id: 1, title: "资料分析｜超大杯第 12 组", subject: "资料分析", minutes: 50, done: true },
  { id: 2, title: "言语理解｜逻辑填空 30 题", subject: "言语理解", minutes: 45, done: true },
  { id: 3, title: "判断推理｜夸夸刷图推", subject: "判断推理", minutes: 40, done: false },
  { id: 4, title: "复习昨日错题 12 道", subject: "错题复盘", minutes: 35, done: false },
];

const seedPractices: Practice[] = [
  { id: 1, date: "07-09", source: "夸夸刷", module: "言语理解", correct: 38, total: 50, minutes: 42 },
  { id: 2, date: "07-10", source: "超大杯", module: "资料分析", correct: 16, total: 20, minutes: 24 },
  { id: 3, date: "07-11", source: "套卷", module: "判断推理", correct: 30, total: 40, minutes: 36 },
  { id: 4, date: "07-12", source: "夸夸刷", module: "判断推理", correct: 34, total: 40, minutes: 33 },
  { id: 5, date: "07-13", source: "超大杯", module: "数量关系", correct: 9, total: 15, minutes: 25 },
  { id: 6, date: "07-14", source: "套卷", module: "言语理解", correct: 32, total: 40, minutes: 35 },
  { id: 7, date: "07-15", source: "超大杯", module: "资料分析", correct: 18, total: 20, minutes: 22 },
];

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

const quotes = [
  "今天不必惊艳，只要比昨天更扎实一点。",
  "把会做的做对，把不会的逐个变成会做。",
  "稳定不是慢，稳定是最可靠的加速。",
  "每一道认真复盘的错题，都在替未来的你加分。",
  "别等状态，先完成今天的第一道题。",
];

const modules = ["资料分析", "判断推理", "言语理解", "数量关系", "常识判断"];
const sources = ["夸夸刷", "超大杯", "套卷", "粉笔题库", "真题", "错题重做"];

function useStoredState<T>(key: string, initial: T) {
  const [value, setValue] = useState<T>(initial);
  const [ready, setReady] = useState(false);
  useEffect(() => {
    try {
      const saved = localStorage.getItem(key);
      // Hydrate browser-only study data after the client mounts.
      // eslint-disable-next-line react-hooks/set-state-in-effect
      if (saved) setValue(JSON.parse(saved));
    } catch { /* retain defaults */ }
    setReady(true);
  }, [key]);
  useEffect(() => {
    if (ready) localStorage.setItem(key, JSON.stringify(value));
  }, [key, ready, value]);
  return [value, setValue] as const;
}

export default function Home() {
  const [tab, setTab] = useState<Tab>("home");
  const [tasks, setTasks] = useStoredState("shore-tasks", seedTasks);
  const [practices, setPractices] = useStoredState("shore-practices", seedPractices);
  const [mistakes, setMistakes] = useStoredState("shore-mistakes", seedMistakes);
  const [idioms, setIdioms] = useStoredState("shore-idioms", seedIdioms);
  const [studySeconds, setStudySeconds] = useStoredState("shore-study-seconds", 7320);
  const [timerOn, setTimerOn] = useState(false);
  const [timerModule, setTimerModule] = useState("资料分析");
  const [mobileNav, setMobileNav] = useState(false);
  const [toast, setToast] = useState("");
  const timerRef = useRef<ReturnType<typeof setInterval> | null>(null);

  const done = tasks.filter((task) => task.done).length;
  const completion = tasks.length ? Math.round((done / tasks.length) * 100) : 0;
  const totalQuestions = practices.reduce((sum, item) => sum + item.total, 0);
  const totalCorrect = practices.reduce((sum, item) => sum + item.correct, 0);
  const accuracy = totalQuestions ? Math.round((totalCorrect / totalQuestions) * 1000) / 10 : 0;
  const pendingMistakes = mistakes.filter((item) => !item.mastered).length;
  const pendingIdioms = idioms.filter((item) => item.level !== "已掌握").length;
  const quote = quotes[new Date().getDate() % quotes.length];

  useEffect(() => {
    if (timerOn) timerRef.current = setInterval(() => setStudySeconds((s) => s + 1), 1000);
    else if (timerRef.current) clearInterval(timerRef.current);
    return () => { if (timerRef.current) clearInterval(timerRef.current); };
  }, [timerOn, setStudySeconds]);

  useEffect(() => {
    const saved = localStorage.getItem("shore-theme");
    if (!saved) return;
    try {
      const theme = JSON.parse(saved);
      const root = document.documentElement;
      root.style.setProperty("--primary", theme.primary);
      root.style.setProperty("--accent", theme.accent);
      root.style.setProperty("--page", theme.bg);
      root.style.setProperty("--ink", theme.ink);
    } catch { /* retain the default palette */ }
  }, []);

  const flash = (message: string) => {
    setToast(message);
    window.setTimeout(() => setToast(""), 2200);
  };

  const timeText = `${String(Math.floor(studySeconds / 3600)).padStart(2, "0")}:${String(Math.floor((studySeconds % 3600) / 60)).padStart(2, "0")}:${String(studySeconds % 60).padStart(2, "0")}`;

  const section = () => {
    if (tab === "home") return <Dashboard {...{ tasks, setTasks, completion, done, studySeconds, timeText, timerOn, setTimerOn, timerModule, setTimerModule, accuracy, totalQuestions, pendingMistakes, pendingIdioms, quote, setTab }} />;
    if (tab === "plan") return <Plan tasks={tasks} setTasks={setTasks} flash={flash} />;
    if (tab === "practice") return <PracticeView practices={practices} setPractices={setPractices} flash={flash} />;
    if (tab === "idiom") return <IdiomView idioms={idioms} setIdioms={setIdioms} flash={flash} />;
    if (tab === "mistakes") return <MistakeView mistakes={mistakes} setMistakes={setMistakes} flash={flash} />;
    if (tab === "review") return <ReviewView mistakes={mistakes} setMistakes={setMistakes} idioms={idioms} setIdioms={setIdioms} />;
    return <ThemeSettings flash={flash} />;
  };

  return (
    <main className="app-shell">
      <aside className={`sidebar ${mobileNav ? "open" : ""}`}>
        <div className="brand"><span className="brand-seal">岸</span><div><strong>上岸手账</strong><small>GONGKAO JOURNAL</small></div></div>
        <nav>{nav.map((item) => <button key={item.id} className={tab === item.id ? "active" : ""} onClick={() => { setTab(item.id); setMobileNav(false); }}><span>{item.mark}</span>{item.label}</button>)}</nav>
        <div className="sidebar-foot"><div className="streak"><span>连续打卡</span><strong>12 <small>天</small></strong></div><p>距离目标，再近一点点。</p></div>
      </aside>
      {mobileNav && <button className="scrim" aria-label="关闭菜单" onClick={() => setMobileNav(false)} />}
      <section className="content">
        <header className="topbar">
          <button className="menu-button" onClick={() => setMobileNav(true)}>☰</button>
          <div><p>{new Date().toLocaleDateString("zh-CN", { month: "long", day: "numeric", weekday: "long" })}</p><h1>{nav.find((item) => item.id === tab)?.label}</h1></div>
          <div className="top-actions"><button className="soft-button" onClick={() => setTab("settings")}>◐ 主题</button><div className="avatar">W</div></div>
        </header>
        {section()}
      </section>
      {toast && <div className="toast">✓ {toast}</div>}
    </main>
  );
}

type DashboardProps = {
  tasks: Task[]; setTasks: Dispatch<SetStateAction<Task[]>>; completion: number; done: number;
  studySeconds: number; timeText: string; timerOn: boolean; setTimerOn: Dispatch<SetStateAction<boolean>>;
  timerModule: string; setTimerModule: Dispatch<SetStateAction<string>>; accuracy: number; totalQuestions: number;
  pendingMistakes: number; pendingIdioms: number; quote: string; setTab: Dispatch<SetStateAction<Tab>>;
};

function Dashboard({ tasks, setTasks, completion, done, studySeconds, timeText, timerOn, setTimerOn, timerModule, setTimerModule, accuracy, totalQuestions, pendingMistakes, pendingIdioms, quote, setTab }: DashboardProps) {
  return <div className="page-stack">
    <section className="hero-card">
      <div className="hero-copy"><p className="eyebrow">DAILY NOTE · 今日寄语</p><blockquote>“{quote}”</blockquote><p>你已经连续学习 12 天，今天也稳稳向前。</p></div>
      <div className="hero-date"><span>JUL</span><strong>15</strong><small>距离省考预计还有 228 天</small></div>
    </section>
    <section className="metric-grid">
      <Metric label="今日完成" value={`${completion}%`} note={`${done}/${tasks.length} 项任务`} color="sage" />
      <Metric label="今日学习" value={`${Math.floor(studySeconds / 3600)}h ${Math.floor((studySeconds % 3600) / 60)}m`} note="目标 4 小时" color="peach" />
      <Metric label="近期正确率" value={`${accuracy}%`} note={`累计 ${totalQuestions} 题`} color="lilac" />
      <Metric label="等待复习" value={pendingMistakes + pendingIdioms} note={`${pendingMistakes} 错题 · ${pendingIdioms} 成语`} color="rose" />
    </section>
    <section className="dashboard-grid">
      <div className="panel task-panel"><PanelTitle title="今日学习清单" action="查看计划" onClick={() => setTab("plan")} /><div className="progress-track"><i style={{ width: `${completion}%` }} /></div>
        <div className="task-list">{tasks.map((task: Task) => <label className={`task-row ${task.done ? "done" : ""}`} key={task.id}><input type="checkbox" checked={task.done} onChange={() => setTasks(tasks.map((x: Task) => x.id === task.id ? { ...x, done: !x.done } : x))} /><span className="custom-check">✓</span><span className="task-copy"><strong>{task.title}</strong><small>{task.subject}</small></span><em>{task.minutes} 分钟</em></label>)}</div>
      </div>
      <div className="panel timer-panel"><PanelTitle title="专注计时" /><div className={`timer-orb ${timerOn ? "running" : ""}`}><small>{timerOn ? "正在专注" : "今日累计"}</small><strong>{timeText}</strong><span>{timerModule}</span></div><select value={timerModule} onChange={(e) => setTimerModule(e.target.value)}>{modules.map((m) => <option key={m}>{m}</option>)}</select><button className="primary-button wide" onClick={() => setTimerOn(!timerOn)}>{timerOn ? "暂停计时" : "开始学习"}</button></div>
    </section>
    <section className="panel"><PanelTitle title="本周学习脉络" action="查看完整统计" onClick={() => setTab("practice")} /><div className="week-chart">{[2.1, 3.4, 2.8, 4.2, 3.7, 4.5, 2.0].map((h, i) => <div key={i}><span>{h}h</span><i style={{ height: `${h * 26}px` }} /><small>{["一", "二", "三", "四", "五", "六", "日"][i]}</small></div>)}</div></section>
  </div>;
}

function Metric({ label, value, note, color }: { label: string; value: string | number; note: string; color: string }) { return <article className={`metric-card ${color}`}><p>{label}</p><strong>{value}</strong><small>{note}</small></article>; }
function PanelTitle({ title, action, onClick }: { title: string; action?: string; onClick?: () => void }) { return <div className="panel-title"><h2>{title}</h2>{action && <button onClick={onClick}>{action} →</button>}</div>; }

function Plan({ tasks, setTasks, flash }: { tasks: Task[]; setTasks: (x: Task[]) => void; flash: (x: string) => void }) {
  const [title, setTitle] = useState(""); const [subject, setSubject] = useState("资料分析"); const [minutes, setMinutes] = useState(45);
  const add = (e: FormEvent) => { e.preventDefault(); if (!title.trim()) return; setTasks([...tasks, { id: Date.now(), title, subject, minutes, done: false }]); setTitle(""); flash("任务已加入今日计划"); };
  return <div className="page-stack"><section className="page-intro"><div><p className="eyebrow">TODAY&apos;S PLAN</p><h2>把今天过得具体一点</h2><p>完成一项，勾掉一项。未完成的任务可留到明天继续。</p></div><div className="completion-ring" style={{ "--percent": `${tasks.length ? tasks.filter(t => t.done).length / tasks.length * 360 : 0}deg` } as React.CSSProperties}><strong>{tasks.filter(t => t.done).length}/{tasks.length}</strong><span>已完成</span></div></section>
    <section className="two-col"><div className="panel"><PanelTitle title="今日任务" />{tasks.map(task => <div className={`plan-row ${task.done ? "done" : ""}`} key={task.id}><button className="check-button" onClick={() => setTasks(tasks.map(x => x.id === task.id ? { ...x, done: !x.done } : x))}>{task.done ? "✓" : ""}</button><div><strong>{task.title}</strong><small>{task.subject} · 预计 {task.minutes} 分钟</small></div><button className="ghost-danger" onClick={() => setTasks(tasks.filter(x => x.id !== task.id))}>删除</button></div>)}</div>
      <form className="panel form-card" onSubmit={add}><PanelTitle title="添加任务" /><label>任务内容<input value={title} onChange={e => setTitle(e.target.value)} placeholder="例如：完成资料分析第 12 组" /></label><div className="form-grid"><label>学习模块<select value={subject} onChange={e => setSubject(e.target.value)}>{[...modules, "错题复盘", "成语积累"].map(x => <option key={x}>{x}</option>)}</select></label><label>预计时长<input type="number" min="5" step="5" value={minutes} onChange={e => setMinutes(Number(e.target.value))} /></label></div><button className="primary-button">加入今日计划</button></form></section>
  </div>;
}

function PracticeView({ practices, setPractices, flash }: { practices: Practice[]; setPractices: (x: Practice[]) => void; flash: (x: string) => void }) {
  const [source, setSource] = useState("超大杯"), [module, setModule] = useState("资料分析"); const [correct, setCorrect] = useState(18), [total, setTotal] = useState(20), [minutes, setMinutes] = useState(24);
  const add = (e: FormEvent) => { e.preventDefault(); setPractices([...practices, { id: Date.now(), date: new Date().toLocaleDateString("zh-CN", { month: "2-digit", day: "2-digit" }).replace("/", "-"), source, module, correct: Math.min(correct, total), total, minutes }]); flash("练习成绩已记录"); };
  const stats = useMemo(() => modules.map(m => { const rows = practices.filter(p => p.module === m); const q = rows.reduce((s, x) => s + x.total, 0); const c = rows.reduce((s, x) => s + x.correct, 0); return { module: m, accuracy: q ? Math.round(c / q * 100) : 0, count: q }; }), [practices]);
  return <div className="page-stack"><section className="page-intro"><div><p className="eyebrow">ACCURACY TRACKER</p><h2>看见每一次进步的轨迹</h2><p>按模块与练习来源拆分，正确率和速度一起看。</p></div></section>
    <section className="metric-grid compact">{stats.slice(0, 4).map((s, i) => <Metric key={s.module} label={s.module} value={`${s.accuracy}%`} note={`累计 ${s.count} 题`} color={["sage", "peach", "lilac", "rose"][i]} />)}</section>
    <section className="analysis-grid"><div className="panel"><PanelTitle title="各模块正确率" /><div className="horizontal-bars">{stats.map(s => <div key={s.module}><span>{s.module}</span><div><i style={{ width: `${s.accuracy}%` }} /></div><strong>{s.accuracy}%</strong></div>)}</div></div>
      <div className="panel"><PanelTitle title="最近练习变化" /><div className="trend-chart"><div className="grid-lines" />{practices.slice(-7).map((p, i) => { const pct = p.correct / p.total * 100; return <div className="trend-point" key={p.id} style={{ left: `${6 + i * 14.5}%`, bottom: `${Math.max(10, pct - 42)}%` }}><i /><span>{Math.round(pct)}%</span><small>{p.date}</small></div>; })}</div></div></section>
    <section className="two-col"><div className="panel"><PanelTitle title="练习明细" /><div className="data-table"><div className="table-head"><span>日期</span><span>来源</span><span>模块</span><span>正确率</span></div>{practices.slice().reverse().slice(0, 8).map(p => <div key={p.id}><span>{p.date}</span><span>{p.source}</span><span>{p.module}</span><strong>{Math.round(p.correct / p.total * 100)}%</strong></div>)}</div></div>
      <form className="panel form-card" onSubmit={add}><PanelTitle title="记录一组练习" /><div className="form-grid"><label>练习来源<select value={source} onChange={e => setSource(e.target.value)}>{sources.map(x => <option key={x}>{x}</option>)}</select></label><label>学习模块<select value={module} onChange={e => setModule(e.target.value)}>{modules.map(x => <option key={x}>{x}</option>)}</select></label><label>正确题数<input type="number" min="0" value={correct} onChange={e => setCorrect(Number(e.target.value))} /></label><label>总题数<input type="number" min="1" value={total} onChange={e => setTotal(Number(e.target.value))} /></label><label>实际用时（分钟）<input type="number" min="1" value={minutes} onChange={e => setMinutes(Number(e.target.value))} /></label></div><button className="primary-button">保存练习记录</button></form></section>
  </div>;
}

function IdiomView({ idioms, setIdioms, flash }: { idioms: Idiom[]; setIdioms: (x: Idiom[]) => void; flash: (x: string) => void }) {
  const [query, setQuery] = useState(""); const [word, setWord] = useState(""), [meaning, setMeaning] = useState(""); const filtered = idioms.filter(x => `${x.word}${x.meaning}`.includes(query));
  const rotate = (id: number) => setIdioms(idioms.map(x => x.id === id ? { ...x, level: x.level === "未掌握" ? "模糊" : x.level === "模糊" ? "已掌握" : "未掌握" } : x));
  return <div className="page-stack"><section className="page-intro"><div><p className="eyebrow">IDIOM NOTEBOOK</p><h2>成语不是背过，是能辨清</h2><p>点击掌握状态即可切换，复习中心会自动收集未掌握内容。</p></div><div className="stat-pills"><span>总数 <b>{idioms.length}</b></span><span>未掌握 <b>{idioms.filter(x => x.level !== "已掌握").length}</b></span></div></section>
    <section className="panel"><div className="toolbar"><input className="search" value={query} onChange={e => setQuery(e.target.value)} placeholder="搜索成语或词义…" /><button className="soft-button" onClick={() => { const blob = new Blob(["成语,词义,例句,掌握状态\n" + idioms.map(x => `${x.word},${x.meaning},${x.example},${x.level}`).join("\n")], { type: "text/csv" }); const a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = "成语积累.csv"; a.click(); }}>导出表格</button></div><div className="idiom-grid">{filtered.map(x => <article className="idiom-card" key={x.id}><div><span className={`level ${x.level}`}>{x.level}</span><button onClick={() => rotate(x.id)}>切换状态</button></div><h3>{x.word}</h3><p>{x.meaning}</p><small>例：{x.example}</small></article>)}</div></section>
    <form className="panel inline-form" onSubmit={e => { e.preventDefault(); if (!word || !meaning) return; setIdioms([...idioms, { id: Date.now(), word, meaning, example: "", level: "未掌握" }]); setWord(""); setMeaning(""); flash("成语已加入积累本"); }}><h3>快速添加</h3><input value={word} onChange={e => setWord(e.target.value)} placeholder="成语" /><input value={meaning} onChange={e => setMeaning(e.target.value)} placeholder="词义或易错点" /><button className="primary-button">添加</button></form>
  </div>;
}

function MistakeView({ mistakes, setMistakes, flash }: { mistakes: Mistake[]; setMistakes: (x: Mistake[]) => void; flash: (x: string) => void }) {
  const [filter, setFilter] = useState("全部"); const fileRef = useRef<HTMLInputElement>(null); const imageRef = useRef<HTMLInputElement>(null);
  const importCsv = (e: ChangeEvent<HTMLInputElement>) => { const file = e.target.files?.[0]; if (!file) return; const reader = new FileReader(); reader.onload = () => { const rows = String(reader.result).split(/\r?\n/).slice(1).filter(Boolean).map((line, i) => { const [module = "未分类", source = "批量导入", title = "导入错题", mine = "", answer = "", reason = "待复盘"] = line.split(","); return { id: Date.now() + i, module, source, title, mine, answer, reason, mastered: false }; }); setMistakes([...mistakes, ...rows]); flash(`成功导入 ${rows.length} 道错题`); }; reader.readAsText(file); e.target.value = ""; };
  const importImage = (e: ChangeEvent<HTMLInputElement>) => { const file = e.target.files?.[0]; if (!file) return; const reader = new FileReader(); reader.onload = () => { setMistakes([{ id: Date.now(), module: "待分类", source: "图片导入", title: file.name.replace(/\.[^.]+$/, ""), answer: "", mine: "", reason: "待复盘", image: String(reader.result), mastered: false }, ...mistakes]); flash("图片错题已导入"); }; reader.readAsDataURL(file); e.target.value = ""; };
  const shown = filter === "全部" ? mistakes : mistakes.filter(x => x.module === filter);
  return <div className="page-stack"><section className="page-intro"><div><p className="eyebrow">MISTAKE LIBRARY</p><h2>错题的终点，是不再错</h2><p>支持图片与 CSV 批量导入，按模块、来源和错因整理。</p></div><div className="import-actions"><input hidden ref={fileRef} type="file" accept=".csv" onChange={importCsv} /><input hidden ref={imageRef} type="file" accept="image/*" onChange={importImage} /><button className="soft-button" onClick={() => fileRef.current?.click()}>导入 CSV</button><button className="primary-button" onClick={() => imageRef.current?.click()}>导入题目图片</button></div></section>
    <section className="panel"><div className="toolbar"><div className="filter-tabs">{["全部", ...modules].map(x => <button key={x} className={filter === x ? "active" : ""} onClick={() => setFilter(x)}>{x}</button>)}</div><span>{shown.length} 道</span></div><div className="mistake-list">{shown.map(x => <article key={x.id} className={x.mastered ? "mastered" : ""}>{x.image && <img src={x.image} alt="导入的错题" />}<div className="mistake-main"><p><span>{x.module}</span><span>{x.source}</span></p><h3>{x.title}</h3><small>我的答案：{x.mine || "待补充"}　正确答案：{x.answer || "待补充"}　·　{x.reason}</small></div><button onClick={() => setMistakes(mistakes.map(y => y.id === x.id ? { ...y, mastered: !y.mastered } : y))}>{x.mastered ? "已掌握" : "标记掌握"}</button></article>)}</div></section>
    <div className="notice-card"><strong>CSV 导入格式</strong><p>首行为表头，字段顺序：模块、来源、题目、我的答案、正确答案、错因。导入前会保留原有错题。</p></div>
  </div>;
}

function ReviewView({ mistakes, setMistakes, idioms, setIdioms }: { mistakes: Mistake[]; setMistakes: (x: Mistake[]) => void; idioms: Idiom[]; setIdioms: (x: Idiom[]) => void }) {
  const pendingM = mistakes.filter(x => !x.mastered), pendingI = idioms.filter(x => x.level !== "已掌握");
  return <div className="page-stack"><section className="page-intro"><div><p className="eyebrow">DAILY REVIEW</p><h2>今日复习队列</h2><p>先清错题，再过成语。少量、多次、稳定重复。</p></div><div className="review-count"><strong>{pendingM.length + pendingI.length}</strong><span>项待复习</span></div></section>
    <section className="two-col"><div className="panel"><PanelTitle title={`错题复盘 · ${pendingM.length}`} />{pendingM.length ? pendingM.slice(0, 6).map(x => <div className="review-row" key={x.id}><span>{x.module}</span><div><strong>{x.title}</strong><small>{x.reason}</small></div><button onClick={() => setMistakes(mistakes.map(y => y.id === x.id ? { ...y, mastered: true } : y))}>掌握</button></div>) : <Empty text="今日错题已清空" />}</div>
      <div className="panel"><PanelTitle title={`成语复习 · ${pendingI.length}`} />{pendingI.length ? pendingI.slice(0, 6).map(x => <div className="review-row idiom" key={x.id}><span>{x.word.slice(0, 1)}</span><div><strong>{x.word}</strong><small>{x.meaning}</small></div><button onClick={() => setIdioms(idioms.map(y => y.id === x.id ? { ...y, level: "已掌握" } : y))}>记住了</button></div>) : <Empty text="今日成语已清空" />}</div></section>
  </div>;
}
function Empty({ text }: { text: string }) { return <div className="empty"><span>✓</span><p>{text}</p></div>; }

function ThemeSettings({ flash }: { flash: (x: string) => void }) {
  const themes = [
    { name: "奶油鼠尾草", primary: "#6f8271", accent: "#e5a48d", bg: "#f4f0e8", ink: "#3e433e" },
    { name: "蜜桃燕麦", primary: "#a76f61", accent: "#d7a85b", bg: "#f7eee6", ink: "#4d3d38" },
    { name: "雾紫可可", primary: "#7c708d", accent: "#c98f98", bg: "#f1edf3", ink: "#403b47" },
    { name: "抹茶红豆", primary: "#71805f", accent: "#a96565", bg: "#f1f0e6", ink: "#3e4339" },
  ];
  const apply = (t: typeof themes[0]) => { const r = document.documentElement; r.style.setProperty("--primary", t.primary); r.style.setProperty("--accent", t.accent); r.style.setProperty("--page", t.bg); r.style.setProperty("--ink", t.ink); localStorage.setItem("shore-theme", JSON.stringify(t)); flash(`已切换为${t.name}`); };
  const custom = (key: string, value: string) => { document.documentElement.style.setProperty(key, value); const current = { primary: getComputedStyle(document.documentElement).getPropertyValue("--primary"), accent: getComputedStyle(document.documentElement).getPropertyValue("--accent"), bg: getComputedStyle(document.documentElement).getPropertyValue("--page"), ink: getComputedStyle(document.documentElement).getPropertyValue("--ink") }; localStorage.setItem("shore-theme", JSON.stringify(current)); };
  return <div className="page-stack"><section className="page-intro"><div><p className="eyebrow">COLOR STUDIO</p><h2>把学习空间调成喜欢的样子</h2><p>低饱和、温和但不寡淡。所有选择都会保存在当前浏览器。</p></div></section><section className="theme-grid">{themes.map(t => <button key={t.name} className="theme-card" onClick={() => apply(t)}><div className="swatches"><i style={{ background: t.bg }} /><i style={{ background: t.primary }} /><i style={{ background: t.accent }} /><i style={{ background: t.ink }} /></div><strong>{t.name}</strong><small>点击应用主题</small></button>)}</section>
    <section className="panel custom-theme"><PanelTitle title="自定义色彩" /><div className="color-controls"><label><input type="color" defaultValue="#6f8271" onChange={e => custom("--primary", e.target.value)} /><span>主色</span></label><label><input type="color" defaultValue="#e5a48d" onChange={e => custom("--accent", e.target.value)} /><span>强调色</span></label><label><input type="color" defaultValue="#f4f0e8" onChange={e => custom("--page", e.target.value)} /><span>背景色</span></label><label><input type="color" defaultValue="#3e433e" onChange={e => custom("--ink", e.target.value)} /><span>文字色</span></label></div></section>
    <section className="notice-card"><strong>数据说明</strong><p>当前版本的数据保存在你的浏览器中。清理浏览器缓存前，请先从对应页面导出重要数据。</p></section>
  </div>;
}
