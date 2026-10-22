/* Builds index.html + System-Design-Atlas.xlsx from data.js — no external deps. */
const fs = require("fs");
const path = require("path");
const { SECTIONS, ROWS, STUDY_PATH } = require("./data.js");

const OUT = __dirname;
const TODAY = "2026-10-22";
const esc = (s) => String(s)
  .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
  .replace(/"/g, "&quot;");
const pad2 = (n) => String(n).padStart(2, "0");
const fmtDur = (s) => (s >= 3600 ? `${Math.floor(s / 3600)}h ${Math.round((s % 3600) / 60)}m` : `${Math.round(s / 60)} min`);

/* ------------------------------------------------------------------ HTML */
const html = `<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>The System Design Atlas — 93 concepts</title>
<meta name="description" content="93 production system-design concepts clustered into 13 areas, each mapped to a curated lecture and a hands-on lab.">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=JetBrains+Mono:wght@400;600&display=swap" rel="stylesheet">
<style>
  :root{
    --bg:#070b12; --panel:#0d1421; --panel2:#111a29; --line:#1e2a3d;
    --ink:#e8eef8; --muted:#8fa3bf; --dim:#64748b;
    --accent:#5eead4; --link:#7dd3fc;
    --radius:14px;
  }
  *{box-sizing:border-box}
  html{scroll-behavior:smooth}
  body{
    margin:0; background:var(--bg); color:var(--ink);
    font-family:Inter,system-ui,-apple-system,"Segoe UI",sans-serif;
    font-size:15px; line-height:1.5; -webkit-font-smoothing:antialiased;
  }
  body::before{
    content:""; position:fixed; inset:0; z-index:-1; pointer-events:none;
    background:
      radial-gradient(900px 500px at 12% -8%, rgba(56,189,248,.16), transparent 60%),
      radial-gradient(800px 480px at 88% 4%, rgba(167,139,250,.14), transparent 62%),
      radial-gradient(900px 600px at 50% 110%, rgba(45,212,191,.10), transparent 60%);
  }
  .wrap{max-width:1560px; margin:0 auto; padding:0 28px}

  /* ---------- hero ---------- */
  header.hero{padding:64px 0 34px}
  .eyebrow{
    display:inline-flex; align-items:center; gap:9px;
    font-family:"JetBrains Mono",monospace; font-size:11px; letter-spacing:.16em;
    text-transform:uppercase; color:var(--accent);
    border:1px solid rgba(94,234,212,.28); background:rgba(94,234,212,.07);
    padding:7px 13px; border-radius:999px;
  }
  .eyebrow .dot{width:6px;height:6px;border-radius:50%;background:var(--accent);box-shadow:0 0 12px var(--accent)}
  h1{
    margin:22px 0 0; font-size:clamp(2.6rem,6vw,4.6rem); line-height:1.02;
    letter-spacing:-.035em; font-weight:800;
  }
  h1 em{
    font-style:normal;
    background:linear-gradient(100deg,#5eead4 0%,#7dd3fc 45%,#c084fc 100%);
    -webkit-background-clip:text; background-clip:text; color:transparent;
  }
  .lede{max-width:760px; margin:20px 0 0; color:var(--muted); font-size:clamp(15px,1.4vw,17.5px)}
  .stats{display:flex; flex-wrap:wrap; gap:12px; margin-top:32px}
  .stat{
    flex:1 1 170px; background:linear-gradient(180deg,rgba(255,255,255,.055),rgba(255,255,255,.015));
    border:1px solid var(--line); border-radius:var(--radius); padding:16px 18px;
  }
  .stat b{display:block; font-size:1.9rem; line-height:1; letter-spacing:-.03em}
  .stat span{display:block; margin-top:7px; font-size:11.5px; letter-spacing:.09em;
    text-transform:uppercase; color:var(--dim); font-family:"JetBrains Mono",monospace}

  /* ---------- toolbar ---------- */
  .toolbar{
    position:sticky; top:0; z-index:40; margin-top:34px;
    background:rgba(7,11,18,.82); backdrop-filter:blur(14px) saturate(140%);
    border-top:1px solid var(--line); border-bottom:1px solid var(--line);
    padding:14px 0;
  }
  .toolbar .wrap{display:flex; align-items:center; gap:14px; flex-wrap:wrap}
  .search{position:relative; flex:1 1 300px; min-width:220px}
  .search input{
    width:100%; background:var(--panel); color:var(--ink);
    border:1px solid var(--line); border-radius:10px; padding:10px 14px 10px 38px;
    font:inherit; font-size:14px; outline:none; transition:border-color .18s, box-shadow .18s;
  }
  .search input:focus{border-color:rgba(94,234,212,.55); box-shadow:0 0 0 3px rgba(94,234,212,.12)}
  .search svg{position:absolute; left:13px; top:50%; transform:translateY(-50%); opacity:.5}
  kbd{
    font-family:"JetBrains Mono",monospace; font-size:11px; color:var(--dim);
    border:1px solid var(--line); border-bottom-width:2px; border-radius:5px; padding:1px 5px;
  }
  .chips{display:flex; gap:7px; overflow-x:auto; padding-bottom:2px; scrollbar-width:none; max-width:100%}
  .chips::-webkit-scrollbar{display:none}
  .chip{
    white-space:nowrap; cursor:pointer; font-size:12.5px; font-weight:500;
    background:transparent; color:var(--muted); border:1px solid var(--line);
    padding:7px 12px; border-radius:999px; transition:.16s; display:inline-flex; align-items:center; gap:7px;
  }
  .chip:hover{color:var(--ink); border-color:#2c3e57}
  .chip.on{background:rgba(255,255,255,.09); color:var(--ink); border-color:#38506f}
  .chip i{width:7px;height:7px;border-radius:50%;background:var(--c,#64748b); display:inline-block}
  .toggles{display:flex; align-items:center; gap:14px; margin-left:auto}
  .switch{display:inline-flex; align-items:center; gap:8px; font-size:12.5px; color:var(--muted); cursor:pointer; user-select:none}
  .switch input{accent-color:var(--accent); width:15px; height:15px}
  .progress{display:flex; align-items:center; gap:10px; font-family:"JetBrains Mono",monospace; font-size:11.5px; color:var(--dim)}
  .bar{width:110px; height:5px; border-radius:99px; background:#182436; overflow:hidden}
  .bar>div{height:100%; width:0%; background:linear-gradient(90deg,#5eead4,#7dd3fc); transition:width .25s}

  /* ---------- table ---------- */
  main{padding:8px 0 80px}
  table{width:100%; border-collapse:separate; border-spacing:0; table-layout:fixed}
  col.c-num{width:58px} col.c-topic{width:182px} col.c-used{width:21%}
  col.c-case{width:22%} col.c-vid{width:15%} col.c-prac{width:13%} col.c-course{width:11%}
  thead th{
    position:sticky; top:var(--thead-top,66px); z-index:30; text-align:left;
    background:#0a111c; border-bottom:1px solid var(--line);
    padding:12px 14px; font-size:10.5px; letter-spacing:.11em; text-transform:uppercase;
    color:var(--dim); font-weight:600; font-family:"JetBrains Mono",monospace;
  }
  tbody.sec{--c:#64748b}
  tr.sechead td{
    padding:30px 14px 12px; border-bottom:1px solid var(--line);
    background:linear-gradient(90deg, color-mix(in srgb, var(--c) 16%, transparent), transparent 65%);
  }
  .sechead .inner{display:flex; align-items:center; gap:12px}
  .sechead .badge{
    width:28px; height:28px; flex:none; border-radius:8px; display:grid; place-items:center;
    font-family:"JetBrains Mono",monospace; font-size:12px; font-weight:600; color:#04070c;
    background:var(--c); box-shadow:0 0 22px color-mix(in srgb, var(--c) 55%, transparent);
  }
  .sechead h2{margin:0; font-size:1.06rem; letter-spacing:-.015em}
  .sechead .rng{font-family:"JetBrains Mono",monospace; font-size:11px; color:var(--dim)}
  .sechead .count{margin-left:auto; font-family:"JetBrains Mono",monospace; font-size:11px; color:var(--dim)}

  tr.item td{
    padding:15px 14px; border-bottom:1px solid rgba(30,42,61,.62);
    vertical-align:top; font-size:13.4px; color:#b8c7dc;
  }
  tr.item{transition:background .14s}
  tr.item:hover td{background:rgba(255,255,255,.028)}
  tr.item.done td{opacity:.42}
  tr.item.done .topic{text-decoration:line-through; text-decoration-color:rgba(148,163,184,.7)}

  td.num{position:relative; font-family:"JetBrains Mono",monospace; font-size:12px; color:var(--dim)}
  td.num .n{display:block; margin-top:6px}
  .chk{display:block; width:15px; height:15px; cursor:pointer}
  .chk input{position:absolute; opacity:0; width:0; height:0}
  .chk span{
    display:block; width:15px; height:15px; border-radius:4px;
    border:1.5px solid #33465f; transition:.15s; position:relative;
  }
  .chk:hover span{border-color:var(--c)}
  .chk input:checked + span{background:var(--c); border-color:var(--c)}
  .chk input:checked + span::after{
    content:""; position:absolute; left:4px; top:1px; width:4px; height:8px;
    border:solid #04070c; border-width:0 2px 2px 0; transform:rotate(42deg);
  }
  td.topic{font-weight:650; color:#f1f6ff; font-size:13.8px; letter-spacing:-.01em}
  td.used,td.case{color:#9fb2cc; line-height:1.48}
  a.res{
    display:inline-flex; gap:7px; align-items:flex-start; text-decoration:none;
    color:var(--link); font-size:12.6px; line-height:1.4;
    border-bottom:1px solid transparent; transition:.15s;
  }
  a.res:hover{border-bottom-color:currentColor}
  a.res .rt{font-size:10px; opacity:.65; margin-top:2px; flex:none}
  .dur{
    display:inline-block; margin-left:5px; font-family:"JetBrains Mono",monospace;
    font-size:9.5px; color:#64748b; background:rgba(255,255,255,.05);
    border:1px solid rgba(255,255,255,.07); border-radius:4px; padding:1px 4px;
  }
  .course a{text-decoration:none; display:inline-block}
  .course .code{
    display:inline-block; font-family:"JetBrains Mono",monospace; font-size:11.5px;
    color:#c7d2fe; border:1px solid #2b3a55; background:rgba(99,102,241,.10);
    padding:3px 8px; border-radius:6px; transition:.15s;
  }
  .course a:hover .code{border-color:#4c5f85; background:rgba(99,102,241,.2)}
  .course .rest{display:block; margin-top:5px; font-size:11px; color:var(--dim); line-height:1.35}
  .none{color:#3f4f66}

  /* ---------- footer ---------- */
  footer{border-top:1px solid var(--line); padding:36px 0 64px; color:var(--dim); font-size:12.5px}
  footer .grid{display:grid; grid-template-columns:repeat(auto-fit,minmax(260px,1fr)); gap:26px; margin-bottom:26px}
  footer h3{margin:0 0 10px; font-size:11px; letter-spacing:.12em; text-transform:uppercase;
    color:var(--muted); font-family:"JetBrains Mono",monospace}
  footer ol{margin:0; padding-left:18px; line-height:1.75}
  footer ol b{color:var(--ink); font-weight:600}
  .path{background:var(--panel); border:1px solid var(--line); border-radius:var(--radius); padding:20px 22px}
  .empty{padding:60px 14px; text-align:center; color:var(--dim); font-family:"JetBrains Mono",monospace; font-size:13px}

  @media (max-width:1080px){
    thead th{top:0}
    col.c-used,col.c-case{width:auto}
    table{min-width:1120px}
    main .wrap{overflow-x:auto; padding-bottom:10px}
  }
  @media print{
    body{background:#fff; color:#111}
    body::before,.toolbar,footer .path{display:none}
    header.hero{padding:0 0 18px}
    h1{font-size:26pt} h1 em{color:#111;background:none;-webkit-text-fill-color:#111}
    .lede{color:#444}
    .stat{border-color:#ddd;background:#fff}
    table{min-width:0; font-size:8pt}
    thead th{position:static; background:#f4f4f4; color:#333}
    tr.item td{color:#222; border-color:#e5e5e5}
    tr.item{break-inside:avoid}
    tr.sechead td{background:#f0f0f0}
    a.res,a{color:#0645ad}
    .chk{display:none}
  }
</style>
</head>
<body>

<header class="hero">
  <div class="wrap">
    <div class="eyebrow"><span class="dot"></span>Learning atlas · ${ROWS.length} concepts</div>
    <h1>The System Design <em>Atlas</em></h1>
    <p class="lede">Every concept from the list, clustered the way it actually shows up in production — with one curated lecture and one place to practise it. Tick a row when it clicks; your progress is saved in this browser.</p>
    <div class="stats">
      <div class="stat"><b>${ROWS.length}</b><span>Concepts</span></div>
      <div class="stat"><b>${SECTIONS.length}</b><span>Clusters</span></div>
      <div class="stat"><b>${new Set(ROWS.flatMap(r => [r.v[1], r.p[1], r.k && r.k[1]]).filter(Boolean)).size}</b><span>Verified resources</span></div>
      <div class="stat"><b id="pctStat">0%</b><span>Completed</span></div>
    </div>
  </div>
</header>

<div class="toolbar">
  <div class="wrap">
    <div class="search">
      <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="11" cy="11" r="7"/><path d="m20 20-3.5-3.5"/></svg>
      <input id="q" type="search" placeholder="Search concepts, e.g. “sharding”, “webhook”, “probe”…" autocomplete="off">
    </div>
    <div class="chips" id="chips">
      <button class="chip on" data-f="all">All</button>
      ${SECTIONS.map(s => `<button class="chip" data-f="${s.n}" style="--c:${s.color}"><i></i>${esc(s.name)}</button>`).join("")}
    </div>
    <div class="toggles">
      <label class="switch"><input type="checkbox" id="hideDone"> hide completed</label>
      <div class="progress"><span id="cnt">0/${ROWS.length}</span><div class="bar"><div id="barFill"></div></div></div>
    </div>
  </div>
</div>

<main>
  <div class="wrap">
    <table>
      <colgroup>
        <col class="c-num"><col class="c-topic"><col class="c-used"><col class="c-case">
        <col class="c-vid"><col class="c-prac"><col class="c-course">
      </colgroup>
      <thead>
        <tr>
          <th>#</th><th>Topic</th><th>Used for</th><th>Use case</th>
          <th>Video lecture · direct</th><th>Practise it</th><th>MIT / Stanford</th>
        </tr>
      </thead>
      ${SECTIONS.map(sec => {
        const items = ROWS.filter(r => r.s === sec.n);
        return `<tbody class="sec" data-s="${sec.n}" style="--c:${sec.color}">
        <tr class="sechead"><td colspan="7">
          <div class="inner">
            <div class="badge">${pad2(sec.n)}</div>
            <h2>${esc(sec.name)}</h2>
            <span class="rng">${esc(sec.range)}</span>
            <span class="count">${items.length} topics</span>
          </div>
        </td></tr>
        ${items.map(r => {
          const [vc, vr] = r.v;
          const [pc, pr] = r.p;
          const course = r.k ? `<a href="${esc(r.k[1])}" target="_blank" rel="noopener" title="${esc(r.k[0])}"><span class="code">${esc(r.k[0].split(" · ")[0])}</span><span class="rest">${esc(r.k[0].split(" · ").slice(1).join(" · ") || "course")}</span></a>` : `<span class="none">—</span>`;
          return `<tr class="item" data-n="${r.n}" data-s="${sec.n}">
            <td class="num"><label class="chk"><input type="checkbox" data-n="${r.n}"><span></span></label><span class="n">${pad2(r.n)}</span></td>
            <td class="topic">${esc(r.t)}</td>
            <td class="used">${esc(r.u)}</td>
            <td class="case">${esc(r.c)}</td>
            <td class="res"><a class="res" href="${esc(vr)}" target="_blank" rel="noopener"><span class="rt">▶</span>${esc(vc)}</a>${r.v[2] ? `<span class="dur">${fmtDur(r.v[2])}</span>` : ""}</td>
            <td class="res"><a class="res" href="${esc(pr)}" target="_blank" rel="noopener"><span class="rt">⌘</span>${esc(pc)}</a></td>
            <td class="course">${course}</td>
          </tr>`;
        }).join("")}
      </tbody>`;
      }).join("")}
    </table>
    <div class="empty" id="empty" style="display:none">no concepts match that filter</div>
  </div>
</main>

<footer>
  <div class="wrap">
    <div class="path">
      <h3>If you only do six things</h3>
      <ol>
        ${STUDY_PATH.map(([t, d, r]) => `<li><b>${esc(t)}</b> — ${esc(d)} <span class="none">(${esc(r)})</span></li>`).join("")}
      </ol>
    </div>
    <div class="grid" style="margin-top:26px">
      <div>
        <h3>How this was built</h3>
        Every row links to one specific YouTube lecture — no channel homepages, no invented video IDs. IDs were pulled from YouTube search and course playlists, then resolved individually. Live-checked ${TODAY}.
      </div>
      <div>
        <h3>Reading the columns</h3>
        <b>Used for</b> = the mechanism. <b>Use case</b> = a concrete scenario. <b>MIT / Stanford</b> = the closest formal course; an em dash means none covers it well.
      </div>
      <div>
        <h3>Companion file</h3>
        <code style="color:#c7d2fe">System-Design-Atlas.xlsx</code> — same dataset, sortable and filterable, for note-taking offline.
      </div>
    </div>
  </div>
</footer>

<script>
(function(){
  var KEY = "sda-done-v1";
  var done = new Set();
  try { done = new Set(JSON.parse(localStorage.getItem(KEY) || "[]")); } catch(e) {}

  var q = document.getElementById("q");
  var hideDone = document.getElementById("hideDone");
  var chips = Array.prototype.slice.call(document.querySelectorAll(".chip"));
  var items = Array.prototype.slice.call(document.querySelectorAll("tr.item"));
  var secs = Array.prototype.slice.call(document.querySelectorAll("tbody.sec"));
  var cnt = document.getElementById("cnt");
  var fill = document.getElementById("barFill");
  var pctStat = document.getElementById("pctStat");
  var empty = document.getElementById("empty");
  var filter = "all";

  function save(){ try { localStorage.setItem(KEY, JSON.stringify(Array.from(done))); } catch(e) {} }

  function paintProgress(){
    var n = done.size, total = items.length, pct = Math.round(n/total*100);
    cnt.textContent = n + "/" + total;
    fill.style.width = pct + "%";
    pctStat.textContent = pct + "%";
  }

  function apply(){
    var term = q.value.trim().toLowerCase();
    var visible = 0;
    items.forEach(function(tr){
      var n = tr.getAttribute("data-n");
      var ok = (filter === "all" || tr.getAttribute("data-s") === filter)
            && (!term || tr.textContent.toLowerCase().indexOf(term) !== -1)
            && (!hideDone.checked || !done.has(n));
      tr.style.display = ok ? "" : "none";
      if (ok) visible++;
      tr.classList.toggle("done", done.has(n));
    });
    secs.forEach(function(tb){
      var any = Array.prototype.some.call(tb.querySelectorAll("tr.item"), function(tr){ return tr.style.display !== "none"; });
      tb.style.display = any ? "" : "none";
    });
    empty.style.display = visible ? "none" : "";
  }

  document.querySelectorAll(".chk input").forEach(function(cb){
    var n = cb.getAttribute("data-n");
    cb.checked = done.has(n);
    cb.addEventListener("change", function(){
      if (cb.checked) done.add(n); else done.delete(n);
      save(); paintProgress(); apply();
    });
  });

  chips.forEach(function(ch){
    ch.addEventListener("click", function(){
      chips.forEach(function(c){ c.classList.remove("on"); });
      ch.classList.add("on");
      filter = ch.getAttribute("data-f");
      apply();
    });
  });

  q.addEventListener("input", apply);
  hideDone.addEventListener("change", apply);
  document.addEventListener("keydown", function(e){
    if (e.key === "/" && document.activeElement !== q){ e.preventDefault(); q.focus(); }
    if (e.key === "Escape"){ q.value = ""; apply(); q.blur(); }
  });

  function syncOffset(){
    var tb = document.querySelector(".toolbar");
    if (tb) document.documentElement.style.setProperty("--thead-top", tb.offsetHeight + "px");
  }
  window.addEventListener("resize", syncOffset);

  syncOffset();
  paintProgress(); apply();
})();
</script>
</body>
</html>`;

fs.writeFileSync(path.join(OUT, "index.html"), html, "utf8");

/* ------------------------------------------------------------------ XLSX */
/* Minimal OOXML writer: ZIP (stored) + SpreadsheetML. No dependencies. */

const CRC = (() => {
  const t = new Int32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
    t[n] = c;
  }
  return t;
})();
function crc32(buf) {
  let c = -1;
  for (let i = 0; i < buf.length; i++) c = CRC[(c ^ buf[i]) & 0xFF] ^ (c >>> 8);
  return (c ^ -1) >>> 0;
}
function zipStore(files) {
  const now = new Date(2026, 9, 22, 12, 0, 0);
  const dosTime = ((now.getHours() << 11) | (now.getMinutes() << 5) | (now.getSeconds() >> 1)) & 0xFFFF;
  const dosDate = (((now.getFullYear() - 1980) << 9) | ((now.getMonth() + 1) << 5) | now.getDate()) & 0xFFFF;
  const locals = [], central = [];
  let offset = 0;
  for (const f of files) {
    const name = Buffer.from(f.name, "utf8");
    const data = Buffer.isBuffer(f.data) ? f.data : Buffer.from(f.data, "utf8");
    const crc = crc32(data);
    const lh = Buffer.alloc(30);
    lh.writeUInt32LE(0x04034b50, 0); lh.writeUInt16LE(20, 4); lh.writeUInt16LE(0x0800, 6);
    lh.writeUInt16LE(0, 8); lh.writeUInt16LE(dosTime, 10); lh.writeUInt16LE(dosDate, 12);
    lh.writeUInt32LE(crc, 14); lh.writeUInt32LE(data.length, 18); lh.writeUInt32LE(data.length, 22);
    lh.writeUInt16LE(name.length, 26); lh.writeUInt16LE(0, 28);
    locals.push(lh, name, data);
    const ch = Buffer.alloc(46);
    ch.writeUInt32LE(0x02014b50, 0); ch.writeUInt16LE(20, 4); ch.writeUInt16LE(20, 6);
    ch.writeUInt16LE(0x0800, 8); ch.writeUInt16LE(0, 10); ch.writeUInt16LE(dosTime, 12); ch.writeUInt16LE(dosDate, 14);
    ch.writeUInt32LE(crc, 16); ch.writeUInt32LE(data.length, 20); ch.writeUInt32LE(data.length, 24);
    ch.writeUInt16LE(name.length, 28); ch.writeUInt16LE(0, 30); ch.writeUInt16LE(0, 32);
    ch.writeUInt16LE(0, 34); ch.writeUInt16LE(0, 36); ch.writeUInt32LE(0, 38); ch.writeUInt32LE(offset, 42);
    central.push(ch, name);
    offset += lh.length + name.length + data.length;
  }
  const cd = Buffer.concat(central);
  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50, 0); eocd.writeUInt16LE(0, 4); eocd.writeUInt16LE(0, 6);
  eocd.writeUInt16LE(files.length, 8); eocd.writeUInt16LE(files.length, 10);
  eocd.writeUInt32LE(cd.length, 12); eocd.writeUInt32LE(offset, 16); eocd.writeUInt16LE(0, 20);
  return Buffer.concat([...locals, cd, eocd]);
}

const SEC_HEX = SECTIONS.map(s => s.color.replace("#", "").toUpperCase());
const xesc = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;")
  .replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&apos;");

/* --- styles --- */
const sectionFills = SEC_HEX.map((c, i) => `<fill><patternFill patternType="solid"><fgColor rgb="FF${c}"/><bgColor indexed="64"/></patternFill></fill>`).join("");
const sectionXfs = SEC_HEX.map((c, i) => `<xf numFmtId="0" fontId="7" fillId="${3 + i}" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center" indent="1"/></xf>`).join("");

const styles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="11">
<font><sz val="10"/><color rgb="FF1F2937"/><name val="Calibri"/></font>
<font><b/><sz val="9"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
<font><b/><sz val="20"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
<font><sz val="9"/><color rgb="FF94A3B8"/><name val="Calibri"/></font>
<font><b/><sz val="10"/><color rgb="FF0F172A"/><name val="Calibri"/></font>
<font><sz val="9.5"/><color rgb="FF334155"/><name val="Calibri"/></font>
<font><u/><sz val="9"/><color rgb="FF2563EB"/><name val="Calibri"/></font>
<font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
<font><b/><sz val="8.5"/><color rgb="FF1E40AF"/><name val="Calibri"/></font>
<font><sz val="8"/><color rgb="FF64748B"/><name val="Calibri"/></font>
<font><b/><sz val="9"/><color rgb="FF94A3B8"/><name val="Calibri"/></font>
</fonts>
<fills count="17">
<fill><patternFill patternType="none"/></fill>
<fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF0F172A"/><bgColor indexed="64"/></patternFill></fill>
${sectionFills}
<fill><patternFill patternType="solid"><fgColor rgb="FFFFFFFF"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFF1F5F9"/><bgColor indexed="64"/></patternFill></fill>
</fills>
<borders count="2">
<border><left/><right/><top/><bottom/><diagonal/></border>
<border><left/><right/><top/><bottom style="thin"><color rgb="FFE2E8F0"/></bottom><diagonal/></border>
</borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="${4 + SEC_HEX.length + 10}">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="2" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="left" vertical="center" indent="1"/></xf>
<xf numFmtId="0" fontId="3" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="left" vertical="center" indent="1"/></xf>
<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
${sectionXfs}
<xf numFmtId="0" fontId="5" fillId="16" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="top" wrapText="1"/></xf>
<xf numFmtId="0" fontId="5" fillId="17" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="top" wrapText="1"/></xf>
<xf numFmtId="0" fontId="4" fillId="16" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="top" wrapText="1"/></xf>
<xf numFmtId="0" fontId="4" fillId="17" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="top" wrapText="1"/></xf>
<xf numFmtId="0" fontId="6" fillId="16" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="top" wrapText="1"/></xf>
<xf numFmtId="0" fontId="6" fillId="17" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="top" wrapText="1"/></xf>
<xf numFmtId="0" fontId="10" fillId="16" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="top"/></xf>
<xf numFmtId="0" fontId="10" fillId="17" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="top"/></xf>
<xf numFmtId="0" fontId="8" fillId="16" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="top" wrapText="1"/></xf>
<xf numFmtId="0" fontId="8" fillId="17" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="top" wrapText="1"/></xf>
<xf numFmtId="0" fontId="9" fillId="16" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="top" wrapText="1"/></xf>
<xf numFmtId="0" fontId="9" fillId="17" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="top" wrapText="1"/></xf>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`;

/* style ids */
const S_TITLE = 1, S_SUB = 2, S_HEAD = 3, S_SEC0 = 4;
const S_BODY = S_SEC0 + SEC_HEX.length;        // 17
const S_BODY_ALT = S_BODY + 1, S_TOPIC = S_BODY + 2, S_TOPIC_ALT = S_BODY + 3;
const S_LINK = S_BODY + 4, S_LINK_ALT = S_BODY + 5, S_NUM = S_BODY + 6, S_NUM_ALT = S_BODY + 7;
const S_COURSE = S_BODY + 8, S_COURSE_ALT = S_BODY + 9;

const WIDTHS = [5, 24, 44, 48, 34, 32, 30];
const cols = WIDTHS.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join("");

let sheetRows = [], merges = [], links = [], rels = [], rowNum = 0;
const cells = [];
function cell(ref, style, value, isNum) {
  if (isNum) return `<c r="${ref}" s="${style}"><v>${value}</v></c>`;
  if (value === null || value === undefined || value === "") return `<c r="${ref}" s="${style}"/>`;
  return `<c r="${ref}" s="${style}" t="inlineStr"><is><t xml:space="preserve">${xesc(value)}</t></is></c>`;
}
function estimatedHeight(vals) {
  const cpw = [24, 44, 48, 34, 32, 30];
  let max = 1;
  vals.forEach((v, i) => {
    if (!v) return;
    const width = cpw[i] || 30;
    const lines = String(v).split("\n").reduce((acc, seg) => acc + Math.ceil(seg.length / (width * 1.06)), 0);
    max = Math.max(max, lines);
  });
  return Math.min(128, Math.max(24, max * 11.6 + 6));
}

// Row 1 — title
rowNum = 1;
sheetRows.push(`<row r="${rowNum}" ht="36" customHeight="1"><c r="A1" s="${S_TITLE}" t="inlineStr"><is><t>The System Design Atlas</t></is></c></row>`);
merges.push("A1:G1");
// Row 2 — subtitle
rowNum = 2;
sheetRows.push(`<row r="${rowNum}" ht="20" customHeight="1"><c r="A2" s="${S_SUB}" t="inlineStr"><is><t>${xesc(`${ROWS.length} concepts · ${SECTIONS.length} clusters · best-rated lecture + hands-on lab + MIT/Stanford course for each · generated ${TODAY}`)}</t></is></c></row>`);
merges.push("A2:G2");
// Row 3 — spacer
rowNum = 3;
sheetRows.push(`<row r="${rowNum}" ht="8" customHeight="1"/>`);
// Row 4 — header
rowNum = 4;
const HEAD = ["#", "Topic", "Used for", "Use case", "Video lecture · direct link", "Practise it (hands-on)", "MIT OCW / Stanford"];
sheetRows.push(`<row r="${rowNum}" ht="30" customHeight="1">` + HEAD.map((h, i) =>
  `<c r="${String.fromCharCode(65 + i)}${rowNum}" s="${S_HEAD}" t="inlineStr"><is><t>${xesc(h)}</t></is></c>`).join("") + `</row>`);

let alt = 0;
for (const sec of SECTIONS) {
  const items = ROWS.filter(r => r.s === sec.n);
  rowNum++;
  sheetRows.push(`<row r="${rowNum}" ht="26" customHeight="1"><c r="A${rowNum}" s="${S_SEC0 + (sec.n - 1)}" t="inlineStr"><is><t>${xesc(`${pad2(sec.n)}  ${sec.name}   ·   ${sec.range}   ·   ${items.length} topics`)}</t></is></c></row>`);
  merges.push(`A${rowNum}:G${rowNum}`);

  for (const r of items) {
    rowNum++;
    const zebra = alt++ % 2 === 1;
    const b = zebra ? S_BODY_ALT : S_BODY;
    const t = zebra ? S_TOPIC_ALT : S_TOPIC;
    const n = zebra ? S_NUM_ALT : S_NUM;
    const c = zebra ? S_COURSE_ALT : S_COURSE;

    // video link
    links.push({ ref: `E${rowNum}`, target: r.v[1] });
    rels.push(`<Relationship Id="rId${links.length}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${xesc(r.v[1])}" TargetMode="External"/>`);
    const l1 = zebra ? S_LINK_ALT : S_LINK;

    links.push({ ref: `F${rowNum}`, target: r.p[1] });
    rels.push(`<Relationship Id="rId${links.length}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${xesc(r.p[1])}" TargetMode="External"/>`);

    let courseCell;
    if (r.k) {
      links.push({ ref: `G${rowNum}`, target: r.k[1] });
      rels.push(`<Relationship Id="rId${links.length}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${xesc(r.k[1])}" TargetMode="External"/>`);
      courseCell = cell(`G${rowNum}`, c, r.k[0]);
    } else {
      courseCell = `<c r="G${rowNum}" s="${c}" t="inlineStr"><is><t>—</t></is></c>`;
    }

    const h = estimatedHeight([String(r.n), r.t, r.u, r.c, r.v[0], r.p[0], r.k ? r.k[0] : "—"]);
    sheetRows.push(`<row r="${rowNum}" ht="${h.toFixed(1)}" customHeight="1">` +
      cell(`A${rowNum}`, n, r.n, true) +
      cell(`B${rowNum}`, t, r.t) +
      cell(`C${rowNum}`, b, r.u) +
      cell(`D${rowNum}`, b, r.c) +
      cell(`E${rowNum}`, l1, r.v[0] + (r.v[2] ? ` (${fmtDur(r.v[2])})` : "")) +
      cell(`F${rowNum}`, l1, r.p[0]) +
      courseCell +
      `</row>`);
  }
}

const hyperlinks = links.length
  ? `<hyperlinks>${links.map((l, i) => `<hyperlink ref="${l.ref}" r:id="rId${i + 1}"/>`).join("")}</hyperlinks>`
  : "";

const sheet = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>
<dimension ref="A1:G${rowNum}"/>
<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane ySplit="4" topLeftCell="A5" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A5" sqref="A5"/></sheetView></sheetViews>
<sheetFormatPr defaultRowHeight="15"/>
<cols>${cols}</cols>
<sheetData>${sheetRows.join("")}</sheetData>
<mergeCells count="${merges.length}">${merges.map(m => `<mergeCell ref="${m}"/>`).join("")}</mergeCells>
${hyperlinks}
<pageMargins left="0.4" right="0.4" top="0.5" bottom="0.5" header="0.3" footer="0.3"/>
<pageSetup orientation="landscape" fitToWidth="1" fitToHeight="0" paperSize="9"/>
</worksheet>`;

const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`;

const rootRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;

const workbook = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="Atlas" sheetId="1" r:id="rId1"/></sheets>
</workbook>`;

const wbRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;

const sheetRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${rels.join("\n")}
</Relationships>`;

const xlsx = zipStore([
  { name: "[Content_Types].xml", data: contentTypes },
  { name: "_rels/.rels", data: rootRels },
  { name: "xl/workbook.xml", data: workbook },
  { name: "xl/_rels/workbook.xml.rels", data: wbRels },
  { name: "xl/styles.xml", data: styles },
  { name: "xl/worksheets/sheet1.xml", data: sheet },
  { name: "xl/worksheets/_rels/sheet1.xml.rels", data: sheetRels }
]);

fs.writeFileSync(path.join(OUT, "System-Design-Atlas.xlsx"), xlsx);

console.log("index.html            ", (Buffer.byteLength(html) / 1024).toFixed(1) + " KB");
console.log("System-Design-Atlas.xlsx", (xlsx.length / 1024).toFixed(1) + " KB");
console.log("rows:", ROWS.length, "| sections:", SECTIONS.length, "| hyperlinks:", links.length, "| sheet rows:", rowNum);
