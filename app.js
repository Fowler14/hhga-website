/* HHGA site — single shared script; page chosen by <body data-page="...">. */
"use strict";

const $ = (sel, el) => (el || document).querySelector(sel);

const state = { data: null, names: {}, players: {} };

async function loadData() {
  const res = await fetch("data/hhga.json");
  state.data = await res.json();
  for (const p of state.data.players) {
    state.names[p.id] = p.name;
    state.players[p.id] = p;
  }
  return state.data;
}

const nameOf = (pid) => state.names[pid] || pid;
const isPhantom = (pid) => !!(state.players[pid] && state.players[pid].phantom);

function fmtDates(yr) {
  const ds = yr.rounds.map((r) => r.date).filter(Boolean).sort();
  if (!ds.length) return "";
  const opt = { month: "short", day: "numeric" };
  const a = new Date(ds[0] + "T12:00:00");
  const b = new Date(ds[ds.length - 1] + "T12:00:00");
  return `${a.toLocaleDateString("en-US", opt)}–${b.toLocaleDateString("en-US", opt)}, ${yr.year}`;
}

function championScore(yr) {
  const top = yr.leaderboard[0];
  return yr.scoring === "gross" ? `${top.gross} (gross)` : `net ${top.net}`;
}

/* ---------------------------------------------------------------- home */

function renderHome() {
  const d = state.data;
  const years = d.years;
  const latest = years[years.length - 1];

  $("#champ-name").textContent = nameOf(latest.champion);
  $("#champ-detail").textContent =
    `${latest.year} · ${latest.location} · ${championScore(latest)}`;

  // titles
  const titles = {};
  for (const y of years) titles[y.champion] = (titles[y.champion] || 0) + 1;
  const n = titles[latest.champion];
  $("#champ-titles").textContent = n > 1 ? `${ordinal(n)} title` : "First title";

  // stats
  let strokes = 0, holes = 0;
  const everyone = new Set();
  for (const y of years) {
    for (const r of y.rounds) {
      for (const [pid, hs] of Object.entries(r.scores)) {
        strokes += hs.reduce((a, b) => a + b, 0);
        holes += hs.length;
        everyone.add(pid);
      }
    }
  }
  $("#stat-years").textContent = years.length;
  $("#stat-strokes").textContent = strokes.toLocaleString();
  $("#stat-holes").textContent = holes.toLocaleString();
  $("#stat-players").textContent = [...everyone].filter((p) => !isPhantom(p)).length;

  // champions wall (newest first)
  const wall = $("#wall");
  for (const y of [...years].reverse()) {
    const a = document.createElement("a");
    a.href = `years.html#${y.year}`;
    a.innerHTML = `<div class="yr">${y.year} · ${esc(shortLoc(y.location))}</div>
                   <div class="who">${esc(nameOf(y.champion))}</div>`;
    wall.appendChild(a);
  }

  // mist-weed running tally
  let mist = 0, weed = 0;
  for (const y of years) {
    const mw = y.sideGames && y.sideGames.mistWeed;
    if (mw && mw.winner === "mist") mist++;
    else if (mw && mw.winner === "weed") weed++;
  }
  $("#mw-tally").innerHTML =
    `<span class="mw-team mist">Mist ${mist}</span> &nbsp;–&nbsp; <span class="mw-team weed">Weed ${weed}</span>`;
}

/* ---------------------------------------------------------------- years */

function renderYears() {
  const years = state.data.years;
  const sel = $("#year-select");
  for (const y of [...years].reverse()) {
    const o = document.createElement("option");
    o.value = y.year;
    o.textContent = `${y.year} — ${shortLoc(y.location)}`;
    sel.appendChild(o);
  }
  const fromHash = parseInt(location.hash.slice(1), 10);
  const start = years.some((y) => y.year === fromHash) ? fromHash : years[years.length - 1].year;
  sel.value = start;
  sel.addEventListener("change", () => {
    location.hash = sel.value;
    showYear(parseInt(sel.value, 10));
  });
  window.addEventListener("hashchange", () => {
    const y = parseInt(location.hash.slice(1), 10);
    if (years.some((v) => v.year === y) && parseInt(sel.value, 10) !== y) {
      sel.value = y;
      showYear(y);
    }
  });
  showYear(start);
}

function showYear(yearNum) {
  const yr = state.data.years.find((y) => y.year === yearNum);
  const out = $("#year-view");
  out.innerHTML = "";

  const head = document.createElement("div");
  head.innerHTML = `
    <h2>${yr.year} — ${esc(yr.location)}</h2>
    <p class="muted">${fmtDates(yr)} · ${yr.rounds.length} rounds ·
      ${yr.leaderboard.length} players · scoring: ${yr.scoring}</p>
    <div class="chips">${yr.courses.map((c) => `<span class="chip">${esc(c)}</span>`).join("")}</div>`;
  out.appendChild(head);

  // leaderboard
  const lbSec = document.createElement("div");
  const hasNet = yr.scoring !== "gross";
  lbSec.innerHTML = `<h3>Leaderboard</h3>`;
  const rows = yr.leaderboard.map((e) => {
    const extra = e.roundsPlayed ? ` <span class="muted">(${e.roundsPlayed} rounds)</span>` : "";
    return `<tr class="${e.place === 1 ? "first" : ""}">
      <td>${e.place ?? "—"}</td>
      <td><a href="players.html#${e.player}">${esc(nameOf(e.player))}</a>${e.place === 1 ? " \u{1F3C6}" : ""}${extra}</td>
      <td>${e.gross ?? "—"}</td>
      ${hasNet ? `<td>${e.net ?? "—"}</td>` : ""}
    </tr>`;
  }).join("");
  lbSec.innerHTML += `<table class="data"><thead><tr>
      <th>Place</th><th>Player</th><th>Gross</th>${hasNet ? "<th>Net</th>" : ""}
    </tr></thead><tbody>${rows}</tbody></table>`;
  out.appendChild(lbSec);

  // mist-weed
  const mw = yr.sideGames && yr.sideGames.mistWeed;
  if (mw && mw.final) renderMistWeed(out, mw);

  // myrtle cup 2008
  const mc = yr.sideGames && yr.sideGames.myrtleCup;
  if (mc) renderMyrtleCup(out, mc);

  // winnings
  const win = yr.sideGames && yr.sideGames.winnings;
  if (win) {
    const sec = document.createElement("div");
    sec.innerHTML = `<h3>Money</h3><table class="data"><thead>
      <tr><th>Player</th><th>Winnings</th></tr></thead><tbody>
      ${Object.entries(win).sort((a, b) => b[1] - a[1]).map(([p, v]) =>
        `<tr><td>${esc(nameOf(p))}</td><td>${v < 0 ? "−" : ""}$${Math.abs(v).toFixed(2)}</td></tr>`).join("")}
      </tbody></table>`;
    out.appendChild(sec);
  }

  // scorecards
  const sc = document.createElement("div");
  sc.innerHTML = "<h3>Scorecards</h3>";
  for (const r of yr.rounds) sc.appendChild(scorecard(yr, r));
  out.appendChild(sc);
}

function renderMistWeed(out, mw) {
  const sec = document.createElement("div");
  const f = mw.final;
  const winBadge = mw.winner === "tie" ? "Tie" : `${mw.winner} wins`;
  sec.innerHTML = `<h3>Mist vs Weed</h3>
    <div class="mw-final">
      <span class="mw-team mist">Mist</span>
      <span class="mw-score">${f.mist}–${f.weed}</span>
      <span class="mw-team weed">Weed</span>
      <span class="badge">${esc(winBadge)}</span>
    </div>`;
  if (mw.teams) {
    sec.innerHTML += `<p class="muted">
      <b>Mist:</b> ${mw.teams.mist.map(nameOf).map(esc).join(", ")}<br>
      <b>Weed:</b> ${mw.teams.weed.map(nameOf).map(esc).join(", ")}</p>`;
  }
  if (mw.days && mw.days.length) {
    sec.innerHTML += `<table class="data"><thead><tr><th>Day</th><th>Mist</th><th>Weed</th></tr></thead>
      <tbody>${mw.days.map((d) =>
        `<tr><td>Day ${d.day}</td><td>${d.mistPoints ?? ""}</td><td>${d.weedPoints ?? ""}</td></tr>`).join("")}
      </tbody></table>`;
  }
  if (mw.individualRecords) {
    const recs = Object.entries(mw.individualRecords)
      .filter(([, r]) => r && r.wins !== undefined)
      .sort((a, b) => (b[1].wins - b[1].losses) - (a[1].wins - a[1].losses));
    if (recs.length) {
      sec.innerHTML += `<details><summary class="muted">Individual records</summary>
        <table class="data"><thead><tr><th>Player</th><th>W</th><th>L</th><th>T</th></tr></thead><tbody>
        ${recs.map(([p, r]) => `<tr><td>${esc(nameOf(p))}${isPhantom(p) ? " \u{1F47B}" : ""}</td>
          <td>${r.wins}</td><td>${r.losses}</td><td>${r.ties}</td></tr>`).join("")}
        </tbody></table></details>`;
    }
  }
  out.appendChild(sec);
}

function renderMyrtleCup(out, mc) {
  const sec = document.createElement("div");
  sec.innerHTML = `<h3>${esc(mc.name)} (Mist vs Weed origins)</h3>
    <p class="muted">${esc(mc.format)}</p>
    <div class="mw-final">
      <span class="mw-team mist">Mist</span>
      <span class="mw-score">${mc.standingsAfterWed.mist}–${mc.standingsAfterWed.weed}</span>
      <span class="mw-team weed">Weed</span>
      <span class="badge">after Wed</span>
    </div>
    <p class="muted"><b>Mist:</b> ${mc.teams.mist.map(nameOf).map(esc).join(", ")}<br>
       <b>Weed:</b> ${mc.teams.weed.map(nameOf).map(esc).join(", ")}</p>
    <p class="muted">${esc(mc.resultNote)}</p>`;
  out.appendChild(sec);
}

function scorecard(yr, r) {
  const div = document.createElement("div");
  const meta = [];
  if (r.rating) meta.push(`rating ${r.rating}`);
  if (r.slope) meta.push(`slope ${r.slope}`);
  const holes = r.par.length;
  const nums = Array.from({ length: holes }, (_, i) => i + 1);

  // order players by leaderboard
  const order = yr.leaderboard.map((e) => e.player).filter((p) => r.scores[p]);

  const rowsHtml = order.map((pid) => {
    const hs = r.scores[pid];
    const hcp = r.handicaps ? r.handicaps[pid] : (yr.handicaps ? yr.handicaps[pid] : null);
    const cells = hs.map((s, i) => {
      const d = s - r.par[i];
      const cls = d <= -2 ? "eagle" : d === -1 ? "under" : "";
      return `<td class="${cls}">${s}</td>`;
    }).join("");
    const label = hcp != null ? `${esc(nameOf(pid))} <span class="muted">(${hcp})</span>` : esc(nameOf(pid));
    return `<tr><td>${label}</td>${cells}<td><b>${hs.reduce((a, b) => a + b, 0)}</b></td></tr>`;
  }).join("");

  div.className = "scorewrap";
  div.innerHTML = `
    <h4>Round ${r.round} · ${esc(r.course || "")}
      <span class="muted">${r.date || ""}${meta.length ? " · " + meta.join(", ") : ""}</span></h4>
    <table class="data scorecard">
      <thead><tr><th></th>${nums.map((n) => `<th>${n}</th>`).join("")}<th>Tot</th></tr></thead>
      <tbody>
        <tr class="par"><td>Par</td>${r.par.map((p) => `<td>${p}</td>`).join("")}
          <td>${r.par.reduce((a, b) => a + b, 0)}</td></tr>
        ${rowsHtml}
      </tbody>
    </table>`;
  return div;
}

/* -------------------------------------------------------------- players */

function playerCareer(pid) {
  const seasons = [];
  for (const y of state.data.years) {
    const e = y.leaderboard.find((v) => v.player === pid);
    if (e) seasons.push({ year: y.year, place: e.place, of: y.leaderboard.length,
                          gross: e.gross, net: e.net, champion: y.champion === pid });
  }
  return seasons;
}

function renderPlayers() {
  const roster = $("#roster");
  const players = state.data.players
    .filter((p) => !p.phantom)
    .filter((p) => playerCareer(p.id).length > 0)
    .sort((a, b) => a.name.localeCompare(b.name));

  for (const p of players) {
    const c = playerCareer(p.id);
    const wins = c.filter((s) => s.champion).length;
    const a = document.createElement("a");
    a.href = `#${p.id}`;
    a.id = `card-${p.id}`;
    a.innerHTML = `<div class="who">${esc(p.name)}</div>
      <div class="sub">${c.length} trips${wins ? ` · ${wins} \u{1F3C6}` : ""}</div>`;
    a.addEventListener("click", () => setTimeout(() => showPlayer(p.id), 0));
    roster.appendChild(a);
  }

  const fromHash = location.hash.slice(1);
  showPlayer(players.some((p) => p.id === fromHash) ? fromHash : players[0].id);
}

let chart = null;

function showPlayer(pid) {
  document.querySelectorAll(".roster a").forEach((a) => a.classList.remove("sel"));
  const card = $(`#card-${pid}`);
  if (card) card.classList.add("sel");

  const c = playerCareer(pid);
  const wins = c.filter((s) => s.champion);
  const placed = c.filter((s) => s.place != null);
  const best = placed.length ? Math.min(...placed.map((s) => s.place)) : null;
  const avg = placed.length ? (placed.reduce((a, s) => a + s.place, 0) / placed.length) : null;

  $("#p-name").textContent = nameOf(pid);
  $("#p-sub").textContent =
    `${c.length} trips (${c[0].year}–${c[c.length - 1].year})` +
    (wins.length ? ` · Champion: ${wins.map((s) => s.year).join(", ")}` : "");
  $("#p-wins").textContent = wins.length;
  $("#p-best").textContent = best ? ordinal(best) : "—";
  $("#p-avg").textContent = avg ? avg.toFixed(1) : "—";
  $("#p-trips").textContent = c.length;

  // seasons table
  $("#p-seasons").innerHTML = `<thead><tr><th>Year</th><th>Place</th><th>Gross</th><th>Net</th></tr></thead>
    <tbody>${[...c].reverse().map((s) => `
      <tr class="${s.champion ? "first" : ""}">
        <td><a href="years.html#${s.year}">${s.year}</a></td>
        <td>${s.place != null ? `${s.place} of ${s.of}` : "—"}${s.champion ? " \u{1F3C6}" : ""}</td>
        <td>${s.gross ?? "—"}</td><td>${s.net ?? "—"}</td>
      </tr>`).join("")}</tbody>`;

  // chart
  if (chart) chart.destroy();
  const pts = placed.map((s) => ({ x: s.year, y: s.place }));
  chart = new Chart($("#p-chart"), {
    type: "line",
    data: { datasets: [{
      label: "Finish",
      data: pts,
      borderColor: "#1e5631",
      backgroundColor: "#c9a227",
      pointRadius: 4,
      tension: 0.25,
    }] },
    options: {
      scales: {
        x: { type: "linear", ticks: { stepSize: 1, callback: (v) => String(v) } },
        y: { reverse: true, min: 1, ticks: { stepSize: 1 }, title: { display: true, text: "Place" } },
      },
      plugins: { legend: { display: false } },
    },
  });
}

/* -------------------------------------------------------------- helpers */

function shortLoc(loc) {
  return loc.split(",")[0].split("/")[0].trim();
}
function ordinal(n) {
  const s = ["th", "st", "nd", "rd"], v = n % 100;
  return n + (s[(v - 20) % 10] || s[v] || s[0]);
}
function esc(s) {
  return String(s).replace(/[&<>"']/g, (c) =>
    ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}

/* ---------------------------------------------------------------- init */

loadData().then(() => {
  const page = document.body.dataset.page;
  if (page === "home") renderHome();
  else if (page === "years") renderYears();
  else if (page === "players") renderPlayers();
}).catch((err) => {
  document.querySelector("main").insertAdjacentHTML("afterbegin",
    `<div class="card">Failed to load data/hhga.json — ${esc(err.message)}.
     If you opened this file directly, serve the folder instead
     (e.g. <code>python3 -m http.server</code>).</div>`);
});
