// 저장소 data/ 폴더 내 기본 Excel 파일 목록
const DEFAULT_DATA_FILES = [
  "(표)(2023연장)투입산출표_생산자가격_통합대분류.xlsx",
  "(표)(2023연장)투입산출표_생산자가격_통합중분류.xlsx",
  "(표)(2023연장)투입산출표_생산자가격_통합소분류.xlsx",
  "(표)(2023연장)투입산출표_생산자가격_기본부문.xlsx",
  "(표)상품분류.xlsx",
  "2023년 산업연관표(연장표).xlsx"
];

const worker = new Worker("./worker.js");
let msgCounter = 0;
const pending = new Map();

const state = {
  files: {},
  levels: [],
  currentLevel: "",
  modelType: "domestic_est",
  structureData: null,
  sectorsByLevel: {}, // level -> [{code, name}]
  shocks: [],         // [{code, name, amount}]
  simResult: null
};

function callWorker(action, payload = {}) {
  return new Promise((resolve, reject) => {
    const id = ++msgCounter;
    pending.set(id, { resolve, reject });
    worker.postMessage({ id, action, payload });
  });
}

worker.onmessage = (e) => {
  const msg = e.data;
  if (msg.type === "status") {
    setProgress(msg.message, "busy");
    return;
  }
  const entry = pending.get(msg.id);
  if (!entry) return;
  pending.delete(msg.id);
  if (msg.ok) entry.resolve(msg.data);
  else entry.reject(new Error(msg.error));
};

function setProgress(text, mode = "idle") {
  document.getElementById("loadProgressText").textContent = text;
  const badge = document.getElementById("engineStatus");
  if (mode === "busy") {
    badge.textContent = "처리 중...";
    badge.className = "header-badge busy";
  } else if (mode === "ready") {
    badge.textContent = "분석 준비 완료";
    badge.className = "header-badge ready";
  } else {
    badge.textContent = "엔진 대기 중";
    badge.className = "header-badge";
  }
}

function fmtNum(v, digits = 0) {
  if (v === null || v === undefined || Number.isNaN(Number(v))) return "-";
  return Number(v).toLocaleString("ko-KR", {
    minimumFractionDigits: digits,
    maximumFractionDigits: digits
  });
}

function fmtJo(vMillion) {
  return (Number(vMillion) / 1_000_000).toFixed(2) + " 조원";
}

function escapeHtml(str) {
  return String(str ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

// 탭 전환
document.querySelectorAll(".tab-btn").forEach((btn) => {
  btn.addEventListener("click", () => {
    document.querySelectorAll(".tab-btn").forEach((b) => b.classList.remove("active"));
    document.querySelectorAll(".tab-panel").forEach((p) => p.classList.remove("active"));
    btn.classList.add("active");
    document.getElementById(btn.dataset.tab).classList.add("active");
  });
});

// 1. data/ 폴더에서 기본 Excel 불러오기
document.getElementById("btnLoadRepoData").addEventListener("click", async () => {
  const btn = document.getElementById("btnLoadRepoData");
  btn.disabled = true;
  try {
    setProgress("Pyodide 엔진을 초기화하는 중입니다...", "busy");
    await callWorker("init");

    let loadedCount = 0;
    const failed = [];
    for (const fname of DEFAULT_DATA_FILES) {
      setProgress(`Excel 불러오는 중 (${loadedCount + 1}/${DEFAULT_DATA_FILES.length}): ${fname}`, "busy");
      const url = "./data/" + encodeURIComponent(fname);
      try {
        const resp = await fetch(url);
        if (!resp.ok) {
          failed.push(`${fname} (HTTP ${resp.status})`);
          continue;
        }
        const buf = await resp.arrayBuffer();
        await callWorker("load_file", { filename: fname, bytes: buf });
        loadedCount++;
      } catch (err) {
        failed.push(`${fname} (${err.message})`);
      }
    }

    if (loadedCount === 0) {
      setProgress(
        `불러오기 실패 원인: ${failed.join(" | ")}`,
        "idle"
      );
    } else {
      const warnMsg = failed.length ? ` (일부 미발견: ${failed.join(", ")})` : "";
      setProgress(`총 ${loadedCount}개 Excel 파일 로딩 완료!${warnMsg}`, "ready");
      await refreshAllAfterLoad();
    }
  } catch (err) {
    setProgress("오류 발생: " + err.message, "idle");
  } finally {
    btn.disabled = false;
  }
});

// 2. 로컬 Excel 파일 직접 선택 업로드
document.getElementById("fileInput").addEventListener("change", async (e) => {
  const files = Array.from(e.target.files || []);
  if (!files.length) return;
  try {
    setProgress("Pyodide 엔진 초기화 및 로컬 파일 읽는 중...", "busy");
    await callWorker("init");
    let count = 0;
    for (const f of files) {
      setProgress(`로컬 Excel 파싱 중 (${++count}/${files.length}): ${f.name}`, "busy");
      const buf = await f.arrayBuffer();
      await callWorker("load_file", { filename: f.name, bytes: buf });
    }
    setProgress(`로컬 Excel 파일 ${count}개 로딩 완료!`, "ready");
    await refreshAllAfterLoad();
  } catch (err) {
    setProgress("파일 로딩 오류: " + err.message, "idle");
  } finally {
    e.target.value = "";
  }
});

async function refreshAllAfterLoad() {
  const summary = await callWorker("get_state");
  state.files = summary.files;
  state.levels = summary.levels;

  // 로딩된 파일 칩 표시
  const chips = document.getElementById("loadedFilesChips");
  chips.innerHTML = Object.entries(state.files)
    .map(([fname, info]) => `<span class="chip">${escapeHtml(fname)} (${info.sheets.length}개 시트)</span>`)
    .join("");

  // 시트 탐색기 파일 목록
  const fileSel = document.getElementById("sheetFileSelect");
  fileSel.innerHTML = Object.keys(state.files)
    .map((fn) => `<option value="${escapeHtml(fn)}">${escapeHtml(fn)}</option>`)
    .join("");
  updateSheetDropdown();
  await loadSheetPreview();

  // 상품분류 검색
  await runClassSearch();

  // 분석 분류 드롭다운
  const lvlSel = document.getElementById("globalLevelSelect");
  if (state.levels.length > 0) {
    lvlSel.disabled = false;
    lvlSel.innerHTML = state.levels
      .map((l) => `<option value="${escapeHtml(l.level)}">${escapeHtml(l.label)}</option>`)
      .join("");
    if (!state.currentLevel || !state.levels.some((l) => l.level === state.currentLevel)) {
      state.currentLevel = state.levels[0].level;
    }
    lvlSel.value = state.currentLevel;

    await loadLevelData();
    await loadVerification();
    await populateMajorDrilldown();
  }
}

// 시트 탐색기
function updateSheetDropdown() {
  const fn = document.getElementById("sheetFileSelect").value;
  const sSel = document.getElementById("sheetNameSelect");
  const sheets = (state.files[fn] && state.files[fn].sheets) || [];
  sSel.innerHTML = sheets.map((s) => `<option value="${escapeHtml(s)}">${escapeHtml(s)}</option>`).join("");
}

document.getElementById("sheetFileSelect").addEventListener("change", () => {
  updateSheetDropdown();
  loadSheetPreview();
});
document.getElementById("sheetNameSelect").addEventListener("change", loadSheetPreview);
document.getElementById("btnSearchSheet").addEventListener("click", loadSheetPreview);

async function loadSheetPreview() {
  const fn = document.getElementById("sheetFileSelect").value;
  const sn = document.getElementById("sheetNameSelect").value;
  if (!fn || !sn) return;
  const q = document.getElementById("sheetSearchInput").value.trim();
  const res = await callWorker("get_sheet_preview", { filename: fn, sheet_name: sn, query: q, limit: 80 });
  document.getElementById("sheetMetaBox").textContent = `${fn} > [${sn}] — ${res.meta}`;
  const tbl = document.getElementById("sheetPreviewTable");
  if (!res.rows.length) {
    tbl.innerHTML = `<tbody><tr><td class="empty-cell">검색 조건에 맞는 행이 없습니다.</td></tr></tbody>`;
    return;
  }
  tbl.innerHTML = "<tbody>" + res.rows.map((r) =>
    `<tr><td class="muted">${r.row_no}</td>` +
    r.cells.map((c) => `<td>${escapeHtml(c)}</td>`).join("") +
    "</tr>"
  ).join("") + "</tbody>";
}

// 상품분류 계층 검색
document.getElementById("btnSearchClass").addEventListener("click", runClassSearch);
document.getElementById("classSearchInput").addEventListener("keydown", (e) => {
  if (e.key === "Enter") runClassSearch();
});

async function runClassSearch() {
  const q = document.getElementById("classSearchInput").value.trim();
  const res = await callWorker("search_classification", { query: q, limit: 150 });
  const box = document.getElementById("classSummaryCount");
  const tbody = document.querySelector("#classTable tbody");
  if (!res.total) {
    box.textContent = "(표)상품분류.xlsx가 아직 로딩되지 않았습니다.";
    return;
  }
  box.textContent = `총 ${res.total}개 기본부문 중 ${res.matched}건 표시 (최대 150건)`;
  tbody.innerHTML = res.items.map((it) => `
    <tr>
      <td><code>${escapeHtml(it.div_code)}</code> ${escapeHtml(it.div_name)}</td>
      <td><code>${escapeHtml(it.major_code)}</code> ${escapeHtml(it.major_name)}</td>
      <td><code>${escapeHtml(it.minor_code)}</code> ${escapeHtml(it.minor_name)}</td>
      <td><code>${escapeHtml(it.basic_code)}</code> ${escapeHtml(it.basic_name)}</td>
      <td>
        <button class="btn btn-chip" onclick="jumpToSector('major33','${escapeHtml(it.div_code)}')">대</button>
        <button class="btn btn-chip" onclick="jumpToSector('mid83','${escapeHtml(it.major_code)}')">중</button>
        <button class="btn btn-chip" onclick="jumpToSector('minor165','${escapeHtml(it.minor_code)}')">소</button>
        <button class="btn btn-chip" onclick="jumpToSector('basic380','${escapeHtml(it.basic_code)}')">기본</button>
      </td>
    </tr>
  `).join("");
}

window.jumpToSector = async function (targetLevel, sectorCode) {
  if (!state.levels.some((l) => l.level === targetLevel)) {
    alert(`${targetLevel} 투입산출표가 먼저 로딩되어야 합니다.`);
    return;
  }
  state.currentLevel = targetLevel;
  document.getElementById("globalLevelSelect").value = targetLevel;
  await loadLevelData();

  document.querySelector('[data-tab="tab-linkage"]').click();
  document.getElementById("sectorSelect").value = sectorCode;
  await loadSectorDetail();
};

// 분류·모형 변경 이벤트
document.getElementById("globalLevelSelect").addEventListener("change", async (e) => {
  state.currentLevel = e.target.value;
  state.shocks = [];
  renderShockList();
  await loadLevelData();
});

document.getElementById("globalModelType").addEventListener("change", async (e) => {
  state.modelType = e.target.value;
  await loadSectorDetail();
  if (state.shocks.length > 0) await runSimulation();
});

async function loadLevelData() {
  if (!state.currentLevel) return;
  const data = await callWorker("get_structure", { level: state.currentLevel });
  state.structureData = data;
  state.sectorsByLevel[state.currentLevel] = data.rows.map((r) => ({ code: r.code, name: r.name }));

  renderMacroKpis(data);
  renderStructureBarChart();
  renderBreakdown(data);
  renderStructureTable(data.rows);
  populateSectorSelectors();
  await loadSectorDetail();
}

function renderMacroKpis(data) {
  const k = data.kpi;
  const cards = document.querySelectorAll("#macroKpiGrid .kpi-card");
  const items = [
    [fmtNum(k.total_supply), `${fmtJo(k.total_supply)} (백만원)`],
    [fmtNum(k.total_output), `${fmtJo(k.total_output)} (백만원)`],
    [fmtNum(k.total_va), `${fmtJo(k.total_va)} (백만원)`],
    [`${fmtNum(k.total_import)} / ${fmtNum(k.total_export)}`, `수입 ${fmtJo(k.total_import)} / 수출 ${fmtJo(k.total_export)}`],
    [`${k.int_input_rate}% / ${k.va_rate}%`, "중간투입률 + 부가가치율 = 100%"],
    [`${k.trade_ratio}%`, "(수출+수입) / 총공급 비율"]
  ];
  cards.forEach((c, idx) => {
    c.querySelector(".kpi-val").textContent = items[idx][0];
    c.querySelector(".kpi-unit").textContent = items[idx][1];
  });
}

document.getElementById("structureChartMetric").addEventListener("change", renderStructureBarChart);

function renderStructureBarChart() {
  if (!state.structureData) return;
  const metric = document.getElementById("structureChartMetric").value;
  const rows = [...state.structureData.rows].sort((a, b) => b[metric] - a[metric]).slice(0, 10);
  const maxVal = Math.max(...rows.map((r) => Math.abs(r[metric])), 1);
  const isPct = metric.endsWith("_rate");

  document.getElementById("structureBarChart").innerHTML = rows.map((r) => {
    const w = Math.min(100, (Math.abs(r[metric]) / maxVal) * 100);
    return `
      <div class="bar-row">
        <div class="bar-name" title="${escapeHtml(r.code + " " + r.name)}">${escapeHtml(r.code)} ${escapeHtml(r.name)}</div>
        <div class="bar-track"><div class="bar-fill" style="width:${w}%"></div></div>
        <div class="bar-val">${fmtNum(r[metric], isPct ? 2 : 0)}${isPct ? "%" : ""}</div>
      </div>
    `;
  }).join("");
}

function renderBreakdown(data) {
  const mkTable = (title, items) => `
    <h4 style="margin:8px 0 6px; font-size:0.86rem;">${title}</h4>
    <table class="data-table" style="margin-bottom:12px;">
      <thead><tr><th>항목</th><th class="num">금액(백만원)</th><th class="num">구성비(%)</th></tr></thead>
      <tbody>
        ${items.map((x) => `<tr><td>${escapeHtml(x.name)}</td><td class="num">${fmtNum(x.value)}</td><td class="num">${x.share}%</td></tr>`).join("")}
      </tbody>
    </table>
  `;
  document.getElementById("macroBreakdownBox").innerHTML =
    mkTable("최종수요 항목별 구성", data.fd_breakdown) +
    mkTable("부가가치 항목별 구성", data.va_breakdown);
}

document.getElementById("structureFilterInput").addEventListener("input", () => {
  if (state.structureData) renderStructureTable(state.structureData.rows);
});

function renderStructureTable(rows) {
  const q = document.getElementById("structureFilterInput").value.trim().toLowerCase();
  const filtered = q ? rows.filter((r) => (r.code + " " + r.name).toLowerCase().includes(q)) : rows;
  document.querySelector("#structureTable tbody").innerHTML = filtered.map((r) => `
    <tr>
      <td><code>${escapeHtml(r.code)}</code></td>
      <td>${escapeHtml(r.name)}</td>
      <td class="num">${fmtNum(r.output)}</td>
      <td class="num">${r.output_share}%</td>
      <td class="num">${fmtNum(r.va)}</td>
      <td class="num">${r.va_rate}%</td>
      <td class="num">${r.int_input_rate}%</td>
      <td class="num">${fmtNum(r.export)}</td>
      <td class="num">${r.export_rate}%</td>
      <td class="num">${fmtNum(r.import)}</td>
      <td class="num">${r.import_rate}%</td>
    </tr>
  `).join("");
}

// 산업 선택 드롭다운 채우기
function populateSectorSelectors() {
  const secs = state.sectorsByLevel[state.currentLevel] || [];
  const filterQ = document.getElementById("sectorFilterInput").value.trim().toLowerCase();
  const filtered = filterQ ? secs.filter((s) => (s.code + " " + s.name).toLowerCase().includes(filterQ)) : secs;

  const sSel = document.getElementById("sectorSelect");
  const prev = sSel.value;
  sSel.innerHTML = filtered.map((s) => `<option value="${escapeHtml(s.code)}">[${escapeHtml(s.code)}] ${escapeHtml(s.name)}</option>`).join("");
  if (filtered.some((s) => s.code === prev)) sSel.value = prev;

  const simSel = document.getElementById("simSectorSelect");
  simSel.innerHTML = secs.map((s) => `<option value="${escapeHtml(s.code)}">[${escapeHtml(s.code)}] ${escapeHtml(s.name)}</option>`).join("");
}

document.getElementById("sectorFilterInput").addEventListener("input", populateSectorSelectors);
document.getElementById("btnAnalyzeSector").addEventListener("click", loadSectorDetail);
document.getElementById("sectorSelect").addEventListener("change", loadSectorDetail);

async function loadSectorDetail() {
  if (!state.currentLevel) return;
  const code = document.getElementById("sectorSelect").value;
  if (!code) return;

  const res = await callWorker("get_sector_detail", {
    level: state.currentLevel,
    code,
    model_type: state.modelType
  });
  if (res.error) return;

  const warnBox = document.getElementById("sectorWarningBox");
  if (res.zero_output_warning) {
    warnBox.classList.remove("hidden");
    warnBox.textContent = `[참고] ${res.sector.name}(${res.sector.code}) 부문은 2023년 국내 총산출이 0(전량 수입 의존 등)이므로, 투입계수 열이 0으로 처리되어 생산유발계수가 1.0으로 계산됩니다.`;
  } else {
    warnBox.classList.add("hidden");
  }

  const k = res.kpi;
  const cards = document.querySelectorAll("#sectorKpiGrid .kpi-card");
  const vals = [
    [`${fmtNum(k.output)} / ${fmtNum(k.va)}`, "백만원 (총산출 / 부가가치)"],
    [k.prod_mult.toFixed(3), "최종수요 1단위당 생산유발"],
    [k.va_mult.toFixed(3), "최종수요 1단위당 부가가치유발"],
    [k.imp_mult.toFixed(3), "국내생산 1단위당 간접 수입유발"],
    [k.backward.toFixed(3), k.backward >= 1 ? "평균 이상 후방연쇄(≥1.0)" : "평균 미만(<1.0)"],
    [k.forward.toFixed(3), k.forward >= 1 ? "평균 이상 전방연쇄(≥1.0)" : "평균 미만(<1.0)"]
  ];
  cards.forEach((c, idx) => {
    c.querySelector(".kpi-val").textContent = vals[idx][0];
    c.querySelector(".kpi-unit").textContent = vals[idx][1];
  });

  document.querySelector("#upstreamTable tbody").innerHTML = res.upstream.map((r) => `
    <tr><td><code>${escapeHtml(r.code)}</code></td><td>${escapeHtml(r.name)}</td><td class="num">${fmtNum(r.value)}</td><td class="num">${r.coeff.toFixed(4)}</td><td class="num">${r.share}%</td></tr>
  `).join("");

  document.querySelector("#downstreamTable tbody").innerHTML = res.downstream.map((r) => `
    <tr><td><code>${escapeHtml(r.code)}</code></td><td>${escapeHtml(r.name)}</td><td class="num">${fmtNum(r.value)}</td><td class="num">${r.coeff.toFixed(4)}</td><td class="num">${r.share}%</td></tr>
  `).join("");

  document.querySelector("#inducedTopTable tbody").innerHTML = res.induced_top.map((r) => `
    <tr><td><code>${escapeHtml(r.code)}</code></td><td>${escapeHtml(r.name)}</td><td class="num">${r.coeff.toFixed(4)}</td><td class="num">${r.share}%</td></tr>
  `).join("");

  renderQuadrantChart(res.scatter);
}

function renderQuadrantChart(points) {
  const box = document.getElementById("quadrantChartBox");
  const W = 460, H = 310, PAD = 40;
  const maxX = Math.max(2.0, ...points.map((p) => p.forward)) * 1.05;
  const maxY = Math.max(2.0, ...points.map((p) => p.backward)) * 1.05;
  const sx = (x) => PAD + (x / maxX) * (W - 2 * PAD);
  const sy = (y) => H - PAD - (y / maxY) * (H - 2 * PAD);

  const lineX = sx(1.0);
  const lineY = sy(1.0);

  const dots = points.map((p) => {
    const cx = sx(p.forward);
    const cy = sy(p.backward);
    const fill = p.selected ? "#d32f2f" : "#1f6feb";
    const r = p.selected ? 6 : 3.2;
    const op = p.selected ? 1 : 0.55;
    return `<circle cx="${cx}" cy="${cy}" r="${r}" fill="${fill}" opacity="${op}">
      <title>[${escapeHtml(p.code)}] ${escapeHtml(p.name)} | 감응도:${p.forward}, 영향력:${p.backward}</title>
    </circle>`;
  }).join("");

  const sel = points.find((p) => p.selected);
  const labelSvg = sel
    ? `<text x="${Math.min(W - 110, sx(sel.forward) + 8)}" y="${Math.max(16, sy(sel.backward) - 8)}" font-size="11" font-weight="700" fill="#b71c1c">${escapeHtml(sel.code)} ${escapeHtml(sel.name.slice(0, 10))}</text>`
    : "";

  box.innerHTML = `
    <svg viewBox="0 0 ${W} ${H}" width="100%" style="max-width:480px; background:#fbfcfe; border:1px solid #e3e9f2; border-radius:6px;">
      <line x1="${PAD}" y1="${H - PAD}" x2="${W - PAD}" y2="${H - PAD}" stroke="#8898aa" />
      <line x1="${PAD}" y1="${PAD}" x2="${PAD}" y2="${H - PAD}" stroke="#8898aa" />
      <line x1="${lineX}" y1="${PAD}" x2="${lineX}" y2="${H - PAD}" stroke="#9aa8b8" stroke-dasharray="4 3" />
      <line x1="${PAD}" y1="${lineY}" x2="${W - PAD}" y2="${lineY}" stroke="#9aa8b8" stroke-dasharray="4 3" />
      <text x="${W - PAD - 75}" y="${PAD + 14}" font-size="10" fill="#5b6778">핵심기간산업형</text>
      <text x="${PAD + 6}" y="${PAD + 14}" font-size="10" fill="#5b6778">최종수요제조형</text>
      <text x="${W - PAD - 68}" y="${H - PAD - 8}" font-size="10" fill="#5b6778">중간재기초형</text>
      <text x="${PAD + 6}" y="${H - PAD - 8}" font-size="10" fill="#5b6778">독립산업형</text>
      <text x="${W / 2 - 42}" y="${H - 10}" font-size="11" fill="#1c2430">감응도계수 (1.0 기준)</text>
      <text x="12" y="${H / 2}" font-size="11" fill="#1c2430" transform="rotate(-90 12 ${H / 2})">영향력계수</text>
      ${dots}
      ${labelSvg}
    </svg>
  `;
}

// TAB 4: 시뮬레이션
document.getElementById("btnAddShock").addEventListener("click", () => {
  const code = document.getElementById("simSectorSelect").value;
  const amount = Number(document.getElementById("simShockInput").value);
  if (!code || Number.isNaN(amount)) return;
  const sec = (state.sectorsByLevel[state.currentLevel] || []).find((s) => s.code === code);
  if (!sec) return;
  state.shocks.push({ code: sec.code, name: sec.name, amount });
  renderShockList();
});

document.getElementById("btnClearShock").addEventListener("click", () => {
  state.shocks = [];
  renderShockList();
});

function renderShockList() {
  const tbody = document.querySelector("#shockListTable tbody");
  if (!state.shocks.length) {
    tbody.innerHTML = `<tr><td colspan="5" class="empty-cell">충격 항목을 추가하거나 위 프리셋을 선택하세요.</td></tr>`;
    return;
  }
  tbody.innerHTML = state.shocks.map((sh, i) => `
    <tr>
      <td><code>${escapeHtml(sh.code)}</code></td>
      <td>${escapeHtml(sh.name)}</td>
      <td class="num">${fmtNum(sh.amount)}</td>
      <td class="num">${fmtJo(sh.amount)}</td>
      <td><button class="btn btn-chip" onclick="removeShock(${i})">삭제</button></td>
    </tr>
  `).join("");
}

window.removeShock = (idx) => {
  state.shocks.splice(idx, 1);
  renderShockList();
};

// 프리셋 시나리오
document.querySelectorAll("[data-preset]").forEach((btn) => {
  btn.addEventListener("click", () => {
    const secs = state.sectorsByLevel[state.currentLevel] || [];
    if (!secs.length) { alert("먼저 Excel 자료를 불러오세요."); return; }

    const findSec = (keywords) => secs.find((s) => keywords.some((kw) => s.code === kw || s.name.includes(kw)));
    const p = btn.dataset.preset;
    state.shocks = [];

    if (p === "semi") {
      const s = findSec(["3102", "310", "31", "C09"]);
      if (s) state.shocks.push({ code: s.code, name: s.name, amount: -10_000_000 });
    } else if (p === "ev") {
      const s1 = findSec(["4012", "401", "40", "C12"]);
      const s2 = findSec(["3730", "373", "37", "C10"]);
      if (s1) state.shocks.push({ code: s1.code, name: s1.name, amount: 3_000_000 });
      if (s2) state.shocks.push({ code: s2.code, name: s2.name, amount: 2_000_000 });
    } else if (p === "soc") {
      const s = findSec(["5111", "51", "F"]);
      if (s) state.shocks.push({ code: s.code, name: s.name, amount: 3_000_000 });
    } else if (p === "service") {
      const s1 = findSec(["5811", "581", "58", "I"]);
      const s2 = findSec(["5500", "55", "H"]);
      if (s1) state.shocks.push({ code: s1.code, name: s1.name, amount: 1_200_000 });
      if (s2) state.shocks.push({ code: s2.code, name: s2.name, amount: 800_000 });
    }
    renderShockList();
    runSimulation();
  });
});

document.getElementById("btnRunSimulation").addEventListener("click", runSimulation);

async function runSimulation() {
  if (!state.currentLevel || !state.shocks.length) {
    alert("충격 항목을 1개 이상 추가하세요.");
    return;
  }
  const res = await callWorker("run_simulation", {
    level: state.currentLevel,
    model_type: state.modelType,
    shocks: state.shocks
  });
  state.simResult = res;

  const k = res.kpi;
  const cards = document.querySelectorAll("#simKpiGrid .kpi-card");
  const vals = [
    [fmtNum(k.total_dy), `${fmtJo(k.total_dy)} (백만원)`],
    [fmtNum(k.total_dx), `${fmtJo(k.total_dx)} | 승수 ${k.prod_multiplier.toFixed(3)}`],
    [fmtNum(k.total_dv), `${fmtJo(k.total_dv)} | 부가가치유발율 ${k.va_ratio.toFixed(3)}`],
    [fmtNum(k.total_dw), `${fmtJo(k.total_dw)} (임금 유발)`],
    [fmtNum(k.total_dm), `${fmtJo(k.total_dm)} (중간재 수입 유발)`]
  ];
  cards.forEach((c, i) => {
    c.querySelector(".kpi-val").textContent = vals[i][0];
    c.querySelector(".kpi-unit").textContent = vals[i][1];
  });

  const top10 = res.results.slice(0, 10);
  const maxAbs = Math.max(...top10.map((r) => Math.abs(r.total_prod)), 1);
  document.getElementById("simBarChart").innerHTML = top10.map((r) => {
    const w = Math.min(100, (Math.abs(r.total_prod) / maxAbs) * 100);
    return `
      <div class="bar-row">
        <div class="bar-name">${escapeHtml(r.code)} ${escapeHtml(r.name)}</div>
        <div class="bar-track"><div class="bar-fill ${r.total_prod < 0 ? "neg" : ""}" style="width:${w}%"></div></div>
        <div class="bar-val">${fmtNum(r.total_prod)}</div>
      </div>
    `;
  }).join("");

  document.querySelector("#simResultTable tbody").innerHTML = res.results.slice(0, 120).map((r) => `
    <tr>
      <td><code>${escapeHtml(r.code)}</code></td>
      <td>${escapeHtml(r.name)}</td>
      <td class="num">${fmtNum(r.direct)}</td>
      <td class="num">${fmtNum(r.indirect)}</td>
      <td class="num"><strong>${fmtNum(r.total_prod)}</strong></td>
      <td class="num">${r.share}%</td>
      <td class="num">${fmtNum(r.va)}</td>
      <td class="num">${fmtNum(r.wage)}</td>
      <td class="num">${fmtNum(r.imp)}</td>
    </tr>
  `).join("");
}

// TAB 5: 정합성 검증 및 드릴다운
async function loadVerification() {
  const res = await callWorker("verify_all");
  document.querySelector("#verifyTable tbody").innerHTML = res.rows.map((r) => `
    <tr>
      <td><strong>${escapeHtml(r.label)}</strong></td>
      <td class="num">${r.n}</td>
      <td class="num">${fmtNum(r.sum_output)}</td>
      <td class="num">${fmtNum(r.sum_input)}</td>
      <td class="num">${fmtNum(r.sum_supply_demand)}</td>
      <td class="num">${fmtNum(r.sum_va)}</td>
      <td class="num">${fmtNum(r.max_error, 2)}</td>
      <td>${r.passed ? '<span class="badge-ok">일치(PASS)</span>' : '<span class="badge-warn">차이 확인</span>'}</td>
    </tr>
  `).join("");
}

async function populateMajorDrilldown() {
  if (!state.levels.some((l) => l.level === "major33")) return;
  const data = await callWorker("get_structure", { level: "major33" });
  const sel = document.getElementById("drillMajorSelect");
  sel.innerHTML = data.rows.map((r) => `<option value="${escapeHtml(r.code)}">[${escapeHtml(r.code)}] ${escapeHtml(r.name)}</option>`).join("");
}

document.getElementById("btnRunDrilldown").addEventListener("click", async () => {
  const div_code = document.getElementById("drillMajorSelect").value;
  const target_level = document.getElementById("drillTargetLevel").value;
  if (!div_code) return;
  const res = await callWorker("drilldown", { div_code, target_level, model_type: state.modelType });
  const tbody = document.querySelector("#drillTable tbody");
  if (res.error) {
    tbody.innerHTML = `<tr><td colspan="8" class="empty-cell">${escapeHtml(res.error)}</td></tr>`;
    return;
  }
  tbody.innerHTML = res.rows.map((r) => `
    <tr>
      <td><code>${escapeHtml(r.code)}</code></td>
      <td>${escapeHtml(r.name)}</td>
      <td class="num">${fmtNum(r.output)}</td>
      <td class="num">${r.share}%</td>
      <td class="num">${r.va_rate}%</td>
      <td class="num">${r.export_rate}%</td>
      <td class="num">${r.prod_mult.toFixed(3)}</td>
      <td class="num">${r.va_mult.toFixed(3)}</td>
    </tr>
  `).join("");
});

// CSV 내보내기
function downloadCsv(filename, headers, rows) {
  const bom = "\uFEFF";
  const csv = [headers.join(","), ...rows.map((r) => r.map((v) => `"${String(v).replace(/"/g, '""')}"`).join(","))].join("\n");
  const blob = new Blob([bom + csv], { type: "text/csv;charset=utf-8;" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  a.click();
}

document.getElementById("btnExportStructureCsv").addEventListener("click", () => {
  if (!state.structureData) return;
  const headers = ["코드", "산업명", "총산출(백만원)", "구성비(%)", "부가가치(백만원)", "부가가치율(%)", "중간투입률(%)", "수출액(백만원)", "수출률(%)", "수입계(백만원)", "수입률(%)"];
  const rows = state.structureData.rows.map((r) => [
    r.code, r.name, r.output, r.output_share, r.va, r.va_rate, r.int_input_rate, r.export, r.export_rate, r.import, r.import_rate
  ]);
  downloadCsv(`IO_2023_structure_${state.currentLevel}.csv`, headers, rows);
});

document.getElementById("btnExportSimCsv").addEventListener("click", () => {
  if (!state.simResult) return;
  const headers = ["코드", "산업명", "직접충격(백만원)", "간접유발(백만원)", "총생산유발(백만원)", "비중(%)", "부가가치유발(백만원)", "피용자보수유발(백만원)", "수입유발(백만원)"];
  const rows = state.simResult.results.map((r) => [
    r.code, r.name, r.direct, r.indirect, r.total_prod, r.share, r.va, r.wage, r.imp
  ]);
  downloadCsv(`IO_2023_simulation_${state.currentLevel}.csv`, headers, rows);
});
