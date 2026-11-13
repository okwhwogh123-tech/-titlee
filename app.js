'use strict';
/* =========================================================================
   2023 산업연관표 분석 웹앱 — 화면 로직
   - 계산은 worker.js(Pyodide) → model.py 에서 수행
   - 이 파일은 버튼 동작, 표·그래프 그리기, CSV 저장만 담당
   ========================================================================= */
const $ = (id) => document.getElementById(id);
const LEVELS = {
  basic: { label: '기본부문', file: './data/io-2023-basic.xlsx', n: 380, note: '한국은행 공식 파일' },
  small: { label: '소분류', file: './data/io-2023-small.xlsx', n: 165, note: '한은 부문분류표에 따라 기본부문 합산' },
  medium: { label: '중분류', file: './data/io-2023-medium.xlsx', n: 83, note: '한은 부문분류표에 따라 기본부문 합산' },
  large: { label: '대분류', file: './data/io-2023-large.xlsx', n: 33, note: '한국은행 공식 파일' },
  s13: { label: '13부문' }, s7: { label: '7부문' }, s1: { label: '1부문(전산업)' },
};
const ORDER = ['basic', 'small', 'medium', 'large', 's13', 's7', 's1'];
const OFFICIAL = ['basic', 'small', 'medium', 'large'];
const CLASS_FILE = './data/classification.xlsx';

const state = {
  ready: false, classLoaded: false, info: {}, level: null, tab: 'data',
  cache: {}, basket: [], priceTargets: [], shockResult: null, priceResult: null,
  tree: [], loading: false,
};

/* ---------- 공통 도우미 ---------- */
const nf = (d) => new Intl.NumberFormat('ko-KR', { maximumFractionDigits: d, minimumFractionDigits: d });
const NF = [0, 1, 2, 3, 4].map(nf);
function fmt(v, d = 0) {
  if (v === null || v === undefined || Number.isNaN(v)) return '–';
  return NF[d].format(v);
}
const won = (v) => fmt(v, 0);
function el(tag, attrs = {}, ...kids) {
  const e = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs)) {
    if (v === undefined || v === null || v === false) continue;
    if (k === 'class') e.className = v;
    else if (k === 'text') e.textContent = v;
    else if (k.startsWith('on')) e[k] = v;
    else e.setAttribute(k, v === true ? '' : v);
  }
  for (const c of kids.flat()) if (c !== null && c !== undefined) e.append(c instanceof Node ? c : String(c));
  return e;
}
function td(v, num = false, cls = '') { return el('td', { class: [num ? 'num' : '', cls].join(' ').trim() || null }, v); }
function setStatus(t) { $('status').textContent = t; }
function q(s) { return (s || '').replace(/\s+/g, '').toLowerCase(); }
function match(item, query) { const k = q(query); return !k || q(item.code + item.name).includes(k); }
function lab(level) { return LEVELS[level]?.label || level; }

/* ---------- Worker 통신 ---------- */
let worker; let seq = 0; const pending = new Map();
function startWorker() {
  worker = new Worker('./worker.js');
  worker.onmessage = ({ data: d }) => {
    if (d.type === 'status') { setStatus(d.text); return; }
    const p = pending.get(d.id); if (!p) return; pending.delete(d.id);
    d.ok ? p.resolve(d.result) : p.reject(Error(d.error));
  };
  worker.onerror = (e) => {
    for (const p of pending.values()) p.reject(Error(e.message || 'Worker 오류'));
    pending.clear(); setStatus('계산 환경 오류: ' + (e.message || '') + ' — 새로고침하세요.');
  };
}
function send(type, payload, transfer) {
  return new Promise((resolve, reject) => {
    const id = ++seq; pending.set(id, { resolve, reject });
    worker.postMessage({ id, type, payload }, transfer || []);
  });
}
const call = (fn, arg) => send('call', { fn, arg: typeof arg === 'string' ? arg : JSON.stringify(arg) });
async function cached(key, fn) {
  const k = state.level + '|' + key;
  if (!state.cache[k]) state.cache[k] = fn();
  try { return await state.cache[k]; } catch (e) { delete state.cache[k]; throw e; }
}

/* ---------- CSV ---------- */
function csvCell(v) {
  let s = v === null || v === undefined ? '' : String(v);
  if (typeof v !== 'number' && /^[=+\-@]/.test(s)) s = "'" + s;
  return '"' + s.replaceAll('"', '""') + '"';
}
function downloadCSV(name, header, rows) {
  const text = '﻿' + [header, ...rows].map((r) => r.map(csvCell).join(',')).join('\r\n');
  const url = URL.createObjectURL(new Blob([text], { type: 'text/csv;charset=utf-8' }));
  const a = el('a', { href: url, download: name }); document.body.append(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

/* ---------- 정렬 가능한 표 ---------- */
function sortableTable(table, cols, rows, opts = {}) {
  const st = table._sort || { key: opts.sortKey || cols[0].key, dir: opts.dir || -1 };
  table._sort = st;
  const thead = table.tHead || table.createTHead();
  const tbody = table.tBodies[0] || table.createTBody();
  thead.replaceChildren(el('tr', {}, cols.map((c) => {
    const th = el('th', { class: c.num ? 'num' : null, scope: 'col', title: c.title || null }, c.label);
    if (st.key === c.key) th.setAttribute('aria-sort', st.dir > 0 ? 'ascending' : 'descending');
    th.onclick = () => { st.dir = st.key === c.key ? -st.dir : (c.num ? -1 : 1); st.key = c.key; sortableTable(table, cols, rows, opts); };
    return th;
  })));
  const col = cols.find((c) => c.key === st.key) || cols[0];
  const get = col.sort || ((r) => r[col.key]);
  const sorted = [...rows].sort((a, b) => {
    const x = get(a), y = get(b);
    if (typeof x === 'number' && typeof y === 'number') return (x - y) * st.dir;
    return String(x ?? '').localeCompare(String(y ?? ''), 'ko') * st.dir;
  });
  tbody.replaceChildren(...sorted.map((r) => {
    const tr = el('tr', { class: opts.rowClass ? opts.rowClass(r) : null });
    for (const c of cols) tr.append(td(c.render ? c.render(r) : c.fmt ? c.fmt(r[c.key]) : r[c.key], c.num));
    return tr;
  }));
}

/* ---------- 가로 막대 ---------- */
function hbars(container, items, { valueFmt = won, split = null } = {}) {
  if (!items.length) { container.replaceChildren(el('p', { class: 'note' }, '표시할 자료가 없습니다.')); return; }
  const max = Math.max(...items.map((d) => Math.abs(d.value)), 1e-12);
  // 양수·음수가 섞일 때만 가운데 축, 한쪽 부호뿐이면 왼쪽 기준(크기 비교), 색으로 부호 구분
  const mixed = items.some((d) => d.value < 0) && items.some((d) => d.value > 0);
  const zero = mixed ? 50 : 0; const span = mixed ? 50 : 100;
  container.replaceChildren(...items.map((d) => {
    const track = el('div', { class: 'track' });
    if (mixed) track.append(el('div', { class: 'axis', style: `left:${zero}%` }));
    const w = (Math.abs(d.value) / max) * span;
    const neg = d.value < 0;
    const left = mixed && neg ? zero - w : zero;
    if (split && d.parts) {
      let x = left;
      const parts = neg && mixed ? [...d.parts].reverse() : d.parts;
      parts.forEach((p, i) => {
        const pw = (Math.abs(p) / max) * span;
        const isDirect = (neg && mixed) ? i === parts.length - 1 : i === 0;
        const outer = (neg && mixed) ? i === 0 : i === parts.length - 1;
        const rad = outer ? ((neg && mixed) ? '4px 0 0 4px' : '0 4px 4px 0') : '0';
        track.append(el('div', { class: `seg ${isDirect ? 'direct' : 'indirect'}`, style: `left:${x}%;width:${pw}%;border-radius:${rad};${i ? 'margin-left:2px;' : ''}` }));
        x += pw;
      });
    } else {
      track.append(el('div', { class: `seg ${neg ? 'negc' : 'pos'}${neg && mixed ? ' neg' : ''}`, style: `left:${left}%;width:${w}%` }));
    }
    return el('div', { class: 'hbar', title: `${d.label}: ${valueFmt(d.value)}` },
      el('div', { class: 'lab' }, d.label), track, el('div', { class: 'val' }, valueFmt(d.value)));
  }));
}

function tiles(container, list) {
  container.replaceChildren(...list.map(([k, v, u]) => el('div', { class: 'tile' },
    el('div', { class: 'k' }, k), el('div', { class: 'v' }, v), el('div', { class: 'u' }, u || ''))));
}

/* ---------- 산업 선택(검색 + 목록) ---------- */
function sectors(level = state.level) { return state.info[level]?.sectors || []; }
function fillSectorSelect(select, query, level = state.level, keep) {
  const prev = keep ?? select.value;
  const list = sectors(level).filter((s) => match(s, query));
  select.replaceChildren(...list.map((s) => el('option', { value: s.code }, `${s.code} ${s.name}${s.active ? '' : ' (산출 0)'}`)));
  if (list.some((s) => s.code === prev)) select.value = prev;
}
function sectorByCode(code, level = state.level) { return sectors(level).find((s) => s.code === code); }

/* =========================================================================
   ① 자료 불러오기
   ========================================================================= */
function renderLevelCards() {
  $('level-cards').replaceChildren(...OFFICIAL.map((lv) => {
    const L = LEVELS[lv]; const info = state.info[lv];
    const card = el('div', { class: 'lcard', id: 'card-' + lv });
    const btn = el('button', { type: 'button', disabled: !state.ready || state.loading || null }, info ? '다시 불러오기' : `${L.label} 불러오기`);
    btn.onclick = () => loadLevels([lv]);
    let stateLine = el('div', { class: 'state note' }, '아직 불러오지 않았습니다.');
    let detail = null;
    if (info) {
      const resid = Object.entries(info.residuals || {});
      const maxr = resid.length ? Math.max(...resid.map(([, v]) => v)) : null;
      const ok = maxr !== null && maxr < 1e-8;
      stateLine = el('div', { class: 'state' },
        el('span', { class: 'badge ' + (ok ? 'ok' : 'warn') }, ok ? '✓ 검증 통과' : '확인 필요'),
        ` ${info.n}개 부문 · ${info.sheets.length}개 시트`);
      detail = el('details', {}, el('summary', {}, '검증 상세(최대 오차)'),
        el('pre', {}, resid.map(([k, v]) => `${k}: ${v.toExponential(1)}`).join('\n') +
          (info.warnings.length ? '\n⚠ ' + info.warnings.join('\n⚠ ') : '') + '\n허용 기준: 1e-8 미만'));
    }
    card.append(el('h3', {}, L.label), el('div', { class: 'meta' }, `${L.n}개 부문 · ${L.note}`), stateLine, btn, detail);
    return card;
  }));
}

async function fetchBytes(url, label) {
  setStatus(`${label} 엑셀 내려받는 중…`);
  const r = await fetch(url);
  if (!r.ok) throw Error(`${label} 엑셀 다운로드 실패 (HTTP ${r.status})`);
  return r.arrayBuffer();
}

async function ensureClassification() {
  if (state.classLoaded) return;
  const buf = await fetchBytes(CLASS_FILE, '부문분류표');
  await send('bytes', { fn: 'load_classification', buffer: buf }, [buf]);
  state.classLoaded = true;
}

async function loadLevels(levels, customBuffer) {
  if (state.loading) return;
  state.loading = true; renderLevelCards(); lockLoadButtons(true);
  try {
    await ensureClassification();
    for (const lv of levels) {
      const buf = customBuffer || await fetchBytes(LEVELS[lv].file, LEVELS[lv].label);
      setStatus(`${LEVELS[lv].label}: 엑셀 읽기·계수 재계산·검증 중… (큰 파일은 시간이 걸립니다)`);
      const info = await send('bytes', { fn: 'load_level', level: lv, buffer: buf }, [buf]);
      state.info[lv] = info;
      for (const d of info.derived || []) state.info[d] = null;
      state.cache = {};
      renderLevelCards();
    }
    await refreshDerived();
    updateLevelSelect(levels[levels.length - 1]);
    setStatus(`준비 완료 · 불러온 분류: ${OFFICIAL.filter((l) => state.info[l]).map(lab).join(', ')} (+13·7·1부문 자동 생성)`);
  } catch (e) {
    setStatus('불러오기 실패: ' + e.message);
  } finally {
    state.loading = false; renderLevelCards(); lockLoadButtons(false);
  }
}

async function refreshDerived() {
  for (const d of ['s13', 's7', 's1']) {
    try { state.info[d] = await call('level_summary_json', d); } catch { state.info[d] = undefined; }
  }
}

function lockLoadButtons(lock) {
  $('load-all').disabled = lock || !state.ready;
  $('upload-file').disabled = lock || !state.ready;
}

function updateLevelSelect(prefer) {
  const avail = ORDER.filter((l) => state.info[l]);
  const sel = $('level');
  sel.replaceChildren(...avail.map((l) => el('option', { value: l }, `${lab(l)} (${state.info[l].n}개 부문)`)));
  sel.disabled = !avail.length;
  const want = state.level && avail.includes(state.level) ? state.level : (avail.includes(prefer) ? prefer : avail[0]);
  if (want) { sel.value = want; setLevel(want); }
  // 원자료 탐색: 원본 엑셀이 있는 분류만
  const ex = $('ex-level'); const prevEx = ex.value;
  const exAvail = OFFICIAL.filter((l) => state.info[l]);
  ex.replaceChildren(...exAvail.map((l) => el('option', { value: l }, lab(l))));
  if (exAvail.includes(prevEx)) ex.value = prevEx;
  ex.disabled = !exAvail.length;
  for (const id of ['find-q', 'st-q', 'st-metric', 'sh-q', 'sh-sector', 'sh-mode', 'sh-amount', 'sh-add', 'sh-preset',
    'lk-q', 'lk-quad', 'fd-q', 'fd-sort', 'pr-kind', 'pr-pct', 'pr-q', 'pr-sector', 'pr-add', 'pr-run',
    'lv-parent-level', 'lv-parent', 'lv-child-level', 'lv-amount', 'lv-run', 'st-csv', 'fd-csv']) $(id).disabled = !avail.length;
  renderSheetSelect();
  state.tree = [];
}

function setLevel(level) {
  const changed = state.level !== level;
  state.level = level;
  if (changed) {
    state.basket = []; renderBasket(); state.priceTargets = []; renderPriceTargets();
    $('sh-results').hidden = true; $('pr-results').hidden = true;
  }
  fillSectorSelect($('sh-sector'), $('sh-q').value);
  fillSectorSelect($('pr-sector'), $('pr-q').value);
  renderCurrentTab();
}

/* =========================================================================
   ② 원자료 탐색
   ========================================================================= */
let findTimer;
$('find-q').oninput = () => {
  clearTimeout(findTimer);
  findTimer = setTimeout(async () => {
    const qv = $('find-q').value.trim();
    if (!qv) { $('find-rows').replaceChildren(); return; }
    const rows = await call('search', { q: qv });
    $('find-rows').replaceChildren(...(rows.length ? rows.map((r) => el('tr', {},
      td(r.levelLabel), td(r.code), td(r.name), td(won(r.x), true),
      td(r.path.filter((p) => !p.level.startsWith('1부문')).map((p) => `${p.level} ${p.code} ${p.name}`).join(' › ')))) :
      [el('tr', {}, el('td', { colspan: 5, class: 'note' }, '일치하는 산업이 없습니다.'))]));
  }, 200);
};

const CHILD = { s1: 'large', large: 'medium', medium: 'small', small: 'basic' };
async function renderTree() {
  const path = state.tree.length ? state.tree : [{ level: 's1', code: '1', name: '전산업' }];
  state.tree = path;
  const cur = path[path.length - 1];
  const child = CHILD[cur.level];
  $('tree-path').replaceChildren(...path.flatMap((p, i) => {
    const b = el('button', { type: 'button', class: i === path.length - 1 ? null : 'ghost' }, i === 0 ? '대분류 전체' : `${p.code} ${p.name}`);
    b.onclick = () => { state.tree = path.slice(0, i + 1); renderTree(); };
    return i ? [el('span', {}, '›'), b] : [b];
  }));
  if (!child) { $('tree-rows').replaceChildren(el('tr', {}, el('td', { colspan: 6, class: 'note' }, '기본부문이 가장 세분된 단계입니다.'))); return; }
  if (!state.info[child]) { $('tree-rows').replaceChildren(el('tr', {}, el('td', { colspan: 6, class: 'note' }, `${lab(child)} 자료를 먼저 불러오세요.`))); return; }
  const rows = await call('children', { parentLevel: cur.level, code: cur.code, childLevel: child });
  $('tree-rows').replaceChildren(...rows.map((r) => {
    const name = el('span', { class: CHILD[child] ? 'clickable' : null }, r.name);
    if (CHILD[child]) name.onclick = () => { state.tree = [...path, { level: child, code: r.code, name: r.name }]; renderTree(); };
    return el('tr', {}, td(r.code), el('td', {}, name), td(won(r.x), true), td(fmt(r.prodMult, 3), true), td(fmt(r.vaMult, 3), true), td(fmt(r.impMult, 3), true));
  }));
}

function renderSheetSelect() {
  const lv = $('ex-level').value; const info = state.info[lv];
  const sel = $('ex-sheet'); const prev = sel.value;
  const sheets = info?.sheets || [];
  sel.replaceChildren(...sheets.map((s) => el('option', { value: s.name }, `${s.name} — ${s.unit}`)));
  if (sheets.some((s) => s.name === prev)) sel.value = prev;
  sel.disabled = !sheets.length;
  for (const id of ['ex-row-search', 'ex-row', 'ex-col-search']) $(id).disabled = !sheets.length;
  if (sheets.length) explore(); else { $('ex-cells').replaceChildren(); $('ex-desc').textContent = ''; }
}
let exData = null; let exSeq = 0;
async function explore() {
  const my = ++exSeq;
  try {
    const d = await call('explore', { level: $('ex-level').value, sheet: $('ex-sheet').value, row: $('ex-row').value });
    if (my !== exSeq) return;
    exData = d;
    $('ex-desc').textContent = d.desc ? `${d.desc}  [${d.meta.join(' · ')}]` : d.meta.join(' · ');
    filterExRows(d.selected);
    $('ex-unit').textContent = '이 행의 단위: ' + (d.unit || '시트 제목 참조');
    $('ex-caption').textContent = `${lab($('ex-level').value)} · ${$('ex-sheet').value} · 행 ${d.selected} ${d.name} (${d.cells.length}개 열)`;
    renderExCells();
  } catch (e) { setStatus(e.message); }
}
function filterExRows(selected) {
  const sel = $('ex-row'); const list = (exData?.rows || []).filter((r) => match(r, $('ex-row-search').value));
  sel.replaceChildren(...list.map((r) => el('option', { value: r.code }, `${r.code} ${r.name}`)));
  if (list.some((r) => r.code === selected)) sel.value = selected;
}
function renderExCells() {
  if (!exData) return;
  const cu = exData.colUnits || {};
  const rowUnit = (exData.unit || '').replace(' (요약 행)', '');
  $('ex-cells').replaceChildren(...exData.cells.filter((c) => match(c, $('ex-col-search').value)).map((c) =>
    el('tr', {}, td(c.code), td(c.name), td(c.value === null ? '빈 셀' : typeof c.value === 'number' ? fmt(c.value, Math.abs(c.value) < 100 && !Number.isInteger(c.value) ? 4 : 0) : c.value, true),
      td(cu[c.code] || rowUnit, false, 'note'))));
}
$('ex-level').onchange = renderSheetSelect;
$('ex-sheet').onchange = () => { $('ex-row-search').value = ''; $('ex-row').replaceChildren(); explore(); };
$('ex-row').onchange = explore;
$('ex-row-search').oninput = () => { filterExRows(); if ($('ex-row').value) explore(); };
$('ex-col-search').oninput = renderExCells;

/* =========================================================================
   ③ 산업 구조
   ========================================================================= */
const ST_COLS = [
  { key: 'code', label: '코드' }, { key: 'name', label: '산업' },
  { key: 'x', label: '총산출(백만원)', num: true, fmt: won },
  { key: 'xShare', label: '산출비중(%)', num: true, fmt: (v) => fmt(v, 2) },
  { key: 'va', label: '부가가치(백만원)', num: true, fmt: won },
  { key: 'vaRate', label: '부가가치율(%)', num: true, fmt: (v) => fmt(v, 1), title: '부가가치 ÷ 총투입' },
  { key: 'midRate', label: '중간투입률(%)', num: true, fmt: (v) => fmt(v, 1) },
  { key: 'impDep', label: '수입의존도(%)', num: true, fmt: (v) => fmt(v, 1), title: '수입 중간재 ÷ 총투입' },
  { key: 'localization', label: '국산화율(%)', num: true, fmt: (v) => fmt(v, 1), title: '국산 중간재 ÷ 중간재' },
  { key: 'wageShare', label: '부가가치 중 피용자보수(%)', num: true, fmt: (v) => fmt(v, 1) },
  { key: 'exportRate', label: '수출률(%)', num: true, fmt: (v) => fmt(v, 1), title: '수출 ÷ 총산출' },
  { key: 'importRate', label: '수입률(%)', num: true, fmt: (v) => fmt(v, 1), title: '수입 ÷ 총공급' },
  { key: 'prodMult', label: '생산유발계수', num: true, fmt: (v) => fmt(v, 3) },
  { key: 'vaMult', label: '부가가치유발계수', num: true, fmt: (v) => fmt(v, 3) },
  { key: 'impMult', label: '수입유발계수', num: true, fmt: (v) => fmt(v, 3) },
];
async function renderStructure() {
  const d = await cached('overview', () => call('overview', state.level));
  const e = d.economy;
  tiles($('st-tiles'), [
    ['총산출', fmt(e.x_total, 1), '조원'], ['부가가치율', fmt(e.va_rate, 1), '% (부가가치/총투입)'],
    ['중간투입률', fmt(e.mid_rate, 1), '%'], ['수입의존도', fmt(e.imp_dep, 1), '% (수입중간재/총투입)'],
    ['수출률', fmt(e.export_rate, 1), '% (수출/총산출)'], ['생산유발계수', fmt(e.prod_mult, 3), '최종수요 1원당, 전체 평균'],
  ]);
  const rows = d.rows.filter((r) => match(r, $('st-q').value));
  const metric = $('st-metric').value; const mcol = ST_COLS.find((c) => c.key === metric);
  const top = [...rows].filter((r) => r.active).sort((a, b) => b[metric] - a[metric]).slice(0, 15);
  $('st-chart-title').textContent = `${mcol.label} 상위 15개 산업 (${lab(state.level)})`;
  hbars($('st-chart'), top.map((r) => ({ label: `${r.code} ${r.name}`, value: r[metric] })), { valueFmt: mcol.fmt });
  sortableTable($('st-table'), ST_COLS, rows, { sortKey: 'x', rowClass: (r) => (r.active ? null : 'muted') });
  $('st-csv').onclick = () => downloadCSV(`io2023_${state.level}_structure.csv`, ST_COLS.map((c) => c.label), d.rows.map((r) => ST_COLS.map((c) => r[c.key])));
}
$('st-q').oninput = () => renderStructure();
$('st-metric').onchange = () => renderStructure();

/* =========================================================================
   ④ 수요충격 시뮬레이터
   ========================================================================= */
function itemsList() { return state.info[state.level]?.items || []; }
$('sh-q').oninput = () => fillSectorSelect($('sh-sector'), $('sh-q').value);
$('sh-mode').onchange = () => {
  const pct = $('sh-mode').value === 'pct';
  $('sh-item-wrap').hidden = !pct;
  $('sh-item').replaceChildren(...itemsList().map((it, i) => el('option', { value: i }, it.name)));
  if (pct) $('sh-item').value = String(itemsList().length - 1);
  $('sh-amount-label').textContent = pct ? '변화율 (%)' : '변화액 (백만원)';
  $('sh-amount').value = pct ? '-10' : '1000000';
};
function renderBasket() {
  $('sh-basket').replaceChildren(...state.basket.map((s, i) => {
    const b = el('button', { type: 'button', class: 'ghost' }, '삭제');
    b.onclick = () => { state.basket.splice(i, 1); renderBasket(); $('sh-results').hidden = true; };
    return el('li', {}, el('span', {}, `${s.code} ${s.name} · ${s.desc} → ${s.amount >= 0 ? '+' : ''}${won(s.amount)}백만원`), b);
  }));
  $('sh-run').disabled = !state.basket.length;
  $('sh-status').textContent = state.basket.length ? `${state.basket.length}개 입력. '파급효과 계산'을 누르세요.` : '시나리오를 추가하세요.';
}
$('sh-form').onsubmit = (e) => {
  e.preventDefault();
  const s = sectorByCode($('sh-sector').value); const v = Number($('sh-amount').value);
  if (!s || !Number.isFinite(v) || v === 0) { $('sh-status').textContent = '산업과 0이 아닌 값을 확인하세요.'; return; }
  let amount = v; let desc = '최종수요 직접 입력';
  if ($('sh-mode').value === 'pct') {
    const k = Number($('sh-item').value); const base = s.fd[k];
    if (!base) { $('sh-status').textContent = `${s.name}의 ${itemsList()[k].name}(국산)이 0이라 %로 계산할 수 없습니다.`; return; }
    amount = base * v / 100; desc = `${itemsList()[k].name} ${won(base)}백만원의 ${v}%`;
  }
  state.basket.push({ code: s.code, name: s.name, amount, desc });
  renderBasket(); $('sh-results').hidden = true;
};
$('sh-clear').onclick = () => { state.basket = []; renderBasket(); $('sh-results').hidden = true; };
$('sh-preset').onclick = () => {
  const cand = sectors().find((s) => /집적회로|반도체/.test(s.name)) || sectors().find((s) => /컴퓨터, 전자/.test(s.name));
  if (!cand) return;
  const k = itemsList().findIndex((it) => it.code === '9140');
  state.basket = [{ code: cand.code, name: cand.name, amount: cand.fd[k] * -0.1, desc: `수출 ${won(cand.fd[k])}백만원의 -10%` }];
  renderBasket(); $('sh-run').click();
};
$('sh-run').onclick = async () => {
  if (!state.basket.length) return;
  const snap = JSON.stringify(state.basket); $('sh-run').disabled = true; $('sh-status').textContent = '계산 중…';
  try {
    const r = await call('shock', { level: state.level, shocks: state.basket.map(({ code, amount }) => ({ code, amount })) });
    if (snap !== JSON.stringify(state.basket)) { $('sh-status').textContent = '입력이 바뀌었습니다. 다시 계산하세요.'; return; }
    state.shockResult = r; renderShock(); $('sh-status').textContent = `계산 완료 (${lab(state.level)} 기준).`;
  } catch (e) { $('sh-status').textContent = '계산 실패: ' + e.message; } finally { $('sh-run').disabled = false; }
};
function renderShock() {
  const r = state.shockResult; if (!r) return;
  $('sh-results').hidden = false;
  const m = (v) => (v === null ? '부호 혼재' : fmt(v, 3) + '배');
  tiles($('sh-tiles'), [
    ['최종수요 변화', won(r.demand), '백만원'], ['생산유발액', won(r.production), `백만원 · 생산유발 ${m(r.prodMult)}`],
    ['부가가치유발액', won(r.valueAdded), `백만원 · ${m(r.vaMult)}`], ['수입유발액', won(r.imports), `백만원 · ${m(r.impMult)}`],
    ['피용자보수 유발', won(r.wage), '백만원 (부가가치 중 임금)'], ['간접효과', won(r.indirect), '백만원 (생산 − 직접)'],
  ]);
  $('sh-note').textContent = `검산: 부가가치유발 + 수입유발 = 최종수요 변화 (오차 ${r.check.toExponential(1)}). 국산품 1원의 수요는 결국 국내 부가가치와 수입으로 모두 나뉩니다. 1조원 = 1,000,000백만원.` + (r.mixedSign ? ' 증가·감소가 섞여 있어 배수는 표시하지 않습니다.' : '');
  hbars($('sh-chart'), r.rows.slice(0, 15).map((x) => ({
    label: `${x.code} ${x.name}`, value: x.production,
    parts: (x.direct >= 0 && x.indirect >= 0) || (x.direct <= 0 && x.indirect <= 0) ? [x.direct, x.indirect] : null,
  })), { split: true });
  const g = $('sh-groups'); g.replaceChildren();
  for (const [lv, list] of Object.entries(r.groups || {})) {
    if (!list.length) continue;
    g.append(el('h3', {}, `${lab(lv)}로 묶어 본 결과 `, el('span', { class: 'sub' }, '— 백만원')),
      el('div', { class: 'scroll short' }, el('table', {}, el('thead', {}, el('tr', {}, el('th', {}, '코드'), el('th', {}, lab(lv)), el('th', { class: 'num' }, '생산유발'), el('th', { class: 'num' }, '부가가치'), el('th', { class: 'num' }, '수입'), el('th', { class: 'num' }, '생산 비중(%)'))),
        el('tbody', {}, list.map((x) => el('tr', {}, td(x.code), td(x.name), td(won(x.production), true), td(won(x.valueAdded), true), td(won(x.imports), true), td(fmt(x.production / r.production * 100, 1), true)))))));
  }
  $('sh-rows').replaceChildren(...r.rows.map((x) => el('tr', {}, td(x.code), td(x.name), td(won(x.direct), true), td(won(x.indirect), true), td(won(x.production), true), td(won(x.valueAdded), true), td(won(x.wage), true), td(won(x.imports), true))));
}
$('sh-csv').onclick = () => {
  const r = state.shockResult; if (!r) return;
  downloadCSV(`io2023_${r.level}_shock.csv`, ['코드', '산업', '직접효과_백만원', '1차간접(Ad·f)_백만원', '간접효과_백만원', '생산유발_백만원', '부가가치유발_백만원', '피용자보수유발_백만원', '수입유발_백만원'],
    r.rows.map((x) => [x.code, x.name, x.direct, x.first, x.indirect, x.production, x.valueAdded, x.wage, x.imports]));
};

/* =========================================================================
   ⑤ 전후방 연관
   ========================================================================= */
const QNAME = { 'Ⅰ': 'Ⅰ 전후방 모두 높음', 'Ⅱ': 'Ⅱ 후방만 높음', 'Ⅲ': 'Ⅲ 전방만 높음', 'Ⅳ': 'Ⅳ 둘 다 낮음' };
const LK_COLS = [
  { key: 'code', label: '코드' }, { key: 'name', label: '산업' },
  { key: 'bl', label: '영향력계수(후방)', num: true, fmt: (v) => fmt(v, 3) },
  { key: 'fl', label: '감응도계수(전방)', num: true, fmt: (v) => fmt(v, 3) },
  { key: 'quad', label: '유형', render: (r) => QNAME[r.quad] },
  { key: 'colSum', label: '생산유발계수(열합)', num: true, fmt: (v) => fmt(v, 3) },
  { key: 'rowSum', label: '행합', num: true, fmt: (v) => fmt(v, 3) },
  { key: 'vaMult', label: '부가가치유발계수', num: true, fmt: (v) => fmt(v, 3) },
  { key: 'impMult', label: '수입유발계수', num: true, fmt: (v) => fmt(v, 3) },
  { key: 'x', label: '총산출(백만원)', num: true, fmt: won },
];
async function renderLinkage() {
  const d = await cached('linkage', () => call('linkage', state.level));
  tiles($('lk-tiles'), Object.entries(d.counts).map(([k, v]) => [QNAME[k], String(v), '개 산업']));
  $('lk-check').textContent = Object.keys(d.check).length ? '원표 대조: ' + Object.entries(d.check).map(([k, v]) => `${k} 최대오차 ${v.toExponential(1)}`).join(' · ') : '';
  const rows = d.rows.filter((r) => match(r, $('lk-q').value) && (!$('lk-quad').value || r.quad === $('lk-quad').value));
  drawScatter(d.rows.filter((r) => r.active), new Set(rows.map((r) => r.code)), rows.length !== d.rows.length);
  sortableTable($('lk-table'), LK_COLS, rows, { sortKey: 'bl', rowClass: (r) => (r.active ? null : 'muted') });
}
$('lk-q').oninput = () => renderLinkage();
$('lk-quad').onchange = () => renderLinkage();
function drawScatter(rows, hl, filtered) {
  const svg = $('lk-svg'); const W = svg.clientWidth || 900; const H = 440;
  svg.setAttribute('viewBox', `0 0 ${W} ${H}`);
  const m = { l: 56, r: 18, t: 18, b: 44 };
  const xs = rows.map((r) => r.bl); const ys = rows.map((r) => r.fl);
  const x0 = Math.min(0.6, ...xs) - 0.02, x1 = Math.max(1.4, ...xs) + 0.02;
  const ly = (v) => Math.log10(Math.max(v, 0.05));
  const y0 = ly(Math.min(0.4, ...ys)) - 0.03, y1 = ly(Math.max(2.5, ...ys)) + 0.03;
  const X = (v) => m.l + (v - x0) / (x1 - x0) * (W - m.l - m.r);
  const Y = (v) => H - m.b - (ly(v) - y0) / (y1 - y0) * (H - m.t - m.b);
  const NS = 'http://www.w3.org/2000/svg';
  const s = (tag, a) => { const e = document.createElementNS(NS, tag); for (const k in a) e.setAttribute(k, a[k]); return e; };
  const kids = [];
  for (let v = Math.ceil(x0 * 10) / 10; v <= x1; v += 0.1) {
    const vv = Math.round(v * 10) / 10;
    kids.push(s('line', { x1: X(vv), x2: X(vv), y1: m.t, y2: H - m.b, class: 'gridline' }));
    const t = s('text', { x: X(vv), y: H - m.b + 16, 'text-anchor': 'middle' }); t.textContent = vv.toFixed(1); kids.push(t);
  }
  for (const v of [0.3, 0.5, 0.7, 1, 1.5, 2, 3, 5, 8, 12]) {
    if (ly(v) < y0 || ly(v) > y1) continue;
    kids.push(s('line', { x1: m.l, x2: W - m.r, y1: Y(v), y2: Y(v), class: 'gridline' }));
    const t = s('text', { x: m.l - 6, y: Y(v) + 4, 'text-anchor': 'end' }); t.textContent = v; kids.push(t);
  }
  kids.push(s('line', { x1: X(1), x2: X(1), y1: m.t, y2: H - m.b, class: 'refline' }));
  kids.push(s('line', { x1: m.l, x2: W - m.r, y1: Y(1), y2: Y(1), class: 'refline' }));
  const ql = [['Ⅰ 전후방 모두 높음', W - m.r - 6, m.t + 14, 'end'], ['Ⅱ 후방만 높음', W - m.r - 6, H - m.b - 8, 'end'],
    ['Ⅲ 전방만 높음', m.l + 6, m.t + 14, 'start'], ['Ⅳ 둘 다 낮음', m.l + 6, H - m.b - 8, 'start']];
  for (const [t, x, y, a] of ql) { const e = s('text', { x, y, 'text-anchor': a, class: 'qlab' }); e.textContent = t; kids.push(e); }
  const xt = s('text', { x: (W + m.l) / 2, y: H - 6, 'text-anchor': 'middle' }); xt.textContent = '영향력계수 (후방연관, 평균=1) →'; kids.push(xt);
  const yt = s('text', { x: 14, y: (H - m.b + m.t) / 2, transform: `rotate(-90 14 ${(H - m.b + m.t) / 2})`, 'text-anchor': 'middle' }); yt.textContent = '감응도계수 (전방연관, 로그) →'; kids.push(yt);
  const maxx = Math.max(...rows.map((r) => r.x));
  const tip = $('lk-tip');
  const ordered = [...rows].sort((a, b) => b.x - a.x);
  for (const r of ordered) {
    const rad = 4 + 10 * Math.sqrt(r.x / maxx);
    const on = filtered && hl.has(r.code);
    kids.push(s('circle', { cx: X(r.bl), cy: Y(r.fl), r: rad, class: on ? 'hl' : '' }));
  }
  for (const r of ordered) {
    const hit = s('circle', { cx: X(r.bl), cy: Y(r.fl), r: Math.max(9, 4 + 10 * Math.sqrt(r.x / maxx)), class: 'hit', tabindex: 0, 'aria-label': `${r.name} 영향력 ${r.bl.toFixed(3)} 감응도 ${r.fl.toFixed(3)}` });
    const show = () => {
      tip.hidden = false;
      tip.replaceChildren(el('b', {}, `${r.code} ${r.name}`), el('br'), `영향력계수 ${fmt(r.bl, 3)} · 감응도계수 ${fmt(r.fl, 3)}`, el('br'), `${QNAME[r.quad]} · 총산출 ${won(r.x)}백만원`);
      const bx = svg.getBoundingClientRect(); const sc = bx.width / W;
      tip.style.left = Math.min(X(r.bl) * sc + 12, bx.width - 290) + 'px'; tip.style.top = Math.max(Y(r.fl) * sc - 60, 4) + 'px';
    };
    hit.onmouseenter = show; hit.onfocus = show; hit.onmouseleave = () => { tip.hidden = true; }; hit.onblur = () => { tip.hidden = true; };
    kids.push(hit);
  }
  svg.replaceChildren(...kids);
}

/* =========================================================================
   ⑥ 최종수요 유발구조
   ========================================================================= */
async function renderFD() {
  const d = await cached('fd', () => call('fd_structure', state.level));
  const head = ['항목', '최종수요(국산)', '최종수요(수입)', '생산유발액', '부가가치유발액', '수입유발액', '생산유발계수', '부가가치유발계수', '수입유발계수', '생산유발비중(%)', '부가가치유발비중(%)', '수입유발비중(%)'];
  const t = $('fd-items');
  t.tHead.replaceChildren(el('tr', {}, head.map((h, i) => el('th', { class: i ? 'num' : null }, h))));
  const showBok = d.bokLevel;
  const row = (r, strong) => {
    const b = showBok && r.bok;
    const coef = (v, i) => el('td', { class: 'num' }, fmt(v, 3), b ? el('div', { class: 'note' }, `한은 ${r.bok[i].toFixed(3)}`) : null);
    return el('tr', { style: strong ? 'font-weight:600;background:var(--surface-2)' : null },
      td(r.name || r.group), td(won(r.fd), true), td(won(r.fm), true), td(won(r.prod), true), td(won(r.va), true), td(won(r.imp), true),
      coef(r.prodMult, 0), coef(r.vaMult, 1), coef(r.impMult, 2),
      td(fmt(r.prodShare, 1), true), td(fmt(r.vaShare, 1), true), td(fmt(r.impShare, 1), true));
  };
  t.tBodies[0].replaceChildren(...d.items.map((r) => row(r)), ...d.groups.map((g) => row(g, true)));
  $('fd-bok').textContent = showBok ? "각 계수 아래 '한은'은 보도자료 공표치입니다. 한국은행은 최종수요 항목별 유발계수를 대분류(33부문) 표로 계산해 공표하므로, 대분류 기준에서 일치합니다."
    : '최종수요 항목별 유발계수는 분류 수준에 따라 조금씩 달라집니다(합계는 동일). 한은 공표치와 대조하려면 분석 기준을 대분류로 바꾸세요. 생산·부가가치 유발계수 = 유발액 ÷ 국산 최종수요, 수입유발계수 = 수입유발액 ÷ (국산+수입) 최종수요.';
  $('fd-legend').replaceChildren(...[['c1', '소비(민간·정부)'], ['c2', '투자(고정자본형성·재고 등)'], ['c3', '수출']].map(([c, n]) => el('span', {}, el('i', { class: c }), n)));
  const key = $('fd-sort').value;
  const rows = d.rows.filter((r) => r.x > 0 && match(r, $('fd-q').value)).sort((a, b) => key === 'x' ? b.x - a.x : b.dep[key] - a.dep[key]);
  $('fd-chart').replaceChildren(...rows.map((r) => {
    const parts = [['소비', 'c1'], ['투자', 'c2'], ['수출', 'c3']];
    const bar = el('div', { class: 'sbar' }, parts.map(([k, c]) => el('span', { class: c, style: `width:${Math.max(r.dep[k], 0)}%`, title: `${k} ${fmt(r.dep[k], 1)}%` })));
    return el('div', { class: 'srow', title: `${r.name} — 소비 ${fmt(r.dep['소비'], 1)}% · 투자 ${fmt(r.dep['투자'], 1)}% · 수출 ${fmt(r.dep['수출'], 1)}%` },
      el('div', { class: 'lab' }, `${r.code} ${r.name}`), el('div', {}, bar, el('div', { class: 'note', style: 'font-size:.75rem' }, `소비 ${fmt(r.dep['소비'], 1)} · 투자 ${fmt(r.dep['투자'], 1)} · 수출 ${fmt(r.dep['수출'], 1)} (%)`)));
  }));
  $('fd-csv').onclick = () => downloadCSV(`io2023_${state.level}_fd_dependence.csv`, ['코드', '산업', '총산출_백만원', '소비의존도_%', '투자의존도_%', '수출의존도_%'],
    d.rows.map((r) => [r.code, r.name, r.x, r.dep['소비'], r.dep['투자'], r.dep['수출']]));
}
$('fd-q').oninput = () => renderFD();
$('fd-sort').onchange = () => renderFD();

/* =========================================================================
   ⑦ 가격 파급
   ========================================================================= */
$('pr-q').oninput = () => fillSectorSelect($('pr-sector'), $('pr-q').value);
$('pr-kind').onchange = () => {
  const k = $('pr-kind').value;
  $('pr-target-box').hidden = k === 'wage';
  $('pr-target-note').textContent = k === 'import' ? '수입품 전체에 적용합니다. 특정 상품만 올리려면 아래에서 추가하세요(예: 원유, 천연가스).'
    : '가격을 올릴 산업을 추가하세요(예: 대분류 D 전력·가스·증기, 기본부문 4502 화력). 여러 개면 모두 같은 상승률을 적용합니다.';
  state.priceTargets = []; renderPriceTargets(); $('pr-results').hidden = true;
};
function renderPriceTargets() {
  $('pr-targets').replaceChildren(...state.priceTargets.map((s, i) => {
    const b = el('button', { type: 'button', class: 'ghost' }, '삭제');
    b.onclick = () => { state.priceTargets.splice(i, 1); renderPriceTargets(); };
    return el('li', {}, `${s.code} ${s.name}`, b);
  }));
}
$('pr-add').onclick = () => {
  const s = sectorByCode($('pr-sector').value);
  if (s && !state.priceTargets.some((t) => t.code === s.code)) { state.priceTargets.push({ code: s.code, name: s.name }); renderPriceTargets(); }
};
$('pr-form').onsubmit = async (e) => {
  e.preventDefault();
  const kind = $('pr-kind').value; const pct = Number($('pr-pct').value);
  if (!Number.isFinite(pct)) { $('pr-status').textContent = '상승률을 확인하세요.'; return; }
  if (kind === 'sector' && !state.priceTargets.length) { $('pr-status').textContent = '가격을 올릴 산업을 한 개 이상 추가하세요.'; return; }
  $('pr-status').textContent = '계산 중…';
  try {
    const r = await call('price', { level: state.level, kind, pct, targets: state.priceTargets.map((t) => ({ code: t.code, pct })) });
    r.label = kind === 'import' ? `수입품${state.priceTargets.length ? '(' + state.priceTargets.map((t) => t.name).join(', ') + ')' : ' 전체'} 가격 ${pct}% 상승`
      : kind === 'wage' ? `임금 ${pct}% 상승` : `${state.priceTargets.map((t) => t.name).join(', ')} 가격 ${pct}% 인상`;
    state.priceResult = r; renderPrice();
    $('pr-status').textContent = `계산 완료: ${r.label} (${lab(state.level)} 기준)`;
  } catch (err) { $('pr-status').textContent = '계산 실패: ' + err.message; }
};
function renderPrice() {
  const r = state.priceResult; if (!r) return;
  $('pr-results').hidden = false;
  tiles($('pr-tiles'), [
    ['소비자물가(민간소비 가중)', fmt(r.index.consumer, 2), '% 상승 (수입 소비재 직접효과 포함)'],
    ['국내 생산자가격(산출 가중)', fmt(r.index.output, 2), '% 상승'],
    ['수출품 가격', fmt(r.index.export, 2), '% 상승 (수출 가중)'],
    ['투자재 가격', fmt(r.index.investment, 2), '% 상승 (고정자본형성 가중)'],
  ]);
  hbars($('pr-chart'), r.rows.slice(0, 20).map((x) => ({ label: `${x.code} ${x.name}${x.exo ? ' (외생 인상)' : ''}`, value: x.dp })), { valueFmt: (v) => fmt(v, 2) + '%' });
  $('pr-rows').replaceChildren(...r.rows.map((x) => el('tr', {}, td(x.code), td(x.name + (x.exo ? ' (외생)' : '')), td(fmt(x.dp, 3), true), td(fmt(x.consW, 2), true), td(fmt(x.consContrib, 3), true))));
}
$('pr-csv').onclick = () => {
  const r = state.priceResult; if (!r) return;
  downloadCSV(`io2023_${state.level}_price.csv`, ['코드', '산업', '가격상승률_%', '외생여부', '민간소비비중_%', '소비자물가기여_%p'], r.rows.map((x) => [x.code, x.name, x.dp, x.exo ? 'Y' : '', x.consW, x.consContrib]));
};

/* =========================================================================
   ⑧ 분류 수준 비교
   ========================================================================= */
async function renderLevels() {
  const d = await call('compare', {});
  const t = $('lv-summary');
  t.tHead.replaceChildren(el('tr', {}, ['분류', '부문 수', '자료', '총산출(조원)', '생산유발계수(전체)', '부가가치유발계수(전체)', '수입유발계수(전체)', '부문별 생산유발계수 최소', '최대', '표준편차'].map((h, i) => el('th', { class: i > 2 ? 'num' : null }, h))));
  const src = (s) => (s === 'official' ? '한은 공식 파일' : s === 'generated' ? '기본부문 합산 파일' : '대분류 합산(자동)');
  t.tBodies[0].replaceChildren(...d.summary.map((s) => el('tr', {}, td(s.label), td(s.n, true), td(src(s.source)), td(fmt(s.x / 1e6, 1), true),
    td(fmt(s.prodMult, 4), true), td(fmt(s.vaMult, 4), true), td(fmt(s.impMult, 4), true), td(fmt(s.min, 3), true), td(fmt(s.max, 3), true), td(fmt(s.sd, 3), true))));
  $('lv-consistency').textContent = (d.consistency.length ? '파일 간 일관성(기본부문을 합산한 값 vs 각 분류 파일): ' + d.consistency.map((c) => `${c.level} ${c.matched}/${c.n}개 부문 일치, 총산출 최대 상대오차 ${c.xMaxRel.toExponential(1)}, 생산유발계수 최대차 ${c.multMaxAbs.toExponential(1)}`).join(' · ') : '')
    + ' 경제 전체 유발계수가 모든 분류에서 같은 것은, 실제 최종수요 전체가 유발한 생산이 분류와 상관없이 총산출과 같기 때문입니다.';
  const lvAvail = ORDER.filter((l) => state.info[l]);
  const pl = $('lv-parent-level'); const prev = pl.value;
  pl.replaceChildren(...lvAvail.filter((l) => l !== 'basic').map((l) => el('option', { value: l }, lab(l))));
  if ([...pl.options].some((o) => o.value === prev)) pl.value = prev; else if ([...pl.options].some((o) => o.value === 'large')) pl.value = 'large';
  fillLvParent();
}
function fillLvParent() {
  const pl = $('lv-parent-level').value;
  fillSectorSelect($('lv-parent'), '', pl, $('lv-parent').value || (pl === 'large' ? 'C09' : undefined));
  const cl = $('lv-child-level'); const prev = cl.value;
  const finer = ORDER.slice(0, ORDER.indexOf(pl)).filter((l) => state.info[l]);
  cl.replaceChildren(...finer.map((l) => el('option', { value: l }, lab(l))));
  if (finer.includes(prev)) cl.value = prev; else if (finer.includes('basic')) cl.value = 'basic';
}
$('lv-parent-level').onchange = fillLvParent;
$('lv-run').onclick = async () => {
  const box = $('lv-detail');
  try {
    const d = await call('compare', { parentLevel: $('lv-parent-level').value, code: $('lv-parent').value, childLevel: $('lv-child-level').value, amount: Number($('lv-amount').value) });
    const x = d.detail; const p = x.parent;
    const diffs = x.rows.map((r) => r.production);
    box.replaceChildren(
      el('div', { class: 'tiles' },
        el('div', { class: 'tile' }, el('div', { class: 'k' }, `${p.label} '${p.name}'에 한꺼번에 줄 때`), el('div', { class: 'v' }, won(p.production)), el('div', { class: 'u' }, `백만원 · 생산유발계수 ${fmt(p.prodMult, 3)}`)),
        el('div', { class: 'tile' }, el('div', { class: 'k' }, `${x.childLabel} ${x.rows.length}개에 실제 최종수요 비중대로 나눌 때`), el('div', { class: 'v' }, won(x.mixProduction)), el('div', { class: 'u' }, `백만원 · 평균 ${fmt(x.mixMult, 3)}`)),
        el('div', { class: 'tile' }, el('div', { class: 'k' }, `하위 부문 하나에 몰아줄 때 범위`), el('div', { class: 'v' }, `${won(Math.min(...diffs))} ~ ${won(Math.max(...diffs))}`), el('div', { class: 'u' }, '백만원')),
      ),
      el('p', { class: 'note' }, `같은 ${won(x.amount)}백만원이라도 하위 부문 중 어디에 수요가 생기느냐에 따라 생산유발계수가 ${fmt(Math.min(...x.rows.map((r) => r.prodMult)), 3)}~${fmt(Math.max(...x.rows.map((r) => r.prodMult)), 3)}로 달라지고, 생산유발액은 최대 ${won(Math.max(...diffs) - Math.min(...diffs))}백만원 차이 납니다. 상위 분류 하나로 계산하면 이 차이가 평균에 묻힙니다(집계 편의). 정책·기업 분석에서 세분류가 필요한 이유입니다.`),
      el('div', { class: 'scroll short' }, el('table', {}, el('thead', {}, el('tr', {}, ['코드', x.childLabel, '최종수요 비중(%)', '생산유발계수', '부가가치유발계수', '수입유발계수', '몰아줄 때 생산유발(백만원)'].map((h, i) => el('th', { class: i > 1 ? 'num' : null }, h)))),
        el('tbody', {}, x.rows.map((r) => el('tr', {}, td(r.code), td(r.name), td(fmt(r.weight, 1), true), td(fmt(r.prodMult, 3), true), td(fmt(r.vaMult, 3), true), td(fmt(r.impMult, 3), true), td(won(r.production), true)))))));
  } catch (e) { box.replaceChildren(el('p', { class: 'bad' }, e.message)); }
};

/* =========================================================================
   ⑨ 검증
   ========================================================================= */
async function renderVerify() {
  const d = await cached('overview', () => call('overview', state.level));
  const judge = (diff, bok, unit) => {
    const tol = unit === '배' ? 0.0005 : 0.05 + 1e-9;
    const ok = Math.abs(diff) <= tol * (unit === '조원' ? 2 : 1);
    return el('span', { class: 'badge ' + (ok ? 'ok' : 'warn') }, ok ? '✓ 반올림 범위 일치' : '차이 있음');
  };
  $('vf-economy').tBodies[0].replaceChildren(...d.bench.map((b) => el('tr', {}, td(b.label), td(fmt(b.bok, b.unit === '배' ? 3 : 1), true), td(fmt(b.calc, b.unit === '배' ? 4 : 2), true), td(fmt(b.diff, b.unit === '배' ? 4 : 2), true), td(b.unit), el('td', {}, judge(b.diff, b.bok, b.unit)))));
  const s = await call('bok_sector_check', '');
  $('vf-sector').tBodies[0].replaceChildren(...s.map((r) => {
    const ok = r.bok.every((v, i) => Math.abs(v - r.calc[i]) < 0.0015);
    return el('tr', {}, td(r.level), td(r.name), td(r.bok.map((v) => v.toFixed(3)).join(' / '), true), td(r.calc.map((v) => v.toFixed(3)).join(' / '), true),
      el('td', {}, el('span', { class: 'badge ' + (ok ? 'ok' : 'warn') }, ok ? '✓ 일치' : '근사(재구성)')));
  }));
  $('vf-files').replaceChildren(...OFFICIAL.filter((l) => state.info[l]).map((l) => {
    const r = state.info[l].residuals;
    return el('details', { class: 'box' }, el('summary', {}, `${lab(l)} (${state.info[l].n}개 부문) — 최대 오차 ${Math.max(...Object.values(r)).toExponential(1)}`),
      el('table', {}, el('tbody', {}, Object.entries(r).map(([k, v]) => el('tr', {}, td(k), td(v.toExponential(2), true))))));
  }));
}

/* =========================================================================
   탭 전환
   ========================================================================= */
const RENDER = { explore: renderTree, structure: renderStructure, linkage: renderLinkage, fd: renderFD, levels: renderLevels, verify: renderVerify };
function showTab(name) {
  if (!document.getElementById(name)) name = 'data';
  state.tab = name;
  for (const s of document.querySelectorAll('section.tab')) s.hidden = s.id !== name;
  for (const a of document.querySelectorAll('#tabs a')) {
    if (a.dataset.tab === name) a.setAttribute('aria-current', 'page'); else a.removeAttribute('aria-current');
  }
  renderCurrentTab();
}
async function renderCurrentTab() {
  const f = RENDER[state.tab];
  if (!f || !state.level) return;
  try { await f(); } catch (e) { setStatus('표시 오류: ' + e.message); }
}
window.addEventListener('hashchange', () => showTab(location.hash.slice(1)));
$('level').onchange = () => setLevel($('level').value);
$('load-all').onclick = () => loadLevels(OFFICIAL);
$('upload-file').onchange = async () => {
  const f = $('upload-file').files[0]; if (!f) return;
  const lv = $('upload-level').value;
  await loadLevels([lv], await f.arrayBuffer());
  $('upload-file').value = '';
};
let resizeT; window.addEventListener('resize', () => { clearTimeout(resizeT); resizeT = setTimeout(() => { if (state.tab === 'linkage') renderLinkage(); }, 200); });

/* ---------- 시작 ---------- */
(async function init() {
  showTab(location.hash.slice(1) || 'data');
  renderLevelCards();
  startWorker();
  try {
    await send('boot');
    state.ready = true; renderLevelCards(); lockLoadButtons(false);
    setStatus("계산 환경 준비 완료. '4개 분류 모두 불러오기' 또는 분류별 버튼을 눌러 엑셀 자료를 불러오세요.");
  } catch (e) {
    setStatus('계산 환경 준비 실패: ' + e.message + ' — 인터넷 연결을 확인하고 새로고침하세요.');
  }
})();
