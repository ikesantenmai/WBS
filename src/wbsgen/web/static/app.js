/* wbsgen — 空の WBS を作る / 記入済み Excel をガントチャートで見る
 *
 * 画面が持つのは「用紙の指定」と描画だけ。日程表の組み立ても Excel の読み取りも
 * サーバ側で行い、返ってきたモデルをそのまま表と SVG に起こす。
 */
'use strict';

// Excel の列幅に合わせたプレビューの列幅 (px)
const TABLE_COLUMNS = [
  { key: 'group', label: '大項目', width: 74, cls: 'group' },
  { key: 'subgroup', label: '中項目', width: 74, cls: 'subgroup' },
  { key: 'no', label: '項番', width: 38 },
  { key: 'name', label: '項目', width: 168, cls: 'name' },
  { key: 'start', label: '開始日', width: 46, group: '予定', plan: true },
  { key: 'days', label: '日数', width: 38, group: '予定', plan: true },
  { key: 'end', label: '終了日', width: 46, group: '予定', plan: true },
  { key: 'actual_start', label: '開始', width: 42, group: '実績' },
  { key: 'actual_days', label: '日数', width: 38, group: '実績' },
  { key: 'actual_end', label: '終了', width: 42, group: '実績' },
  { key: 'delay', label: '遅れ', width: 38, group: '実績' },
  { key: 'progress', label: '進捗', width: 38, group: '実績' },
  { key: 'effort', label: '工数', width: 38 },
  { key: 'predecessor', label: '先行', width: 38 },
  { key: 'member', label: '担当', width: 56, cls: 'member' },
  { key: 'status', label: '状態', width: 62, cls: 'status' },
];

const CHART_WIDTH = { day: 22, week: 42, month: 58 };
const ROW_H = 24;
const HEAD_H = 24;

//: 空行のプレビューはこれ以上描いても読めないので省略する
const PREVIEW_ROW_LIMIT = 60;

let meta = null;
let workdays = ['mon', 'tue', 'wed', 'thu', 'fri'];
let mode = 'blank';        // 'blank' = 新規作成 / 'chart' = 読み込んだ WBS
let pending = null;
let loadedFile = null;     // 読み込み中のファイル (単位を変えて読み直すため)

const $ = (sel) => document.querySelector(sel);
const el = (tag, attrs = {}, ...kids) => {
  const node = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs)) {
    if (k === 'class') node.className = v;
    else if (k === 'text') node.textContent = v;
    else if (k.startsWith('on')) node.addEventListener(k.slice(2), v);
    else if (v !== null && v !== undefined && v !== false) node.setAttribute(k, v);
  }
  for (const kid of kids) if (kid) node.append(kid);
  return node;
};
const svg = (tag, attrs = {}) => {
  const node = document.createElementNS('http://www.w3.org/2000/svg', tag);
  for (const [k, v] of Object.entries(attrs)) {
    if (v !== null && v !== undefined) node.setAttribute(k, String(v));
  }
  return node;
};

// ---------------------------------------------------------------- 通信
function banner(message, ok = false) {
  const box = $('#banner');
  if (!message) { box.hidden = true; box.textContent = ''; return; }
  box.textContent = message;
  box.className = ok ? 'banner ok' : 'banner';
  box.hidden = false;
  if (ok) setTimeout(() => { if (box.textContent === message) banner(''); }, 3500);
}

async function api(path, options = {}) {
  const response = await fetch(path, options);
  if (!response.ok) {
    let detail = `${response.status} ${response.statusText}`;
    try {
      const body = await response.json();
      if (body && body.detail) {
        detail = typeof body.detail === 'string' ? body.detail : JSON.stringify(body.detail);
      }
    } catch (_) { /* JSON でないレスポンス */ }
    throw new Error(detail);
  }
  return response;
}

function currentSpec() {
  const selected = document.querySelector('input[name=mode]:checked').value;
  const spec = {
    title: $('#title').value.trim(),
    start: $('#start').value,
    unit: $('#unit').value,
    rows: Number($('#rows').value) || 0,
    workdays,
    japanese_holidays: $('#jp-holidays').checked,
    holidays: parseDates($('#holidays').value),
    members: $('#members').value.split('\n').map((s) => s.trim()).filter(Boolean),
  };
  if (selected === 'months') spec.months = Number($('#months').value) || 12;
  else spec.end = $('#end').value;
  return spec;
}

function parseDates(text) {
  return String(text || '')
    .split(/[,\s、]+/)
    .map((s) => s.trim())
    .filter(Boolean);
}

async function refresh() {
  if (mode !== 'blank') return;
  if (pending) pending.abort();
  const controller = new AbortController();
  pending = controller;
  try {
    const response = await api('/api/preview', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(currentSpec()),
      signal: controller.signal,
    });
    const model = await response.json();
    banner('');
    render({
      title: model.spec.title,
      timeline: model.timeline,
      rows: [],
      blankRows: model.spec.rows,
      nowX: null,
      totals: null,
      warnings: [],
      info: `${model.spec.start} 〜 ${model.spec.end}`
            + ` / ${model.timeline.columns.length} 列 / 空行 ${model.spec.rows} 行`,
    });
    $('#holiday-note').textContent = model.holidays.length
      ? `この期間の祝日・休業日は ${model.holidays.length} 日です。`
      : 'この期間に祝日・休業日はありません。';
    $('#btn-build').disabled = false;
  } catch (error) {
    if (error.name === 'AbortError') return;
    banner(error.message);
    $('#preview-info').textContent = '';
    $('#sheet').replaceChildren(el('p', { class: 'empty-note', text: '指定を直してください。' }));
    $('#btn-build').disabled = true;
  } finally {
    if (pending === controller) pending = null;
  }
}

let timer = null;
function scheduleRefresh(delay = 250) {
  clearTimeout(timer);
  timer = setTimeout(refresh, delay);
}

// ---------------------------------------------------------------- 描画
/**
 * 表 (左・固定) と日程表 (右・SVG) を並べて描く。
 * 記入済みを読み込んだときは日程表にバーが載る。
 */
function render(view) {
  const { timeline } = view;
  const colW = CHART_WIDTH[timeline.unit] || 42;
  const withWeekday = timeline.unit === 'day';
  const headH = HEAD_H * (withWeekday ? 3 : 2);
  const rowCount = view.rows.length
    || Math.min(view.blankRows || 0, PREVIEW_ROW_LIMIT);
  const bodyH = Math.max(rowCount, 1) * ROW_H;
  const width = timeline.columns.length * colW;

  const grid = el('div', { class: 'grid' },
    el('div', { class: 'table-wrap' }, buildTable(view, rowCount, withWeekday)),
    el('div', { class: 'chart-wrap' },
      chartHead(timeline, colW, width, headH, withWeekday),
      chartBody(view, colW, width, bodyH, rowCount)),
  );

  const parts = [el('div', { class: 'sheet-title', text: view.title })];
  if (view.warnings && view.warnings.length) parts.push(warningBox(view.warnings));
  parts.push(grid);
  $('#sheet').replaceChildren(...parts);

  $('#preview-info').textContent = view.info;
  renderTotals(view.totals);
}

function warningBox(warnings) {
  return el('div', { class: 'warnings' },
    el('strong', { text: '読み飛ばした行があります' }),
    el('ul', {}, ...warnings.slice(0, 5).map((w) => el('li', { text: w }))));
}

// ---------------------------------------------------------------- 表
function buildTable(view, rowCount, withWeekday) {
  const cols = el('colgroup');
  for (const column of TABLE_COLUMNS) cols.append(el('col', { style: `width:${column.width}px` }));

  const head = el('thead');
  const bandRow = el('tr');
  let i = 0;
  while (i < TABLE_COLUMNS.length) {
    const group = TABLE_COLUMNS[i].group || '';
    let span = 1;
    while (i + span < TABLE_COLUMNS.length
           && (TABLE_COLUMNS[i + span].group || '') === group) span++;
    bandRow.append(el('th', { colspan: span, text: group }));
    i += span;
  }
  head.append(bandRow);
  head.append(el('tr', {}, ...TABLE_COLUMNS.map((c) => el('th', { text: c.label }))));
  if (withWeekday) {
    head.append(el('tr', {}, ...TABLE_COLUMNS.map(() => el('th', {}))));
  }

  const body = el('tbody');
  for (let r = 0; r < rowCount; r++) {
    const row = view.rows[r] || null;
    const tr = el('tr');
    for (const column of TABLE_COLUMNS) tr.append(tableCell(row, column, view.rows[r - 1]));
    body.append(tr);
  }
  return el('table', { class: 'wbs' }, cols, head, body);
}

function tableCell(row, column, previous) {
  const td = el('td', { class: `${column.cls || ''}${column.plan ? ' plan' : ''}` });
  if (!row) return td;

  let value = row[column.key];
  // 大項目・中項目は変わったところにだけ出す (Excel と同じ見え方)
  if ((column.key === 'group' || column.key === 'subgroup')
      && previous && previous[column.key] === value) value = '';

  if (column.key === 'status' && value) {
    td.append(el('span', {
      text: value,
      style: row.status_bg ? `background:${row.status_bg};color:${row.status_fg}` : '',
    }));
    return td;
  }
  if (column.key === 'progress') value = value == null ? '' : `${Math.round(value * 100)}%`;
  else if (['days', 'actual_days', 'delay'].includes(column.key)) {
    value = value == null ? '' : `${value} 日`;
  } else if (/start|end$/.test(column.key)) {
    if (column.key === 'end' && row.end_derived) td.classList.add('derived');
    value = shortDate(value);
  } else if (value == null) value = '';

  td.textContent = value;
  if (column.key === 'end' && row.end_derived) td.title = '日数から補った終了日';
  return td;
}

function shortDate(iso) {
  if (!iso) return '';
  const [, m, d] = iso.split('-');
  return `${Number(m)}/${d}`;
}

// ---------------------------------------------------------------- 日程表の見出し
function chartHead(timeline, colW, width, height, withWeekday) {
  const root = svg('svg', {
    class: 'chart-head', width, height, viewBox: `0 0 ${width} ${height}`,
    'data-unit': timeline.unit,
  });

  // 上段: 月 (月表示なら年)。Excel と同じく区切りの列にだけ置く。
  root.append(svg('rect', { x: 0, y: 0, width, height: HEAD_H, fill: 'var(--month-band)' }));
  for (const band of timeline.bands) {
    const x = band.start * colW;
    root.append(svg('line', {
      x1: x, y1: 0, x2: x, y2: height, stroke: '#b8791f', 'stroke-width': 1,
    }));
    const label = svg('text', {
      x: x + colW / 2, y: HEAD_H / 2 + 4, 'text-anchor': 'middle',
      'font-size': 11, 'font-weight': 700, fill: '#fff',
    });
    label.textContent = colW >= 26 ? band.label : '';
    root.append(label);
  }

  // 下段: 週の開始日 (日表示なら日、月表示なら月)
  root.append(svg('rect', {
    x: 0, y: HEAD_H, width, height: height - HEAD_H, fill: 'var(--band)',
  }));
  timeline.columns.forEach((column, i) => {
    if (column.rest) {
      root.append(svg('rect', {
        x: i * colW, y: HEAD_H, width: colW, height: height - HEAD_H,
        fill: column.saturday ? 'var(--saturday)' : 'var(--holiday)',
      }));
    }
    const label = svg('text', {
      x: i * colW + colW / 2, y: HEAD_H + HEAD_H / 2 + 4, 'text-anchor': 'middle',
      'font-size': 10, fill: column.rest ? '#c0392b' : '#4a4a2c',
    });
    label.textContent = colW >= 18 ? column.label : '';
    root.append(label);

    if (withWeekday && colW >= 18) {
      const day = svg('text', {
        x: i * colW + colW / 2, y: HEAD_H * 2 + HEAD_H / 2 + 4, 'text-anchor': 'middle',
        'font-size': 10,
        fill: column.rest ? '#c0392b' : (column.saturday ? '#0070c0' : '#8a8a70'),
      });
      day.textContent = column.weekday;
      root.append(day);
    }
  });

  for (let y = HEAD_H; y <= height; y += HEAD_H) {
    root.append(svg('line', {
      x1: 0, y1: y, x2: width, y2: y, stroke: '#b8791f', 'stroke-width': 1,
    }));
  }
  return root;
}

// ---------------------------------------------------------------- 日程表の本体
function chartBody(view, colW, width, height, rowCount) {
  const { timeline } = view;
  const root = svg('svg', {
    class: 'chart-body', width, height, viewBox: `0 0 ${width} ${height}`,
  });

  // 休日の列と罫線
  timeline.columns.forEach((column, i) => {
    if (column.rest) {
      root.append(svg('rect', {
        x: i * colW, y: 0, width: colW, height,
        fill: column.saturday ? 'var(--saturday)' : 'var(--holiday)',
      }));
    }
    root.append(svg('line', {
      x1: i * colW, y1: 0, x2: i * colW, y2: height,
      stroke: 'var(--line-soft)', 'stroke-width': 1,
    }));
  });
  for (let r = 0; r <= rowCount; r++) {
    root.append(svg('line', {
      x1: 0, y1: r * ROW_H, x2: width, y2: r * ROW_H,
      stroke: 'var(--line-soft)', 'stroke-width': 1,
    }));
  }

  view.rows.forEach((row, i) => root.append(...bars(row, i, colW)));

  if (view.nowX != null) {
    root.append(svg('line', {
      x1: view.nowX * colW, y1: 0, x2: view.nowX * colW, y2: height,
      stroke: 'var(--now)', 'stroke-width': 1.5, 'stroke-dasharray': '4 3',
    }));
  }
  return root;
}

/** 1 行ぶんのバー (予定・進捗・実績)。 */
function bars(row, index, colW) {
  const out = [];
  const y = index * ROW_H;
  if (!row.plan && !row.actual) return out;

  const withActual = !!row.actual;
  const planY = withActual ? y + 3 : y + 5;
  const planH = withActual ? ROW_H / 2 - 4 : ROW_H - 10;

  if (row.plan) {
    const x = row.plan.x1 * colW;
    const w = Math.max(row.plan.x2 * colW - x, 2);
    const bar = svg('rect', {
      x, y: planY, width: w, height: planH, rx: 2,
      fill: row.color, 'fill-opacity': 0.9,
      stroke: shade(row.color, 0.5), 'stroke-width': 0.75,
    });
    bar.append(tooltip(row));
    out.push(bar);

    if (row.progress) {
      out.push(svg('rect', {
        x, y: planY + 2, width: Math.max(w * Math.min(row.progress, 1), 1),
        height: planH - 4, rx: 1, fill: shade(row.color, 0.55),
      }));
    }
  }
  if (row.actual) {
    const x = row.actual.x1 * colW;
    const w = Math.max(row.actual.x2 * colW - x, 2);
    const done = row.progress != null && row.progress >= 1;
    const bar = svg('rect', {
      x, y: y + ROW_H / 2 + 1, width: w, height: ROW_H / 2 - 4, rx: 2,
      fill: shade(row.color, 1.35), 'fill-opacity': done ? 0.95 : 0.65,
      stroke: shade(row.color, 0.5), 'stroke-width': 0.75,
    });
    bar.append(tooltip(row));
    out.push(bar);
  }
  return out;
}

function tooltip(row) {
  const node = svg('title');
  const lines = [row.name];
  if (row.start) lines.push(`予定 ${row.start} 〜 ${row.end || '-'}`);
  if (row.actual_start) lines.push(`実績 ${row.actual_start} 〜 ${row.actual_end || '進行中'}`);
  if (row.progress != null) lines.push(`進捗 ${Math.round(row.progress * 100)}%`);
  if (row.member) lines.push(`担当 ${row.member}`);
  if (row.status) lines.push(`状態 ${row.status}`);
  node.textContent = lines.join('\n');
  return node;
}

function shade(hex, factor) {
  const value = hex.replace('#', '');
  const rgb = [0, 2, 4].map((i) => parseInt(value.slice(i, i + 2), 16));
  const out = rgb.map((c) => (factor <= 1
    ? Math.round(c * factor)
    : Math.round(c + (255 - c) * (factor - 1))));
  return `#${out.map((c) => Math.max(0, Math.min(255, c)).toString(16).padStart(2, '0')).join('')}`;
}

// ---------------------------------------------------------------- 集計
function renderTotals(totals) {
  const box = $('#totals');
  if (!totals) { box.hidden = true; box.replaceChildren(); return; }
  const items = [
    ['行数', `${totals.rows}`],
    ['完了', `${totals.done}`],
    ['進行中', `${totals.running}`],
    ['遅延', `${totals.delayed}`, totals.delayed > 0],
    ['全体進捗', `${Math.round(totals.progress * 100)}%`],
    ['工数', `${totals.effort} 人日`],
    ['期間', `${totals.first_day ?? '-'} 〜 ${totals.last_day ?? '-'}`],
  ];
  box.replaceChildren(...items.map(([term, value, warn]) => el('div', {},
    el('dt', { text: term }),
    el('dd', { text: value, class: warn ? 'warn' : '' }))));
  box.hidden = false;
}

// ---------------------------------------------------------------- 読み込み
async function importFile(file, unit = null) {
  const body = new FormData();
  body.append('file', file);
  const query = unit ? `?unit=${encodeURIComponent(unit)}` : '';
  try {
    const response = await api(`/api/import${query}`, { method: 'POST', body });
    const model = await response.json();
    loadedFile = file;
    setMode('chart');
    $('#chart-unit').value = model.timeline.unit;
    render({
      title: model.title,
      timeline: model.timeline,
      rows: model.rows,
      blankRows: 0,
      nowX: model.now_x,
      totals: model.totals,
      warnings: model.warnings,
      info: `${model.timeline.start} 〜 ${model.timeline.end}`
            + ` / ${model.timeline.columns.length} 列 / ${model.rows.length} 行`,
    });
    if (!unit) banner(`${file.name} を読み込みました。`, true);
  } catch (error) {
    banner(`読み込めません: ${error.message}`);
  }
}

function setMode(next) {
  mode = next;
  const chart = next === 'chart';
  // 前の表示の集計・注意書きが残らないように先に消す
  if (!chart) renderTotals(null);
  for (const id of ['#period-section', '#paper-section',
                    '#calendar-section', '#members-section']) {
    $(id).hidden = chart;
  }
  document.querySelector('.layout').classList.toggle('chart-mode', chart);
  $('#chart-unit-field').hidden = !chart;
  $('#btn-back').hidden = !chart;
  $('#btn-export').hidden = !chart;
  $('#btn-build').hidden = chart;
  $('#preview-title').textContent = chart ? 'ガントチャート' : 'プレビュー';
  $('#preview-hint').textContent = chart
    ? '記入済みの Excel を読み込んで表示しています'
    : '実際の Excel と同じ列・同じ見出しです';
  $('#tagline').textContent = chart
    ? '記入済みの WBS をガントチャートで表示しています。'
    : '期間を指定すると、日程表と記入用の空行だけの Excel を書き出します。';
}

// ---------------------------------------------------------------- 入力
function buildForm() {
  $('#unit').replaceChildren(...meta.units.map(
    (u) => el('option', { value: u.value, text: u.label })));
  $('#unit').value = 'week';
  $('#rows').value = meta.default_rows;
  $('#rows').max = meta.max_rows;
  $('#start').value = meta.suggested.start;
  $('#end').value = meta.suggested.end;
  $('#title').value = meta.suggested.title;
  $('#title').placeholder = meta.suggested.title;
  renderWeekdays();
}

function renderWeekdays() {
  $('#workdays').replaceChildren(...meta.weekdays.map(({ value, label }) => el('button', {
    type: 'button',
    text: label,
    'aria-pressed': String(workdays.includes(value)),
    onclick: () => {
      const at = workdays.indexOf(value);
      if (at >= 0) workdays.splice(at, 1); else workdays.push(value);
      renderWeekdays();
      refresh();
    },
  })));
}

function bind() {
  for (const id of ['#start', '#end', '#months', '#unit', '#rows', '#jp-holidays']) {
    $(id).addEventListener('change', () => refresh());
  }
  for (const id of ['#title', '#holidays', '#members']) {
    $(id).addEventListener('input', () => scheduleRefresh(500));
  }
  for (const radio of document.querySelectorAll('input[name=mode]')) {
    radio.addEventListener('change', () => {
      const months = radio.value === 'months' && radio.checked;
      $('#field-end').hidden = months;
      $('#field-months').hidden = !months;
      refresh();
    });
  }
  $('#spec-form').addEventListener('submit', (event) => event.preventDefault());
  $('#btn-build').addEventListener('click', download);
  $('#btn-export').addEventListener('click', exportChart);
  $('#btn-back').addEventListener('click', () => {
    loadedFile = null;
    setMode('blank');
    banner('');
    refresh();
  });
  $('#chart-unit').addEventListener('change', () => {
    if (loadedFile) importFile(loadedFile, $('#chart-unit').value);
  });
  $('#import-file').addEventListener('change', (event) => {
    const file = event.target.files[0];
    if (file) importFile(file);
    event.target.value = '';
  });
}

/** 空の WBS を書き出す。 */
function download() {
  return save($('#btn-build'), '/api/build', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(currentSpec()),
  });
}

/** 表示中のガントチャートを図形として描き込んだ Excel を書き出す。 */
function exportChart() {
  if (!loadedFile) return Promise.resolve();
  const body = new FormData();
  body.append('file', loadedFile);
  const unit = $('#chart-unit').value;
  return save($('#btn-export'), `/api/export?unit=${encodeURIComponent(unit)}`,
              { method: 'POST', body });
}

async function save(button, path, options) {
  button.disabled = true;
  try {
    const response = await api(path, options);
    const blob = await response.blob();
    const disposition = response.headers.get('Content-Disposition') || '';
    const match = /filename\*=UTF-8''([^;]+)/.exec(disposition);
    const name = match ? decodeURIComponent(match[1]) : 'wbs.xlsx';
    const url = URL.createObjectURL(blob);
    const link = el('a', { href: url, download: name });
    document.body.append(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
    banner(`${name} を書き出しました。`, true);
  } catch (error) {
    banner(`書き出せません: ${error.message}`);
  } finally {
    button.disabled = false;
  }
}

async function start() {
  bind();
  try {
    meta = await (await api('/api/meta')).json();
  } catch (error) {
    banner(`起動できません: ${error.message}`);
    return;
  }
  buildForm();
  await refresh();
}

start();
