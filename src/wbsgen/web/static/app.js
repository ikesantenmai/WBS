/* wbsgen — WBS / ガントチャート Web UI
 *
 * 状態は「プロジェクト定義 (state)」だけが真。編集のたびにサーバへ送り、
 * 返ってきた描画モデル (model) で表とチャートを描き直す。日程計算も
 * チャート座標もサーバ側の同じエンジンを通るので、画面と Excel が一致する。
 */
'use strict';

const WEEKDAYS = [
  ['mon', '月'], ['tue', '火'], ['wed', '水'], ['thu', '木'],
  ['fri', '金'], ['sat', '土'], ['sun', '日'],
];

const SHOW_FLAGS = [
  ['results', '実績'], ['progress', '進捗塗り'], ['manpower', '工数'],
  ['status', '状態'], ['predecessor_line', '先行線'], ['now_line', '現在日線'],
  ['inazuma_line', 'イナズマ線'], ['member_color', '担当色'],
];

const STATUS_COLORS = {
  '完了': ['#C0C0C0', '#666'],
  '実行中': ['#FFFF99', '#B36B00'],
  '残り': ['#FFCC00', '#800000'],
  '遅れ': ['#FF99CC', '#D00000'],
  'あと': ['#CCFFFF', '#3D8BC4'],
};

const COLUMN_WIDTH = { day: 22, week: 45, month: 62 };
const ROW_H = 26;
const HEAD_H = 26;

let state = null;    // プロジェクト定義
let model = null;    // サーバが返した描画モデル
let meta = null;
let editingRow = -1;
let pending = null;  // 実行中のプレビュー要求 (AbortController)

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
    if (v !== null && v !== undefined) node.setAttribute(k, v);
  }
  return node;
};

// ---------------------------------------------------------------- 通信
function banner(message, ok = false) {
  const box = $('#banner');
  if (!message) { box.hidden = true; return; }
  box.textContent = message;
  box.className = ok ? 'banner ok' : 'banner';
  box.hidden = false;
  if (ok) setTimeout(() => { if (box.textContent === message) box.hidden = true; }, 3500);
}

async function api(path, options = {}) {
  const response = await fetch(path, options);
  if (!response.ok) {
    let detail = `${response.status} ${response.statusText}`;
    try {
      const body = await response.json();
      if (body && body.detail) {
        detail = typeof body.detail === 'string'
          ? body.detail
          : JSON.stringify(body.detail);
      }
    } catch (_) { /* JSON でないレスポンス */ }
    throw new Error(detail);
  }
  return response;
}

/** 定義をサーバに送り、描画モデルを取り直して描画する。 */
async function refresh() {
  if (pending) pending.abort();
  const controller = new AbortController();
  pending = controller;
  $('#status-chip').textContent = '計算中';
  $('#status-chip').className = 'chip busy';
  try {
    const response = await api('/api/preview', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(state),
      signal: controller.signal,
    });
    model = await response.json();
    banner('');
    render();
  } catch (error) {
    if (error.name === 'AbortError') return;
    banner(`計算できません: ${error.message}`);
    $('#status-chip').textContent = '';
  } finally {
    if (pending === controller) pending = null;
  }
}

let refreshTimer = null;
function scheduleRefresh(delay = 220) {
  clearTimeout(refreshTimer);
  refreshTimer = setTimeout(refresh, delay);
}

// ---------------------------------------------------------------- 設定
function buildSettings() {
  const flags = $('#show-flags');
  flags.replaceChildren(...SHOW_FLAGS.map(([key, label]) => {
    const box = el('input', { type: 'checkbox' });
    box.checked = !!state.chart.show[key];
    box.addEventListener('change', () => {
      state.chart.show[key] = box.checked;
      refresh();
    });
    return el('label', {}, box, el('span', { text: label }));
  }));

  const days = $('#workdays');
  days.replaceChildren(...WEEKDAYS.map(([key, label]) => {
    const on = state.calendar.workdays.includes(key);
    return el('button', {
      type: 'button', text: label, 'aria-pressed': String(on),
      onclick: () => {
        const list = state.calendar.workdays;
        const at = list.indexOf(key);
        if (at >= 0) list.splice(at, 1); else list.push(key);
        buildSettings();
        refresh();
      },
    });
  }));

  $('#jp-holidays').checked = !!state.calendar.japanese_holidays;
  $('#extra-holidays').value = (state.calendar.holidays || []).join(', ');
  $('#extra-workdays').value = (state.calendar.extra_workdays || []).join(', ');
  $('#title').value = state.project.title || '';
  $('#chart-start').value = state.chart.start || '';
  $('#chart-period').value = state.chart.period_days ?? 360;
  $('#chart-unit').value = state.chart.unit || 'week';
  $('#chart-base').value = state.chart.base_date || '';
  $('#th-remain').value = state.chart.thresholds.exec_remain_days;
  $('#th-near').value = state.chart.thresholds.start_near_days;

  buildMembers();
}

function buildMembers() {
  const box = $('#members');
  box.replaceChildren(...state.members.map((member, index) => {
    const name = el('input', { type: 'text', value: member.name, placeholder: '担当者名' });
    name.addEventListener('change', () => { member.name = name.value.trim(); refresh(); });
    const color = el('input', { type: 'color', value: normalizeColor(member.color) });
    color.addEventListener('change', () => { member.color = color.value; refresh(); });
    const remove = el('button', {
      type: 'button', text: '×', title: '削除',
      onclick: () => { state.members.splice(index, 1); buildMembers(); refresh(); },
    });
    return el('div', { class: 'member-row' }, name, color, remove);
  }));

  $('#member-list').replaceChildren(
    ...state.members.map((m) => el('option', { value: m.name })));
}

function normalizeColor(value) {
  const text = String(value || '#4472C4');
  return /^#[0-9a-f]{6}$/i.test(text) ? text : '#4472C4';
}

function parseDates(text) {
  return String(text || '')
    .split(/[,\s、]+/)
    .map((s) => s.trim())
    .filter((s) => /^\d{4}-\d{2}-\d{2}$/.test(s));
}

// ---------------------------------------------------------------- 表
const COLUMNS = [
  { key: 'group', label: '大項目', width: 78, cls: 'group' },
  { key: 'subgroup', label: '中項目', width: 78, cls: 'subgroup' },
  { key: 'no', label: '項番', width: 40 },
  { key: 'name', label: '項目', width: 186, cls: 'name' },
  { key: 'start', label: '開始日', width: 48, cls: 'plan num', group: '予定' },
  { key: 'days', label: '日数', width: 40, cls: 'plan num', group: '予定' },
  { key: 'end', label: '終了日', width: 48, cls: 'plan num', group: '予定' },
  { key: 'actual_start', label: '開始', width: 44, cls: 'num', group: '実績', flag: 'results' },
  { key: 'actual_days', label: '日数', width: 40, cls: 'num', group: '実績', flag: 'results' },
  { key: 'actual_end', label: '終了', width: 44, cls: 'num', group: '実績', flag: 'results' },
  { key: 'delay', label: '遅れ', width: 40, cls: 'num delay', group: '実績', flag: 'results' },
  { key: 'progress', label: '進捗', width: 40, cls: 'num', group: '実績' },
  { key: 'effort', label: '工数', width: 40, cls: 'num', flag: 'manpower' },
  { key: 'predecessor', label: '先行', width: 40 },
  { key: 'member', label: '担当', width: 58 },
  { key: 'status', label: '状態', width: 70, cls: 'status', flag: 'status' },
];

/** 表を畳んだときも残す列。 */
const ESSENTIAL = new Set(['no', 'name', 'status']);
let collapsed = false;

function visibleColumns() {
  return COLUMNS.filter((c) => !c.flag || model.show[c.flag]);
}

function renderTable() {
  const columns = visibleColumns();
  const optional = (c) => (collapsed && !ESSENTIAL.has(c.key) ? ' opt' : '');
  const cols = el('colgroup');
  for (const column of columns) {
    cols.append(el('col', { class: optional(column).trim(), style: `width:${column.width}px` }));
  }
  const head = el('thead');

  const bandRow = el('tr');
  let index = 0;
  while (index < columns.length) {
    const group = columns[index].group || '';
    let span = 1;
    while (index + span < columns.length && (columns[index + span].group || '') === group) span++;
    const shown = columns.slice(index, index + span).filter((c) => !optional(c)).length;
    if (shown > 0) bandRow.append(el('th', { colspan: shown, text: group }));
    index += span;
  }
  head.append(bandRow);
  head.append(el('tr', {}, ...columns.map(
    (c) => el('th', { text: c.label, title: c.label, class: optional(c).trim() }))));

  const body = el('tbody');
  model.rows.forEach((row, at) => {
    const tr = el('tr', { draggable: 'true', 'data-index': String(at) });
    if (row.kind === 'summary') tr.classList.add('is-summary');
    else if (row.group && (at === 0 || model.rows[at - 1].group !== row.group)) {
      tr.classList.add('is-group');
    }
    for (const column of columns) {
      const td = cell(row, column, at);
      if (optional(column)) td.classList.add('opt');
      tr.append(td);
    }
    tr.addEventListener('click', () => openTaskDialog(at));
    body.append(tr);
  });

  const table = el('table', { class: 'wbs' }, cols, head, body);
  $('#table-wrap').classList.toggle('collapsed', collapsed);
  enableRowDrag(body);
  $('#table-wrap').replaceChildren(table);
}

function cell(row, column, at) {
  const previous = at > 0 ? model.rows[at - 1] : null;
  let text = row[column.key];

  if (column.key === 'group' && previous && previous.group === row.group) text = '';
  if (column.key === 'subgroup' && previous
      && previous.group === row.group && previous.subgroup === row.subgroup) text = '';
  if (column.key === 'progress') text = text == null ? '' : `${Math.round(text * 100)}%`;
  else if (column.key === 'days' || column.key === 'actual_days' || column.key === 'delay') {
    text = text == null ? '' : `${text} 日`;
  } else if (column.key === 'effort') text = text == null ? '' : String(text);
  else if (/start|end$/.test(column.key)) text = shortDate(text);

  const td = el('td', { class: column.cls || '' });
  if (column.key === 'status' && row.status) {
    const [bg, fg] = statusColor(row.status);
    td.append(el('span', { text: row.status, style: `background:${bg};color:${fg}` }));
  } else {
    td.textContent = text ?? '';
  }
  if (column.key === 'delay' && !row.delay) td.classList.remove('delay');
  return td;
}

function statusColor(status) {
  for (const [key, value] of Object.entries(STATUS_COLORS)) {
    if (status.startsWith(key)) return value;
  }
  return ['transparent', '#b6bbc2'];
}

function shortDate(iso) {
  if (!iso) return '';
  const [, m, d] = iso.split('-');
  return `${Number(m)}/${d}`;
}

// ---------------------------------------------------------------- 並べ替え
function enableRowDrag(body) {
  let from = -1;
  body.addEventListener('dragstart', (event) => {
    const tr = event.target.closest('tr');
    if (!tr) return;
    from = Number(tr.dataset.index);
    tr.classList.add('dragging');
    event.dataTransfer.effectAllowed = 'move';
    event.dataTransfer.setData('text/plain', String(from));
  });
  body.addEventListener('dragover', (event) => {
    event.preventDefault();
    const tr = event.target.closest('tr');
    if (!tr) return;
    body.querySelectorAll('.drop-target').forEach((n) => n.classList.remove('drop-target'));
    tr.classList.add('drop-target');
  });
  body.addEventListener('dragend', () => {
    body.querySelectorAll('.dragging,.drop-target')
      .forEach((n) => n.classList.remove('dragging', 'drop-target'));
  });
  body.addEventListener('drop', (event) => {
    event.preventDefault();
    const tr = event.target.closest('tr');
    if (!tr || from < 0) return;
    const to = Number(tr.dataset.index);
    if (from === to) return;
    const [moved] = state.tasks.splice(from, 1);
    state.tasks.splice(to, 0, moved);
    from = -1;
    refresh();
  });
}

// ---------------------------------------------------------------- チャート
function renderChart() {
  const timeline = model.timeline;
  const colW = COLUMN_WIDTH[timeline.unit] || 45;
  const width = timeline.columns.length * colW;
  const height = HEAD_H * 2 + model.rows.length * ROW_H;
  const root = svg('svg', {
    class: 'gantt', width, height, viewBox: `0 0 ${width} ${height}`,
    'data-unit': timeline.unit,
  });

  root.append(defs());
  root.append(chartBackground(timeline, colW, width, height));
  root.append(chartHeader(timeline, colW));
  root.append(chartBars(colW));
  root.append(chartLinks(colW));
  if (model.show.now_line && model.now_x != null) root.append(nowLine(colW, height));
  if (model.show.inazuma_line) root.append(inazumaLine(colW, height));

  $('#chart-wrap').replaceChildren(root);
  document.documentElement.style.setProperty('--col-w', `${colW}px`);
}

function defs() {
  const node = svg('defs');
  const marker = svg('marker', {
    id: 'arrow', viewBox: '0 0 8 8', refX: '7', refY: '4',
    markerWidth: '6', markerHeight: '6', orient: 'auto-start-reverse',
  });
  marker.append(svg('path', { d: 'M0,1 L7,4 L0,7 z', fill: 'var(--link)' }));
  node.append(marker);
  return node;
}

function chartBackground(timeline, colW, width, height) {
  const g = svg('g');
  const top = HEAD_H * 2;
  timeline.columns.forEach((column, i) => {
    const fill = column.rest
      ? (column.saturday ? 'var(--saturday)' : 'var(--holiday)')
      : null;
    if (fill) {
      g.append(svg('rect', { x: i * colW, y: top, width: colW, height: height - top, fill }));
    }
    g.append(svg('line', {
      x1: i * colW, y1: top, x2: i * colW, y2: height,
      stroke: 'var(--line-soft)', 'stroke-width': 1,
    }));
  });
  for (let r = 0; r <= model.rows.length; r++) {
    const y = top + r * ROW_H;
    g.append(svg('line', {
      x1: 0, y1: y, x2: width, y2: y, stroke: 'var(--line-soft)', 'stroke-width': 1,
    }));
  }
  return g;
}

function chartHeader(timeline, colW) {
  const g = svg('g', { class: 'chart-head' });
  g.append(svg('rect', {
    x: 0, y: 0, width: timeline.columns.length * colW, height: HEAD_H * 2,
    fill: 'var(--band)',
  }));
  for (const band of timeline.bands) {
    const x = band.start * colW;
    const w = band.span * colW;
    g.append(svg('line', {
      x1: x, y1: 0, x2: x, y2: HEAD_H * 2, stroke: '#c9ce9a', 'stroke-width': 1,
    }));
    const label = svg('text', {
      x: x + w / 2, y: HEAD_H / 2 + 4, 'text-anchor': 'middle',
      'font-size': 11, 'font-weight': 700, fill: '#3b3b1f',
    });
    label.textContent = w >= 34 ? band.label : '';
    g.append(label);
  }
  g.append(svg('line', {
    x1: 0, y1: HEAD_H, x2: timeline.columns.length * colW, y2: HEAD_H,
    stroke: '#c9ce9a', 'stroke-width': 1,
  }));
  timeline.columns.forEach((column, i) => {
    const label = svg('text', {
      x: i * colW + colW / 2, y: HEAD_H + HEAD_H / 2 + 4, 'text-anchor': 'middle',
      'font-size': 10, fill: column.rest ? '#c0392b' : '#4a4a2c',
    });
    label.textContent = colW >= 18 ? column.label : '';
    g.append(label);
    if (column.weekday && colW >= 18) {
      label.setAttribute('y', HEAD_H + 12);
      const day = svg('text', {
        x: i * colW + colW / 2, y: HEAD_H + 23, 'text-anchor': 'middle',
        'font-size': 9, fill: column.rest ? '#c0392b' : '#8a8a70',
      });
      day.textContent = column.weekday;
      g.append(day);
    }
  });
  return g;
}

function chartBars(colW) {
  const g = svg('g');
  const top = HEAD_H * 2;
  model.rows.forEach((row, i) => {
    const y = top + i * ROW_H;
    if (row.milestone) {
      g.append(milestone(row, row.milestone.x * colW, y));
      return;
    }
    if (!row.plan) return;
    const x = row.plan.x1 * colW;
    const w = Math.max(row.plan.x2 * colW - x, 2);

    if (row.kind === 'summary') {
      g.append(chevron(x, y + 4, w, ROW_H - 8, row.color, row.name));
      return;
    }
    const withActual = !!row.actual;
    const planY = withActual ? y + 3 : y + 6;
    const planH = withActual ? ROW_H / 2 - 4 : ROW_H - 12;
    g.append(bar(x, planY, w, planH, row.color, 0.95));

    if (model.show.progress && row.progress) {
      g.append(svg('rect', {
        x, y: planY + 2, width: Math.max(w * Math.min(row.progress, 1), 1),
        height: planH - 4, fill: shade(row.color, 0.55), rx: 1,
      }));
    }
    if (withActual) {
      const ax = row.actual.x1 * colW;
      const aw = Math.max(row.actual.x2 * colW - ax, 2);
      g.append(bar(ax, y + ROW_H / 2 + 1, aw, ROW_H / 2 - 4,
                   shade(row.color, 1.35), row.progress >= 1 ? 0.95 : 0.65));
    }
    const title = svg('title');
    title.textContent = `${row.name}\n予定 ${row.start ?? '-'} 〜 ${row.end ?? '-'}`
      + (row.actual_start ? `\n実績 ${row.actual_start} 〜 ${row.actual_end ?? '進行中'}` : '')
      + (row.status ? `\n状態 ${row.status}` : '');
    g.lastChild.append(title);
  });
  return g;
}

function bar(x, y, w, h, color, opacity) {
  return svg('rect', {
    x, y, width: w, height: h, rx: 2,
    fill: color, 'fill-opacity': opacity,
    stroke: shade(color, 0.5), 'stroke-width': 0.75,
  });
}

function chevron(x, y, w, h, color, label) {
  const g = svg('g');
  const tip = Math.min(h / 2, w / 2);
  g.append(svg('path', {
    d: `M${x},${y} H${x + w - tip} L${x + w},${y + h / 2} L${x + w - tip},${y + h}`
       + ` H${x} L${x + tip},${y + h / 2} Z`,
    fill: shade(color, 0.7),
  }));
  if (w > 60) {
    const text = svg('text', {
      x: x + w / 2, y: y + h / 2 + 4, 'text-anchor': 'middle',
      'font-size': 10, 'font-weight': 700, fill: '#fff',
    });
    text.textContent = label;
    g.append(text);
  }
  return g;
}

function milestone(row, x, y) {
  const size = 9;
  const cy = y + ROW_H / 2;
  const color = shade(row.color, 0.75);
  const shape = row.milestone.shape;
  let node;
  if (shape === 'ellipse') {
    node = svg('circle', { cx: x, cy, r: size / 2 + 1, fill: color });
  } else if (shape === 'triangle') {
    node = svg('polygon', {
      points: `${x},${cy - size} ${x + size},${cy + size / 2} ${x - size},${cy + size / 2}`,
      fill: color,
    });
  } else if (shape === 'star5') {
    node = svg('polygon', { points: star(x, cy, size + 1, (size + 1) / 2.4), fill: color });
  } else {
    node = svg('polygon', {
      points: `${x},${cy - size} ${x + size},${cy} ${x},${cy + size} ${x - size},${cy}`,
      fill: color,
    });
  }
  node.setAttribute('stroke', shade(row.color, 0.4));
  node.setAttribute('stroke-width', '0.75');
  const g = svg('g');
  g.append(node);
  const label = svg('text', {
    x: x + size + 4, y: cy + 4, 'font-size': 10, fill: '#444',
  });
  label.textContent = row.name;
  g.append(label);
  return g;
}

function star(cx, cy, outer, inner) {
  const points = [];
  for (let i = 0; i < 10; i++) {
    const r = i % 2 ? inner : outer;
    const a = (Math.PI / 5) * i - Math.PI / 2;
    points.push(`${cx + r * Math.cos(a)},${cy + r * Math.sin(a)}`);
  }
  return points.join(' ');
}

function chartLinks(colW) {
  const g = svg('g');
  const top = HEAD_H * 2;
  for (const link of model.links) {
    const y1 = top + (link.from_row - 1) * ROW_H + ROW_H / 2;
    const y2 = top + (link.to_row - 1) * ROW_H + ROW_H / 2;
    const x1 = link.x1 * colW;
    const x2 = link.x2 * colW;
    const mid = x1 + 8;
    g.append(svg('path', {
      d: `M${x1},${y1} H${mid} V${y2} H${Math.max(x2, mid)}`,
      fill: 'none', stroke: 'var(--link)', 'stroke-width': 1,
      'marker-end': 'url(#arrow)',
    }));
  }
  return g;
}

function nowLine(colW, height) {
  const x = model.now_x * colW;
  const g = svg('g');
  g.append(svg('line', {
    x1: x, y1: HEAD_H, x2: x, y2: height,
    stroke: 'var(--now)', 'stroke-width': 1.5, 'stroke-dasharray': '4 3',
  }));
  return g;
}

function inazumaLine(colW, height) {
  if (!model.inazuma.length || model.now_x == null) return svg('g');
  const top = HEAD_H * 2;
  const nowX = model.now_x * colW;
  const points = [`${nowX},${HEAD_H * 2}`];
  for (const point of model.inazuma) {
    points.push(`${point.x * colW},${top + (point.row - 1) * ROW_H + ROW_H / 2}`);
  }
  const last = model.inazuma[model.inazuma.length - 1];
  points.push(`${nowX},${top + (last.row - 1) * ROW_H + ROW_H / 2}`);
  return svg('polyline', {
    points: points.join(' '), fill: 'none',
    stroke: 'var(--inazuma)', 'stroke-width': 1.5, 'stroke-linejoin': 'round',
  });
}

function shade(hex, factor) {
  const value = hex.replace('#', '');
  const rgb = [0, 2, 4].map((i) => parseInt(value.slice(i, i + 2), 16));
  const out = rgb.map((c) => (factor <= 1
    ? Math.round(c * factor)
    : Math.round(c + (255 - c) * (factor - 1))));
  return `#${out.map((c) => Math.max(0, Math.min(255, c)).toString(16).padStart(2, '0')).join('')}`;
}

// ---------------------------------------------------------------- サマリ
function renderSummary() {
  const t = model.totals;
  const items = [
    ['タスク', `${t.tasks} 件`],
    ['完了', `${t.done} 件`],
    ['進行中', `${t.running} 件`],
    ['遅延', `${t.delayed} 件`, t.delayed > 0],
    ['全体進捗', `${Math.round(t.progress * 100)}%`],
    ['工数合計', `${t.effort} 人日`],
    ['開始', t.first_day ?? '-'],
    ['終了', t.last_day ?? '-'],
  ];
  $('#summary').replaceChildren(...items.flatMap(([term, value, warn]) => [
    el('dt', { text: term }),
    el('dd', { text: value, class: warn ? 'warn' : '' }),
  ]));
}

function render() {
  renderTable();
  renderChart();
  renderSummary();
  $('#status-chip').className = 'chip';
  $('#status-chip').textContent = `${model.rows.length} 行 / ${model.timeline.columns.length} 列`;
}

// ---------------------------------------------------------------- 行編集
function openTaskDialog(index) {
  editingRow = index;
  const task = state.tasks[index] || {};
  const form = $('#task-form');
  form.reset();
  for (const [key, value] of Object.entries(task)) {
    const field = form.elements[key];
    if (!field) continue;
    field.value = key === 'progress' && value != null
      ? Math.round(value * 100)
      : (value ?? '');
  }
  form.elements.kind.value = task.kind || 'task';
  $('#task-dialog-title').textContent = task.name ? `編集 — ${task.name}` : '新しい行';
  toggleShapeField();
  $('#task-dialog').showModal();
  form.elements.name.focus();
}

function toggleShapeField() {
  const kind = $('#task-form').elements.kind.value;
  $('#shape-field').style.display = kind === 'milestone' ? '' : 'none';
}

function saveTaskDialog() {
  const form = $('#task-form');
  if (!form.elements.name.value.trim()) return false;
  const task = {};
  for (const field of form.elements) {
    if (!field.name) continue;
    const value = field.value.trim();
    if (!value) continue;
    if (field.name === 'progress') task.progress = Number(value) / 100;
    else if (['days', 'actual_days'].includes(field.name)) task[field.name] = Number(value);
    else if (field.name === 'effort') task.effort = Number(value);
    else task[field.name] = value;
  }
  if (task.kind === 'task') delete task.kind;
  if (task.kind !== 'milestone') delete task.shape;
  state.tasks[editingRow] = task;
  refresh();
  return true;
}

// ---------------------------------------------------------------- 起動
function bindSettings() {
  const wire = (selector, apply, event = 'change') => {
    $(selector).addEventListener(event, () => { apply($(selector)); refresh(); });
  };
  $('#title').addEventListener('input', () => {
    state.project.title = $('#title').value;
    scheduleRefresh(600);
  });
  wire('#chart-start', (n) => { state.chart.start = n.value || null; });
  wire('#chart-period', (n) => { state.chart.period_days = Number(n.value) || 360; });
  wire('#chart-unit', (n) => { state.chart.unit = n.value; });
  wire('#chart-base', (n) => { state.chart.base_date = n.value || null; });
  wire('#th-remain', (n) => { state.chart.thresholds.exec_remain_days = Number(n.value) || 0; });
  wire('#th-near', (n) => { state.chart.thresholds.start_near_days = Number(n.value) || 0; });
  wire('#jp-holidays', (n) => { state.calendar.japanese_holidays = n.checked; });
  wire('#extra-holidays', (n) => { state.calendar.holidays = parseDates(n.value); });
  wire('#extra-workdays', (n) => { state.calendar.extra_workdays = parseDates(n.value); });

  $('#btn-add-member').addEventListener('click', () => {
    state.members.push({ name: `担当${state.members.length + 1}`, color: '#4472C4' });
    buildMembers();
    refresh();
  });
}

function addTask(kind) {
  const last = state.tasks[state.tasks.length - 1] || {};
  const task = { name: '新しいタスク', group: last.group || '', days: 5 };
  if (kind === 'summary') { task.kind = 'summary'; task.name = '新しい工程'; delete task.days; }
  if (kind === 'milestone') {
    task.kind = 'milestone';
    task.shape = 'diamond';
    task.name = '新しいマイルストーン';
    task.days = 1;
  }
  if (!kind) task.start = state.chart.base_date || meta.today;
  state.tasks.push(task);
  refresh().then(() => openTaskDialog(state.tasks.length - 1));
}

function bindActions() {
  $('#btn-add-task').addEventListener('click', () => addTask(null));
  $('#btn-add-summary').addEventListener('click', () => addTask('summary'));
  $('#btn-add-milestone').addEventListener('click', () => addTask('milestone'));

  $('#task-form').addEventListener('submit', (event) => {
    if (!saveTaskDialog()) event.preventDefault();
  });
  $('#task-form').elements.kind.addEventListener('change', toggleShapeField);
  $('#task-cancel').addEventListener('click', () => $('#task-dialog').close());
  $('#task-delete').addEventListener('click', () => {
    state.tasks.splice(editingRow, 1);
    $('#task-dialog').close();
    if (!state.tasks.length) state.tasks.push({ name: '新しいタスク', days: 5 });
    refresh();
  });

  $('#btn-template').addEventListener('click', async () => {
    if (!confirm('現在の内容を破棄して雛形を読み込みます。よろしいですか？')) return;
    try {
      const response = await api('/api/template/standard');
      state = await response.json();
      buildSettings();
      await refresh();
      banner('雛形を読み込みました。', true);
    } catch (error) {
      banner(`雛形を読み込めません: ${error.message}`);
    }
  });

  $('#import-file').addEventListener('change', async (event) => {
    const file = event.target.files[0];
    if (!file) return;
    const body = new FormData();
    body.append('file', file);
    try {
      const response = await api('/api/import', { method: 'POST', body });
      state = await response.json();
      buildSettings();
      await refresh();
      banner(`${file.name} を読み込みました。`, true);
    } catch (error) {
      banner(`読み込めません: ${error.message}`);
    } finally {
      event.target.value = '';
    }
  });

  $('#btn-collapse').addEventListener('click', (event) => {
    collapsed = !collapsed;
    event.currentTarget.setAttribute('aria-pressed', String(collapsed));
    event.currentTarget.textContent = collapsed ? '表を戻す' : '表を畳む';
    $('#table-wrap').classList.toggle('collapsed', collapsed);
    renderTable();
  });
  $('#btn-settings').addEventListener('click', (event) => {
    const hidden = document.querySelector('.layout').classList.toggle('no-settings');
    event.currentTarget.setAttribute('aria-pressed', String(!hidden));
  });

  $('#btn-yaml').addEventListener('click', () => download('/api/export/yaml', 'wbs.yaml'));
  $('#btn-build').addEventListener('click', () => download('/api/build', 'wbs.xlsx'));
}

async function download(path, fallbackName) {
  const button = path.includes('build') ? $('#btn-build') : $('#btn-yaml');
  button.disabled = true;
  try {
    const response = await api(path, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(state),
    });
    const blob = await response.blob();
    const disposition = response.headers.get('Content-Disposition') || '';
    const match = /filename\*=UTF-8''([^;]+)/.exec(disposition);
    const name = match ? decodeURIComponent(match[1]) : fallbackName;
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
  bindSettings();
  bindActions();
  try {
    meta = await (await api('/api/meta')).json();
    state = await (await api('/api/template/standard')).json();
  } catch (error) {
    banner(`起動できません: ${error.message}`);
    return;
  }
  buildSettings();
  await refresh();
}

start();
