/* wbsgen — 空の WBS を作る
 *
 * 画面が持つのは「用紙の指定」だけ。日程表の組み立てはサーバ側の
 * wbsgen.timeline が行い、返ってきた見出しをそのまま表に起こすので、
 * プレビューと書き出される Excel が一致する。
 */
'use strict';

// Excel の列幅 (px) に合わせたプレビューの列幅
const TABLE_COLUMNS = [
  { key: 'group', label: '大項目', width: 74, group: '' },
  { key: 'subgroup', label: '中項目', width: 74, group: '' },
  { key: 'no', label: '項番', width: 38, group: '' },
  { key: 'name', label: '項目', width: 168, group: '' },
  { key: 'start', label: '開始日', width: 46, group: '予定', plan: true },
  { key: 'days', label: '日数', width: 38, group: '予定', plan: true },
  { key: 'end', label: '終了日', width: 46, group: '予定', plan: true },
  { key: 'actual_start', label: '開始', width: 42, group: '実績' },
  { key: 'actual_days', label: '日数', width: 38, group: '実績' },
  { key: 'actual_end', label: '終了', width: 42, group: '実績' },
  { key: 'delay', label: '遅れ', width: 38, group: '実績' },
  { key: 'progress', label: '進捗', width: 38, group: '実績' },
  { key: 'effort', label: '工数', width: 38, group: '' },
  { key: 'predecessor', label: '先行', width: 38, group: '' },
  { key: 'member', label: '担当', width: 56, group: '' },
  { key: 'status', label: '状態', width: 62, group: '' },
];

const CHART_WIDTH = { day: 22, week: 42, month: 58 };

//: プレビューに描く空行の上限 (これ以上は描いても読めないので省略する)
const PREVIEW_ROW_LIMIT = 60;

let meta = null;
let workdays = ['mon', 'tue', 'wed', 'thu', 'fri'];
let pending = null;

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
        detail = typeof body.detail === 'string' ? body.detail : JSON.stringify(body.detail);
      }
    } catch (_) { /* JSON でないレスポンス */ }
    throw new Error(detail);
  }
  return response;
}

/** 画面の入力を、サーバに送る指定にまとめる。 */
function currentSpec() {
  const mode = document.querySelector('input[name=mode]:checked').value;
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
  if (mode === 'months') spec.months = Number($('#months').value) || 12;
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
    render(model);
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

// ---------------------------------------------------------------- プレビュー
function render(model) {
  const { spec, timeline, holidays } = model;
  const chartW = CHART_WIDTH[timeline.unit] || 42;
  const withWeekday = timeline.unit === 'day';

  const cols = el('colgroup');
  for (const column of TABLE_COLUMNS) cols.append(el('col', { style: `width:${column.width}px` }));
  for (let i = 0; i < timeline.columns.length; i++) {
    cols.append(el('col', { style: `width:${chartW}px` }));
  }

  const head = el('thead');

  // 1 段目: 予定 / 実績 のグループ見出し + 月 (日程表の上段)
  const bandRow = el('tr');
  let i = 0;
  while (i < TABLE_COLUMNS.length) {
    const group = TABLE_COLUMNS[i].group;
    let span = 1;
    while (i + span < TABLE_COLUMNS.length && TABLE_COLUMNS[i + span].group === group) span++;
    bandRow.append(el('th', { colspan: span, text: group }));
    i += span;
  }
  // 月は結合せず、区切りの列にだけ置く (Excel と同じ)
  const bandAt = new Map(timeline.bands.map((b) => [b.start, b.label]));
  timeline.columns.forEach((_column, index) => {
    bandRow.append(el('th', { class: 'month', text: bandAt.get(index) || '' }));
  });
  head.append(bandRow);

  // 2 段目: 表側の見出し + 週の開始日 (日程表の下段)
  const headRow = el('tr');
  for (const column of TABLE_COLUMNS) headRow.append(el('th', { text: column.label }));
  for (const column of timeline.columns) {
    headRow.append(el('th', {
      class: `axis${column.rest ? ' rest' : ''}${column.saturday ? ' saturday' : ''}`,
      text: column.label,
      title: column.start,
    }));
  }
  head.append(headRow);

  // 3 段目: 曜日 (日単位のときだけ)
  if (withWeekday) {
    const dayRow = el('tr');
    for (const _column of TABLE_COLUMNS) dayRow.append(el('th', {}));
    for (const column of timeline.columns) {
      dayRow.append(el('th', {
        class: `axis${column.rest ? ' rest' : ''}${column.saturday ? ' saturday' : ''}`,
        text: column.weekday,
      }));
    }
    head.append(dayRow);
  }

  // 記入用の空行
  const body = el('tbody');
  const shown = Math.min(spec.rows, PREVIEW_ROW_LIMIT);
  for (let r = 0; r < shown; r++) {
    const tr = el('tr');
    for (const column of TABLE_COLUMNS) {
      tr.append(el('td', { class: column.plan ? 'plan' : '' }));
    }
    for (const column of timeline.columns) {
      tr.append(el('td', {
        class: `chart${column.rest ? ' rest' : ''}${column.saturday ? ' saturday' : ''}`,
      }));
    }
    body.append(tr);
  }

  $('#sheet').replaceChildren(
    el('div', { class: 'sheet-title', text: spec.title }),
    el('table', { class: 'wbs' }, cols, head, body),
  );

  const omitted = spec.rows > shown ? `（先頭 ${shown} 行を表示）` : '';
  $('#preview-info').textContent =
    `${spec.start} 〜 ${spec.end} / ${timeline.columns.length} 列`
    + ` / 空行 ${spec.rows} 行${omitted}`;
  // 数えるのは祝日と休業日だけ。週末は稼働曜日の設定で決まる。
  $('#holiday-note').textContent = holidays.length
    ? `この期間の祝日・休業日は ${holidays.length} 日です。`
    : 'この期間に祝日・休業日はありません。';
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
}

async function download() {
  const button = $('#btn-build');
  button.disabled = true;
  try {
    const response = await api('/api/build', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(currentSpec()),
    });
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
