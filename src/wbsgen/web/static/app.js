/* wbsgen — 空の WBS を作る / 記入済み Excel をガントチャートで見る
 *
 * 画面が持つのは「用紙の指定」と描画だけ。日程表の組み立ても Excel の読み取りも
 * サーバ側で行い、返ってきたモデルをそのまま表と SVG に起こす。
 */
'use strict';

// Excel の列幅に合わせたプレビューの列幅 (px)
// Excel の列幅に合わせたプレビューの列幅 (px)。見出しは言語ごとに引く。
const TABLE_COLUMNS = [
  { key: 'group', width: 74, cls: 'group' },
  { key: 'subgroup', width: 74, cls: 'subgroup' },
  { key: 'no', width: 38 },
  { key: 'name', width: 168, cls: 'name' },
  { key: 'start', width: 46, group: 'plan', plan: true },
  { key: 'days', width: 38, group: 'plan', plan: true },
  { key: 'end', width: 46, group: 'plan', plan: true },
  { key: 'actual_start', width: 42, group: 'actual' },
  { key: 'actual_days', width: 38, group: 'actual' },
  { key: 'actual_end', width: 42, group: 'actual' },
  { key: 'delay', width: 38, group: 'actual' },
  { key: 'progress', width: 38, group: 'actual' },
  { key: 'effort', width: 38 },
  { key: 'predecessor', width: 38 },
  { key: 'member', width: 56, cls: 'member' },
  { key: 'status', width: 62, cls: 'status' },
];

/*
 * 狭い画面で出す列。スマートフォンの幅では表と日程表を全部は置けないので、
 * 出す列を選べるようにする (値は列幅 px)。
 *
 *   min  項目だけ    — 日程表を広く使う (狭い画面の既定)
 *   key  主要な列    — 日付・進捗・状態まで
 *   all  すべての列  — Excel と同じ
 */
const COLUMN_SETS = {
  min: { name: 128 },
  key: {
    name: 118, start: 42, end: 42, actual_start: 42, actual_end: 42,
    progress: 38, status: 58,
  },
};
const COLUMN_MODES = ['min', 'key', 'all'];

//: この幅より狭いと、表を全部出すと日程表が見えなくなる (列を選べるようにする)
const NARROW = window.matchMedia('(max-width: 900px)');
//: この幅より狭いと、指定フォームと日程表を横に並べられない (タブで切り替える)
const PHONE = window.matchMedia('(max-width: 760px)');

// 画面の文言。既定は日本語。
const UI = {
  ja: {
    app_title: 'WBS ジェネレータ',
    tagline_blank: '期間を指定すると、日程表と記入用の空行だけの Excel を書き出します。',
    tagline_chart: '記入済みの WBS をガントチャートで表示しています。',
    import: 'Excel を読み込む',
    back: '新規作成に戻る',
    export: 'ガントチャート付きで書き出す',
    download: 'Excel をダウンロード',
    section_period: '期間',
    section_sheet: '用紙',
    section_calendar: '稼働日',
    section_members: '担当者一覧に載せる名前',
    field_title: 'プロジェクト名',
    field_start: '開始日',
    field_end: '終了日',
    field_months: '月数',
    field_unit: '日程表の単位',
    field_rows: '記入用の空行',
    field_jp_holidays: '日本の祝日を休みにする',
    field_holidays: '休業日を追加',
    field_members: '担当者',
    members_note: '1 行に 1 人。空でもかまいません。',
    members_placeholder: '設計\n製造\nテスト',
    mode_end: '終了日で指定',
    mode_months: '月数で指定',
    preview: 'プレビュー',
    chart: 'ガントチャート',
    chart_unit: '表示単位',
    columns_shown: '列',
    columns_min: '項目のみ',
    columns_key: '主要な列',
    columns_all: 'すべて',
    pane_form: '指定',
    pane_preview: 'プレビュー',
    hint_blank: '実際の Excel と同じ列・同じ見出しです',
    hint_chart: '記入済みの Excel を読み込んで表示しています',
    columns: {
      group: '大項目', subgroup: '中項目', no: '項番', name: '項目',
      start: '開始日', days: '日数', end: '終了日',
      actual_start: '開始', actual_days: '日数', actual_end: '終了',
      delay: '遅れ', progress: '進捗',
      effort: '工数', predecessor: '先行', member: '担当', status: '状態',
    },
    group_plan: '予定',
    group_actual: '実績',
    totals: {
      rows: '行数', done: '完了', running: '進行中', delayed: '遅延',
      progress: '全体進捗', effort: '工数', period: '期間',
    },
    unit_effort: '人日',
    unit_days: '日',
    holidays_some: 'この期間の祝日・休業日は {n} 日です。',
    holidays_none: 'この期間に祝日・休業日はありません。',
    info_blank: '{start} 〜 {end} / {columns} 列 / 空行 {rows} 行',
    info_chart: '{start} 〜 {end} / {columns} 列 / {rows} 行',
    info_limited: '（先頭 {n} 行を表示）',
    warning_title: '読めなかった項目があります（その項目だけ空にしました）',
    derived_days: '開始日と終了日から数えた稼働日数',
    derived_end: '日数から補った終了日',
    derived_actual_days: '実績の開始日と終了日から数えた稼働日数',
    derived_actual_end: '実績日数から補った終了日',
    derived_progress: '実績の終了日が入っているので 100%',
    imported: '{name} を読み込みました。',
    busy_import: '{name} を読み込んでいます…',
    busy_build: 'Excel を作っています…',
    busy_export: 'ガントチャート付きの Excel を作っています…',
    saved: '{name} を書き出しました。',
    cannot_import: '読み込めません: {reason}',
    cannot_save: '書き出せません: {reason}',
    cannot_start: '起動できません: {reason}',
    fix_input: '指定を直してください。',
    tip_plan: '予定 {start} 〜 {end}',
    tip_actual: '実績 {start} 〜 {end}',
    tip_running: '進行中',
    tip_progress: '進捗 {value}%',
    tip_member: '担当 {value}',
    tip_status: '状態 {value}',
  },
  en: {
    app_title: 'WBS Generator',
    tagline_blank: 'Choose a period and download an Excel file with the calendar and blank rows.',
    tagline_chart: 'Showing a filled-in WBS as a Gantt chart.',
    import: 'Import Excel',
    back: 'Back to new sheet',
    export: 'Export with Gantt chart',
    download: 'Download Excel',
    section_period: 'Period',
    section_sheet: 'Sheet',
    section_calendar: 'Working days',
    section_members: 'Owners for the members sheet',
    field_title: 'Project name',
    field_start: 'Start date',
    field_end: 'End date',
    field_months: 'Months',
    field_unit: 'Calendar unit',
    field_rows: 'Blank rows',
    field_jp_holidays: 'Treat Japanese public holidays as days off',
    field_holidays: 'Extra days off',
    field_members: 'Owners',
    members_note: 'One name per line. May be left empty.',
    members_placeholder: 'Design\nBuild\nTest',
    mode_end: 'By end date',
    mode_months: 'By months',
    preview: 'Preview',
    chart: 'Gantt chart',
    chart_unit: 'Unit',
    columns_shown: 'Columns',
    columns_min: 'Name only',
    columns_key: 'Key columns',
    columns_all: 'All',
    pane_form: 'Settings',
    pane_preview: 'Preview',
    hint_blank: 'The same columns and headers as the Excel file',
    hint_chart: 'Imported from a filled-in Excel file',
    columns: {
      group: 'Group', subgroup: 'Sub-group', no: 'No.', name: 'Task',
      start: 'Start', days: 'Days', end: 'End',
      actual_start: 'Start', actual_days: 'Days', actual_end: 'End',
      delay: 'Delay', progress: 'Progress',
      effort: 'Effort', predecessor: 'Pred.', member: 'Owner', status: 'Status',
    },
    group_plan: 'Planned',
    group_actual: 'Actual',
    totals: {
      rows: 'Rows', done: 'Done', running: 'In progress', delayed: 'Delayed',
      progress: 'Overall', effort: 'Effort', period: 'Period',
    },
    unit_effort: 'person-days',
    unit_days: 'd',
    holidays_some: '{n} public holidays / days off in this period.',
    holidays_none: 'No public holidays or days off in this period.',
    info_blank: '{start} - {end} / {columns} columns / {rows} blank rows',
    info_chart: '{start} - {end} / {columns} columns / {rows} rows',
    info_limited: ' (showing the first {n})',
    warning_title: 'Some cells could not be read (only those were left empty)',
    derived_days: 'Working days counted from the start and end dates',
    derived_end: 'End date derived from the number of days',
    derived_actual_days: 'Working days counted from the actual start and end dates',
    derived_actual_end: 'End date derived from the actual number of days',
    derived_progress: '100% because an actual end date is filled in',
    imported: 'Imported {name}.',
    busy_import: 'Reading {name}…',
    busy_build: 'Building the Excel file…',
    busy_export: 'Building the Excel file with the Gantt chart…',
    saved: 'Saved {name}.',
    cannot_import: 'Cannot import: {reason}',
    cannot_save: 'Cannot save: {reason}',
    cannot_start: 'Cannot start: {reason}',
    fix_input: 'Please correct the settings.',
    tip_plan: 'Planned {start} - {end}',
    tip_actual: 'Actual {start} - {end}',
    tip_running: 'in progress',
    tip_progress: 'Progress {value}%',
    tip_member: 'Owner {value}',
    tip_status: 'Status {value}',
  },
};

const LANGUAGE_KEY = 'wbsgen.language';
const DEFAULT_LANGUAGE = 'ja';

let language = DEFAULT_LANGUAGE;

/** 現在の言語の文言。``{key}`` を values で置き換える。 */
function t(key, values = {}) {
  const table = UI[language] || UI[DEFAULT_LANGUAGE];
  const raw = table[key];
  if (typeof raw !== 'string') return raw;
  return raw.replace(/\{(\w+)\}/g, (_m, name) => (
    values[name] === undefined ? `{${name}}` : values[name]));
}

const CHART_WIDTH = { day: 22, week: 42, month: 58 };

//: 読み込んだ WBS を最初に見せるときの表示単位。日ごとの動きが判るよう
//: 日単位で開く (ファイルの「設定」より優先。上の選択で切り替えられる)。
const DEFAULT_CHART_UNIT = 'day';
const ROW_H = 24;
const HEAD_H = 24;

//: 空行のプレビューはこれ以上描いても読めないので省略する
const PREVIEW_ROW_LIMIT = 60;

//: 保存用の URL を残しておく時間 (これより早く捨てると保存できない端末がある)
const SAVE_KEEP_MS = 60000;

let meta = null;
let workdays = ['mon', 'tue', 'wed', 'thu', 'fri'];
let mode = 'blank';        // 'blank' = 新規作成 / 'chart' = 読み込んだ WBS
let pending = null;
let loaded = null;         // 読み込んだファイルの名前と中身 ({ name, bytes })
let importToken = 0;       // 読み込みの通し番号 (古い応答を捨てるため)
let pane = 'form';         // 狭い画面でどちらを見せているか ('form' / 'preview')
let columnsMode = 'min';   // 狭い画面で出す列 ('min' / 'key' / 'all')
let lastView = null;       // 直前に描いた内容 (画面幅が変わったら描き直す)
let busyCount = 0;         // 進行中の読み込み・書き出しの数
let suggestedTitle = '';   // 提案したプロジェクト名 (書き換えられたか判る)

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

// ------------------------------------------------------------ 処理中の表示
/**
 * 読み込み・書き出しの間、何をしているかを出して二重の操作を止める。
 * 必ず :func:`endBusy` と対にして呼ぶこと。
 */
function startBusy(key, params) {
  busyCount += 1;
  $('#busy-text').textContent = t(key, params);
  $('#busy').hidden = false;
  applyBusy();
}

function endBusy() {
  busyCount = Math.max(0, busyCount - 1);
  if (busyCount === 0) $('#busy').hidden = true;
  applyBusy();
}

function applyBusy() {
  document.body.classList.toggle('is-busy', busyCount > 0);
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
    language,
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
    const response = await api(`/api/preview?lang=${language}`, {
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
      info: t('info_blank', {
        start: model.spec.start, end: model.spec.end,
        columns: model.timeline.columns.length, rows: model.spec.rows,
      }),
    });
    $('#holiday-note').textContent = model.holidays.length
      ? t('holidays_some', { n: model.holidays.length })
      : t('holidays_none');
    $('#btn-build').disabled = false;
  } catch (error) {
    if (error.name === 'AbortError') return;
    banner(error.message);
    $('#preview-info').textContent = '';
    $('#sheet').replaceChildren(el('p', { class: 'empty-note', text: t('fix_input') }));
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
  lastView = view;
  const { timeline } = view;
  const colW = CHART_WIDTH[timeline.unit] || 42;
  const withWeekday = timeline.unit === 'day';
  const headH = HEAD_H * (withWeekday ? 3 : 2);
  const rowCount = view.rows.length
    || Math.min(view.blankRows || 0, PREVIEW_ROW_LIMIT);
  const bodyH = Math.max(rowCount, 1) * ROW_H;
  const width = timeline.columns.length * colW;

  // 表が広いまま貼り付くと日程表に届かないので、そのときは一緒に流す
  const grid = el('div', { class: `grid${isCompact() ? '' : ' scroll-table'}` },
    el('div', { class: 'table-wrap' }, buildTable(view, rowCount, withWeekday)),
    el('div', { class: 'chart-wrap' },
      chartHead(timeline, colW, width, headH, withWeekday),
      chartBody(view, colW, width, bodyH, rowCount)),
  );

  const parts = [el('div', { class: 'sheet-title', text: view.title })];
  if (view.warnings && view.warnings.length) parts.push(warningBox(view.warnings));
  parts.push(grid);
  $('#sheet').replaceChildren(...parts);

  const limited = view.blankRows > rowCount ? t('info_limited', { n: rowCount }) : '';
  $('#preview-info').textContent = view.info + limited;
  renderTotals(view.totals);
}

/** 画面幅や列の指定が変わったときに、同じ内容を描き直す。 */
function redraw() {
  if (lastView) render(lastView);
}

/** 表を全部出すと日程表が見えなくなる幅か。 */
function isNarrow() {
  return NARROW.matches;
}

/** 指定フォームと日程表を横に並べられない幅か。 */
function isPhone() {
  return PHONE.matches;
}

/** いま使う列の組。広い画面や「すべて」を選んでいるときは ``null``。 */
function columnSet() {
  return isNarrow() ? COLUMN_SETS[columnsMode] || null : null;
}

/** 表を貼り付けたままにできるほど列が少ないか。 */
function isCompact() {
  return isNarrow() && columnsMode === 'min';
}

/** いま描く列。狭い画面では選んだ組に絞り、幅も詰める。 */
function tableColumns() {
  const set = columnSet();
  if (!set) return TABLE_COLUMNS;
  return TABLE_COLUMNS
    .filter((c) => c.key in set)
    .map((c) => ({ ...c, width: set[c.key] }));
}

function warningBox(warnings) {
  return el('div', { class: 'warnings' },
    el('strong', { text: t('warning_title') }),
    el('ul', {}, ...warnings.slice(0, 5).map((w) => el('li', { text: w }))));
}

// ---------------------------------------------------------------- 表
function buildTable(view, rowCount, withWeekday) {
  const columns = tableColumns();
  const cols = el('colgroup');
  for (const column of columns) cols.append(el('col', { style: `width:${column.width}px` }));
  // 表そのものの幅も決めておく。決めないと、長い担当名などに引きずられて
  // 列が広がり、その分だけ日程表が右へ押し出されてしまう。
  const width = columns.reduce((total, column) => total + column.width, 0);

  const head = el('thead');
  const bandRow = el('tr');
  let i = 0;
  while (i < columns.length) {
    const group = columns[i].group || '';
    let span = 1;
    while (i + span < columns.length
           && (columns[i + span].group || '') === group) span++;
    const label = group ? t(group === 'plan' ? 'group_plan' : 'group_actual') : '';
    bandRow.append(el('th', { colspan: span, text: label }));
    i += span;
  }
  head.append(bandRow);
  head.append(el('tr', {}, ...columns.map(
    (c) => el('th', { text: t('columns')[c.key] }))));
  if (withWeekday) {
    head.append(el('tr', {}, ...columns.map(() => el('th', {}))));
  }

  const body = el('tbody');
  for (let r = 0; r < rowCount; r++) {
    const row = view.rows[r] || null;
    const tr = el('tr');
    for (const column of columns) tr.append(tableCell(row, column, view.rows[r - 1]));
    body.append(tr);
  }
  return el('table', { class: 'wbs', style: `width:${width}px` }, cols, head, body);
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
    value = value == null ? '' : `${value} ${t('unit_days')}`;
  } else if (/start|end$/.test(column.key)) {
    value = shortDate(value);
  } else if (value == null) value = '';

  td.textContent = value;
  // 記入内容から導き出した値は薄く見せ、理由を添える
  if ((row.derived || []).includes(column.key)) {
    td.classList.add('derived');
    td.title = t(`derived_${column.key}`);
  }
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
  if (row.start) lines.push(t('tip_plan', { start: row.start, end: row.end || '-' }));
  if (row.actual_start) {
    lines.push(t('tip_actual', {
      start: row.actual_start, end: row.actual_end || t('tip_running'),
    }));
  }
  if (row.progress != null) {
    lines.push(t('tip_progress', { value: Math.round(row.progress * 100) }));
  }
  if (row.member) lines.push(t('tip_member', { value: row.member }));
  if (row.status) lines.push(t('tip_status', { value: row.status }));
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
  const names = t('totals');
  const items = [
    [names.rows, `${totals.rows}`],
    [names.done, `${totals.done}`],
    [names.running, `${totals.running}`],
    [names.delayed, `${totals.delayed}`, totals.delayed > 0],
    [names.progress, `${Math.round(totals.progress * 100)}%`],
    [names.effort, `${totals.effort} ${t('unit_effort')}`],
    [names.period, `${totals.first_day ?? '-'} - ${totals.last_day ?? '-'}`],
  ];
  box.replaceChildren(...items.map(([term, value, warn]) => el('div', {},
    el('dt', { text: term }),
    el('dd', { text: value, class: warn ? 'warn' : '' }))));
  box.hidden = false;
}

// ---------------------------------------------------------------- 読み込み
async function importFile(file) {
  let bytes;
  try {
    // 端末によっては、選んだファイルの参照が後で使えなくなる。
    // 中身をここで読み切って持っておき、書き出しにもこれを使う。
    bytes = await file.arrayBuffer();
  } catch (error) {
    banner(t('cannot_import', { reason: error.message }));
    return;
  }
  return importBytes(file.name, bytes, DEFAULT_CHART_UNIT, true);
}

/**
 * 持っている中身を送って読み込む (表示単位を変えたときは読み直す)。
 * ``announce`` は「読み込みました」を出すかどうか。
 */
async function importBytes(name, bytes, unit = null, announce = false) {
  const body = new FormData();
  body.append('file', new Blob([bytes]), name);
  const query = unit ? `&unit=${encodeURIComponent(unit)}` : '';
  const token = ++importToken;
  startBusy('busy_import', { name });
  try {
    const response = await api(`/api/import?lang=${language}${query}`,
                               { method: 'POST', body });
    const model = await response.json();
    // 待っている間に「戻る」や別の読み込みが起きていたら、この結果は捨てる
    if (token !== importToken) return;
    loaded = { name, bytes };
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
      info: t('info_chart', {
        start: model.timeline.start, end: model.timeline.end,
        columns: model.timeline.columns.length, rows: model.rows.length,
      }),
    });
    if (announce) banner(t('imported', { name }), true);
  } catch (error) {
    if (token !== importToken) return;
    banner(t('cannot_import', { reason: error.message }));
  } finally {
    endBusy();
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
  $('#preview-title').textContent = t(chart ? 'chart' : 'preview');
  $('#preview-hint').textContent = t(chart ? 'hint_chart' : 'hint_blank');
  $('#tagline').textContent = t(chart ? 'tagline_chart' : 'tagline_blank');
  applyNarrow();
}

// ------------------------------------------------------ スマートフォン向け
/**
 * 狭い画面では、指定フォームとプレビューを切り替えて 1 つずつ見せる。
 * 読み込んだ WBS を見ている間はフォームが無いので、切り替えも要らない。
 */
function applyNarrow() {
  const narrow = isNarrow();
  const tabs = isPhone() && mode === 'blank';
  $('#pane-tabs').hidden = !tabs;
  $('#columns-field').hidden = !narrow;

  const layout = document.querySelector('.layout');
  layout.classList.toggle('pane-form', tabs && pane === 'form');
  layout.classList.toggle('pane-preview', tabs && pane === 'preview');
  for (const button of document.querySelectorAll('.pane-tab')) {
    button.setAttribute('aria-pressed', String(button.dataset.pane === pane));
  }
}

function showPane(next) {
  pane = next;
  applyNarrow();
}

// ---------------------------------------------------------------- 言語
/** 保存してある言語を読む。読めない環境 (プライベートウィンドウ等) は既定。 */
function storedLanguage() {
  try {
    const saved = localStorage.getItem(LANGUAGE_KEY);
    if (saved && UI[saved]) return saved;
  } catch (_) { /* 保存領域が使えない */ }
  return DEFAULT_LANGUAGE;
}

function rememberLanguage(value) {
  try {
    localStorage.setItem(LANGUAGE_KEY, value);
  } catch (_) { /* 保存できなくても動作に影響はない */ }
}

/** ``data-i18n`` の付いた要素に、現在の言語の文言を入れる。 */
function applyLanguage() {
  document.documentElement.lang = language;
  document.title = t('app_title');
  for (const node of document.querySelectorAll('[data-i18n]')) {
    node.textContent = t(node.dataset.i18n);
  }
  $('#members').placeholder = t('members_placeholder');
  $('#preview-title').textContent = t(mode === 'chart' ? 'chart' : 'preview');
  $('#preview-hint').textContent = t(mode === 'chart' ? 'hint_chart' : 'hint_blank');
  $('#tagline').textContent = t(mode === 'chart' ? 'tagline_chart' : 'tagline_blank');
}

/** 言語を切り替え、選択肢と表示を作り直す。 */
async function switchLanguage(value) {
  language = UI[value] ? value : DEFAULT_LANGUAGE;
  rememberLanguage(language);
  applyLanguage();
  try {
    meta = await (await api(`/api/meta?lang=${language}`)).json();
  } catch (error) {
    banner(t('cannot_start', { reason: error.message }));
    return;
  }
  fillChoices();
  // 提案のままなら新しい言語の提案に差し替える (書き換えてあれば触らない)
  if ($('#title').value === suggestedTitle) $('#title').value = meta.suggested.title;
  suggestedTitle = meta.suggested.title;
  $('#title').placeholder = suggestedTitle;

  if (mode === 'chart' && loaded) {
    await importBytes(loaded.name, loaded.bytes, $('#chart-unit').value);
  }
  else await refresh();
}

/** 言語で変わる選択肢 (表示単位・曜日) を入れ直す。 */
function fillChoices() {
  const unit = $('#unit').value || 'week';
  const chartUnit = $('#chart-unit').value || DEFAULT_CHART_UNIT;
  const options = () => meta.units.map(
    (u) => el('option', { value: u.value, text: u.label }));
  $('#unit').replaceChildren(...options());
  $('#unit').value = unit;
  $('#chart-unit').replaceChildren(...options());
  $('#chart-unit').value = chartUnit;
  $('#columns-mode').replaceChildren(...COLUMN_MODES.map(
    (m) => el('option', { value: m, text: t(`columns_${m}`) })));
  $('#columns-mode').value = columnsMode;
  renderWeekdays();
}

// ---------------------------------------------------------------- 入力
function buildForm() {
  $('#unit').value = 'week';
  $('#chart-unit').value = DEFAULT_CHART_UNIT;
  $('#rows').value = meta.default_rows;
  $('#rows').max = meta.max_rows;
  $('#start').value = meta.suggested.start;
  $('#end').value = meta.suggested.end;
  suggestedTitle = meta.suggested.title;
  $('#title').value = suggestedTitle;
  $('#title').placeholder = suggestedTitle;
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
  $('#language').addEventListener('change',
                                  (event) => switchLanguage(event.target.value));
  $('#spec-form').addEventListener('submit', (event) => event.preventDefault());
  $('#btn-build').addEventListener('click', download);
  $('#btn-export').addEventListener('click', exportChart);
  $('#btn-back').addEventListener('click', () => {
    importToken += 1;   // 読み込み中なら、その結果は捨てる
    loaded = null;
    setMode('blank');
    banner('');
    refresh();
  });
  $('#chart-unit').addEventListener('change', () => {
    if (loaded) importBytes(loaded.name, loaded.bytes, $('#chart-unit').value);
  });
  $('#import-file').addEventListener('change', (event) => {
    const file = event.target.files[0];
    if (file) importFile(file);
    event.target.value = '';
  });

  for (const button of document.querySelectorAll('.pane-tab')) {
    button.addEventListener('click', () => showPane(button.dataset.pane));
  }
  $('#columns-mode').addEventListener('change', (event) => {
    columnsMode = event.target.value;
    redraw();
  });
  // 画面を回したり幅が変わったら、その幅に合わせて描き直す
  for (const query of [NARROW, PHONE]) {
    query.addEventListener('change', () => { applyNarrow(); redraw(); });
  }
}

/** 空の WBS を書き出す。 */
function download() {
  return save($('#btn-build'), 'busy_build', `/api/build?lang=${language}`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(currentSpec()),
  });
}

/** 表示中のガントチャートを図形として描き込んだ Excel を書き出す。 */
function exportChart() {
  if (!loaded) return Promise.resolve();
  const body = new FormData();
  body.append('file', new Blob([loaded.bytes]), loaded.name);
  const unit = $('#chart-unit').value;
  return save($('#btn-export'), 'busy_export',
              `/api/export?lang=${language}&unit=${encodeURIComponent(unit)}`,
              { method: 'POST', body });
}

/**
 * 受け取った中身をファイルとして保存させる。
 *
 * すぐに URL を捨てると、端末によっては保存が始まる前に消えてしまうので、
 * しばらく残してから片付ける。
 */
function saveBlob(blob, name) {
  const url = URL.createObjectURL(blob);
  const link = el('a', { href: url, download: name, style: 'display:none' });
  document.body.append(link);
  link.click();
  setTimeout(() => { link.remove(); URL.revokeObjectURL(url); }, SAVE_KEEP_MS);
}

async function save(button, busyKey, path, options) {
  button.disabled = true;
  startBusy(busyKey);
  try {
    const response = await api(path, options);
    const blob = await response.blob();
    const disposition = response.headers.get('Content-Disposition') || '';
    const match = /filename\*=UTF-8''([^;]+)/.exec(disposition);
    const name = match ? decodeURIComponent(match[1]) : 'wbs.xlsx';
    saveBlob(blob, name);
    banner(t('saved', { name }), true);
  } catch (error) {
    banner(t('cannot_save', { reason: error.message }));
  } finally {
    button.disabled = false;
    endBusy();
  }
}

async function start() {
  bind();
  language = storedLanguage();
  applyLanguage();
  try {
    meta = await (await api(`/api/meta?lang=${language}`)).json();
  } catch (error) {
    banner(t('cannot_start', { reason: error.message }));
    return;
  }
  $('#language').replaceChildren(...meta.languages.map(
    (l) => el('option', { value: l.value, text: l.label })));
  $('#language').value = language;
  fillChoices();
  buildForm();
  applyNarrow();
  await refresh();
}

start();
