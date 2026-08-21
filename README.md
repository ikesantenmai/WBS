# wbsgen — WBS / ガントチャート Excel 生成アプリケーション

テキストで書いたプロジェクト定義 (YAML / JSON / CSV) から、
**WBS 一覧 + ガントチャート**の Excel ブックを生成します。

添付の WBS ファイル (`DS_WEB…PJ_WBS….xls`) および付属の VBA アドイン
`km_module.xla` の仕様を解析し、同じシート構成・同じ日程計算規則を
Python で再実装したものです。Excel やマクロを使わずに、CI やスクリプトから
同じ成果物を作り直せます。

ブラウザで編集する **Web アプリ**と、定義ファイルから一括生成する **CLI** の
どちらからでも同じ Excel を作れます。

```
wbsgen serve                 # ブラウザで編集する (http://127.0.0.1:8000)
wbsgen init -o wbs.yaml      # 雛形を書き出す
wbsgen build wbs.yaml        # wbs.xlsx を生成する
wbsgen check wbs.yaml        # 遅延タスクを一覧する
```

---

## 生成されるもの

`スケジュール` シートは元ファイルと同じレイアウトです。

| 列 | 内容 | | 列 | 内容 |
|----|------|-|----|------|
| B | 大項目 | | K | 実績 終了日 |
| C | 中項目 | | L | 遅れ (稼働日) |
| D | 項番 | | M | 進捗 |
| E | 項目 | | N | 工数 (人日) |
| F | 予定 開始日 | | O | 先行 |
| G | 予定 日数 (稼働日) | | Q | 担当 |
| H | 予定 終了日 | | R | 状態 |
| I | 実績 開始日 | | S〜 | ガントチャート |
| J | 実績 日数 | | | |

ブックには 4 シートが入ります。

- **スケジュール** — WBS 一覧とガントチャート (`S4` で枠固定)
- **担当者一覧** — 担当者と担当色、個別の休日/出勤設定
- **標準工程** — テンプレートとして再利用する工程定義
- **設定** — 表示設定・状態しきい値・稼働日設定・休日一覧の書き出し

### ガントチャートの描画

バーは**セルの塗りつぶしではなく浮動図形 (DrawingML)** として描かれます。
元ツールが VBA でオートシェイプを描いているのと同じ方式で、週表示・月表示でも
バーの端が日付どおりの位置に落ちます。

- **予定バー** — 担当色のグラデーション矩形
- **実績バー** — 予定バーの下段。未完了は半透明
- **工程バー** (`kind: summary`) — 期間つきの山形 (シェブロン)
- **マイルストーン** (`kind: milestone`) — 菱形・三角・星などの記号
- **先行タスク線** — 先行 → 後続を結ぶ L 字の矢印
- **現在日線** — 基準日の赤い破線
- **イナズマ線** — 各タスクの進捗到達日を結ぶ折れ線

---

## インストール

```bash
pip install -e .          # CLI のみ (openpyxl と PyYAML だけ)
pip install -e '.[web]'   # Web アプリも使う (FastAPI / uvicorn を追加)
```

Python 3.9 以上。

---

## Web アプリ

```bash
wbsgen serve                       # http://127.0.0.1:8000
wbsgen serve --host 0.0.0.0 --port 8080
```

画面は左に設定、中央に WBS 一覧、右にガントチャートが並びます。

- **表を編集** — 行をクリックするとダイアログが開き、日付・日数・進捗・担当などを
  編集できます。ドラッグで行の並べ替え、ボタンでタスク／工程／マイルストーンの追加。
- **その場で再計算** — 編集するたびにサーバへ送られ、終了日・遅れ・状態・
  イナズマ線が引き直されます。日程計算もチャートの座標も **CLI と同じエンジン**を
  通るので、**画面のプレビューと生成される Excel は一致します**。
- **設定** — 表示単位 (日/週/月)、表示期間、基準日、状態しきい値、稼働曜日、
  祝日の扱い、担当者と担当色をその場で切り替えられます。
- **読み込み** — 既存の YAML / JSON / CSV をドロップして続きから編集できます
  (Excel が書き出した CP932 の CSV も読めます)。
- **書き出し** — 「Excel を生成」で `.xlsx`、「YAML 保存」で定義ファイル。
  YAML はそのまま CLI の入力になるので、画面で組んだものを CI に載せられます。
- **表を畳む** — 一覧を項番・項目・状態だけにして、チャートを広く使えます。

### API

画面を使わずに HTTP から叩くこともできます。

| メソッド | パス | 内容 |
|---------|------|------|
| `GET` | `/api/meta` | 選択肢 (表示単位・種別・記号・雛形) |
| `GET` | `/api/template/{name}` | 雛形をプロジェクト定義として返す |
| `POST` | `/api/preview` | 日程を解決し、表とチャートの描画モデルを返す |
| `POST` | `/api/build` | Excel ブックを返す |
| `POST` | `/api/export/yaml` | 定義を YAML テキストで返す |
| `POST` | `/api/import` | YAML / JSON / CSV を読み込んで定義に変換する |

リクエストボディは定義ファイルと同じ構造の JSON です。定義に不備があると
`422` と日本語のメッセージが返ります。

```bash
curl -X POST http://127.0.0.1:8000/api/build \
     -H 'Content-Type: application/json' \
     -d @project.json -o wbs.xlsx
```

---

## 定義ファイル

```yaml
project:
  title: システム開発 スケジュール
  updated_at: 2026-04-01 09:00     # 省略時は生成時刻

chart:
  start: 2026-04-01                # チャート表示開始日 (省略時はタスクから自動)
  period_days: 300                 # 表示期間 (暦日)
  unit: week                       # day / week / month
  base_date: 2026-08-07            # 現在日。状態と遅れの判定基準 (省略時は実行日)
  man_month_days: 20
  thresholds:
    exec_remain_days: 5            # 「残り n 日」に切り替わる残稼働日数
    start_near_days: 10            # 「あと n 日」に切り替わる着手前稼働日数
  show:
    results: true                  # 実績バー
    progress: false                # 予定バー内の進捗塗り
    manpower: true                 # 工数列
    status: true                   # 状態列
    start_date: true
    end_date: true
    predecessor: true
    predecessor_line: true
    now_line: true
    inazuma_line: true
    member_color: true

calendar:
  workdays: [mon, tue, wed, thu, fri]
  japanese_holidays: true          # 日本の祝日を自動で休日にする
  holidays: [2026-12-30, 2026-12-31]
  extra_workdays: [2026-05-02]     # 休日だが稼働する日

members:
  - { name: DS,     color: "#4472C4" }
  - { name: DS開発, color: "#70AD47" }
  # 担当者ごとに稼働日を変えることもできる
  - { name: 外注,   color: "#ED7D31", workdays: [mon, tue, wed, thu, fri, sat] }

tasks:
  - { group: 開発, "no": "10", name: 設計, kind: summary, member: DS }
  - { subgroup: 基本設計, "no": "11", name: 画面設計, start: 2026-04-01, days: 5,
      effort: 5, actual_start: 2026-04-01, progress: 100%, member: DS }
  - { "no": "12", name: 実装, predecessor: "11", days: 10, effort: 20, member: DS開発 }
  - { "no": "M1", name: リリース, kind: milestone, shape: diamond,
      predecessor: "12", days: 1, member: DS }
```

### タスクの属性

| 属性 | 説明 |
|------|------|
| `group` / `subgroup` | 大項目 / 中項目。省略すると直前の行を引き継ぐ |
| `no` (`id`) | 項番。`predecessor` から参照される |
| `name` | 項目名 (必須) |
| `kind` | `task` (既定) / `summary` (工程バー) / `milestone` (記号) |
| `start` / `days` / `end` | 予定。`days` は**稼働日数**。2 つ与えれば残り 1 つは自動算出 |
| `actual_start` / `actual_days` / `actual_end` | 実績 |
| `progress` | `0.8` `80%` `80` のいずれでも可 |
| `effort` | 工数 (人日) |
| `predecessor` | 先行タスクの項番。`"11,12"` で複数指定 |
| `member` | 担当者名。バーの色に反映される |
| `shape` | マイルストーンの記号: `diamond` `triangle` `circle` `star` `arrow` `chevron` |
| `comment` | 項目セルに付くコメント |

> **注意**: YAML 1.1 では素の `no:` は真偽値として解釈されます。
> `"no": "11"` と引用するか、別名の `id:` を使ってください
> (引用しない `no:` も項番として受け付けます)。

### 自動配置

`start` を書かずに `predecessor` と `days` だけ書いたタスクは、
**先行タスクの予定終了日の翌稼働日**から自動的に配置されます (FS 関係)。
起点となるタスクにだけ `start` を書けば、あとは連鎖して日程が決まります。

---

## CSV から作る

日本語ヘッダのタスク表 CSV をそのまま読み込めます
(`examples/tasks.csv` 参照)。

```bash
wbsgen build examples/tasks.csv -o out/wbs.xlsx
```

使える列名: `大項目` `中項目` `項番` `項目` `開始日` `日数` `終了日`
`実績開始` `実績日数` `実績終了` `進捗` `工数` `先行` `担当` `コメント` `記号` `種別`
(英語名 `group` `name` `start` … も可)

---

## 日程計算の規則

元ツールと同じ規則を実装しています。

- **予定終了日** = 予定開始日を 1 日目とした `日数` 稼働日目。
  開始日が休日なら翌稼働日から数え始めます。
- **実績日数** (未完了) = 実績開始日から基準日までの稼働日数。
  完了 (`progress: 100%`) すると実績終了日が入ります。
- **遅れ** = 予定終了日から基準日までの稼働日数 − 1。
  未着手なら予定開始日を起点にした「開始遅れ」になります。
- **状態** は次の順で判定します。

  | 条件 | 表示 |
  |------|------|
  | 実績終了日あり | `完了` |
  | 遅れあり | `遅れ n 日` |
  | 着手済み・残り稼働日 ≦ `exec_remain_days` | `残り n 日` |
  | 着手済み | `実行中` |
  | 未着手・開始まで ≦ `start_near_days` | `あと n 日` |
  | それ以外 | `-` |

### 元ファイルとの照合

添付 WBS の 128 行を `examples/sbi_web_wbs.yaml` に写して生成したところ、
**予定終了日・実績終了日・遅れ・状態のすべてが 126 行で完全に一致**しました。

```bash
wbsgen build examples/sbi_web_wbs.yaml -o out/sbi.xlsx
```

残る 2 行 (リハーサル) の差は、元ツールの祝日テーブルが 2007 年当時のままで
**2026-12-23 を天皇誕生日として休日扱いしている**ためです。天皇誕生日は
2020 年から 2/23 に移っているので、本アプリの計算 (12/25 終了) が現行法どおりです。
元の値を再現したい場合は `calendar.holidays` に `2026-12-23` を足してください。

---

## コマンド

```
wbsgen init  [-t standard|minimal] [-o wbs.yaml] [-f]
wbsgen build <定義ファイル> [-o out.xlsx] [--unit day|week|month] [--base-date YYYY-MM-DD]
wbsgen check <定義ファイル>
wbsgen serve [--host 127.0.0.1] [--port 8000] [--reload]
```

`build` の `--unit` / `--base-date` は定義ファイルの設定を一時的に上書きします。
週次報告のたびに `--base-date` を変えれば、同じ定義から最新の状態・遅れ・
イナズマ線を引き直せます。

---

## 開発

```bash
pip install -e '.[web]' pytest
pytest -q
```

`tests/test_web_ui.py` は Playwright と Chromium がある環境でだけ実行され、
無ければ skip します (JavaScript が実際に動くことの確認用)。

### 構成

```
src/wbsgen/
  workcal.py       稼働日カレンダーと日本の祝日
  model.py         データモデル
  scheduling.py    終了日・遅れ・状態・自動配置の導出
  style.py         配色と書式 (元ファイルから抽出)
  loader.py        YAML / JSON / CSV の読み込みと検証
  cli.py           コマンドライン
  serialize.py     定義の辞書化 / YAML 書き戻し
  preview.py       ブラウザ描画用のチャートモデル
  templates/       定義ファイルの雛形
  render/
    workbook.py    4 シートの組み立て
    timeline.py    チャート横軸 (日/週/月)
    gantt.py       バー・記号・線の配置計算
    drawing.py     DrawingML の生成
    inject.py      xlsx への図形パート注入
  web/
    app.py         FastAPI のエンドポイント
    static/        画面 (index.html / app.js / style.css)
```

Web UI は素の HTML + JavaScript で、ビルド手順も外部 CDN もありません。
チャートは SVG で描いていますが、座標は `preview.py` がサーバ側で計算した
ものをそのまま使うため、Excel 出力とずれません。
