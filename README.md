# ワークショップ スケジュール表示 Webアプリ(PoC)

ワークショップ運営のタスク・スケジュールを、ガントチャートで表示するWebアプリのPoCです。

**スケジュールの唯一の正(Single Source of Truth)は Excel ブック**(`data/workshop_schedule.xlsx`)です。
WSの追加・日付の決定・担当者の入力はすべて Excel で行い、Webアプリはそのファイルを読み取って
表示するだけです(書き換えない・DBは持たない)。

```
Excel ブック(人が編集・数式でタスク生成と日付計算)
   │  保存
   ▼
Webアプリ(Flask)── ファイルを読むだけ ──▶ ブラウザ(ガントチャート)
```

## 目次

1. [起動方法](#起動方法)
1. [GitHub Pages 版(docs/)](#github-pages-版docs)
2. [Excel ブックの構成](#excel-ブックの構成)
3. [運用の流れ](#運用の流れ)
4. [日付の計算ルール](#日付の計算ルール)
5. [ディレクトリ構成](#ディレクトリ構成)
6. [Webアプリの仕組み](#webアプリの仕組み)
7. [テスト](#テスト)
8. [注意点](#注意点)
9. [今後の拡張](#今後の拡張)

## 起動方法

Python 3.10 以上を想定しています。

```bash
cd workshop_scheduler
pip install -r requirements.txt
python run.py            # ブラウザで http://127.0.0.1:5000/ を開く
```

読み込むファイルは既定で `data/workshop_schedule.xlsx`。別の場所のファイル(OneDrive の同期フォルダなど)を
直接読ませる場合は環境変数で指定します。個人用 OneDrive でも、PC に同期されたファイルなら API なしで読めます。

```bash
SCHEDULE_EXCEL_PATH="$HOME/OneDrive/workshop_schedule.xlsx" python run.py
```

ファイルが保存し直されると(更新日時・サイズの変化を検知して)、次に画面を開いたときに自動で読み直します。

## GitHub Pages 版(docs/)

`docs/` は、OneDrive 上の Excel をブラウザが直接読む静的版です(サーバー不要。GitHub Pages にそのまま置ける)。
各メンバーが自分の個人用 Microsoft アカウントでサインインし、Microsoft Graph で共有リンクの Excel を取得して、
`docs/js/excel_source.js`(`app/services/excel_source.py` の JavaScript 版)で読みます。
Excel を読めるのは OneDrive でファイルを共有されたアカウントだけです。

```
ブラウザ ── 画面のファイル ──▶ GitHub Pages(docs/)
   │  Microsoft サインイン(MSAL.js)
   └── Graph API ──▶ 個人用 OneDrive の workshop_schedule.xlsx
```

| ファイル | 内容 |
|---|---|
| `docs/index.html` | 画面 |
| `docs/js/config.js` | クライアント ID・共有リンク・要求する権限(公開前に記入。どちらも秘密情報ではない) |
| `docs/js/onedrive.js` | サインインと Excel の取得 |
| `docs/js/excel_source.js` | Excel の読み取り(Python 版と同じ結果を返す) |
| `docs/js/schedule.js` | ガントチャートの描画(Flask 版と同じ。データの取り方だけ違う) |

手元で試す(Entra のアプリに `http://localhost:8000/` を「シングルページ アプリケーション」で登録しておく):

```bash
cd docs
python3 -m http.server 8000     # http://localhost:8000/ を開く(127.0.0.1 ではなく localhost)
```

`config.js` が空のときは、最初に画面でクライアント ID と共有リンクを入力します(そのブラウザにだけ保存)。

Python 版と JavaScript 版の読み取り結果が同じかは、次で確かめます(Node.js が必要)。
片方の読み取り処理を直したら、もう片方にも同じ修正を入れて確認してください。

```bash
cd tools/js_parity && npm install && cd ../..
python -m tools.js_parity.compare data/workshop_schedule.xlsx
```

## Excel ブックの構成

マクロは使っていません。タスクの生成と日付の計算は数式で行うため、ブラウザ版 Excel でもそのまま動きます。
見出しが青いシート・列が入力する場所、緑が自動計算です。

| シート | 入力/自動 | 内容 |
|---|---|---|
| 説明 | ― | 使い方とルール |
| タスク一覧 | 自動(数式) | WS × タスクマスタのタスクと、計算済みの開始日・終了日・担当者・警告。**Webアプリが読む** |
| WS | 入力 | WS名・種別(ワークショップ/その他)・並び順・開始日・ステータス |
| 日程入力 | 左4列は自動・右は入力 | WSを追加すると行が自動で並ぶ。開始日・終了日・ステータス・担当者1〜10・メモを入力 |
| タスクマスタ | 入力 | タスク名・標準作業日数・標準並び順・日付決定有無(Yes/No)・WS自動生成(対象/対象外) |
| メンバー | 入力 | 担当者のプルダウンと、画面での担当者の並び順 |
| 個別タスク | 入力 | マスタから作らない作業(社内MTGなど)。開始日・終了日(または作業日数)・ステータス・担当者1〜10。**Webアプリが読む** |

ステータスは「未着手 / 着手中 / 完了」(空欄は未着手)。タスク一覧・日程入力・個別タスクは見出しのフィルター(▼)で
WS名やステータスごとに絞り込めます。

ブックは次のコマンドで作り直せます(作成後は Excel で一度開いて保存すること。[注意点](#注意点)参照)。

```bash
python -m tools.make_schedule_workbook data/workshop_schedule.xlsx            # 空(タスクマスタのみ)
python -m tools.make_schedule_workbook data/sample_schedule.xlsx --sample   # サンプルデータ入り
```

ブックの形式を変えたときは、古いブックの入力値を新しい形式へ移せます(見出し名で対応させて書き写す)。

```bash
python -m tools.upgrade_workbook data/workshop_schedule.xlsx data/workshop_schedule_new.xlsx
```

## 運用の流れ

1. **WSを追加**:「WS」シートの一番下の空き行に WS名・種別「ワークショップ」・並び順・開始日を入れる。
   → 「タスク一覧」と「日程入力」に、タスクマスタ(WS自動生成=対象)のタスクが自動で並ぶ。
2. **担当者・ステータスを入れる**:「日程入力」の該当行の担当者1〜10・ステータスに入力する。
3. **日付を決める**:リハーサルなど日付決定有無=Yes のタスクの行に開始日を入れる。
   → 後続タスクの開始日・終了日が自動で入る。
4. **保存**すると、Webアプリの画面に反映される(ブラウザを再読み込み)。

## 日付の計算ルール

| 対象 | 開始日 |
|---|---|
| 並び順が最初のタスク | WS の開始日 |
| 日付決定有無=Yes のタスク | 日程入力に入れた日付(無ければ「未定」) |
| Yes のタスクの次の並び順 | その Yes タスクの終了日の翌日 |
| それ以外 | 直前の並び順の最も遅い終了日の翌日(直列) |
| 同じ並び順のタスク | 同じ期間(並列) |
| 日程入力に開始日を入れた No のタスク | その日付で固定(手動) |
| 日程入力に終了日を入れたタスク | 終了日はその日付(開始日だけ・終了日だけ・両方どれでも可) |

- 起点になるタスクに1件でも「未定」があれば、その後の自動のタスクも「未定」。
- 終了日は、日程入力に終了日があればその日付、無ければ 開始日 + 標準作業日数 − 1(暦日)。
- 終了日が開始日より前だと「警告」列に表示される。
- Yes のタスクの開始日が前のタスクの終了日以前だと「警告」列に表示され、画面上部にも出る。

## ディレクトリ構成

```
workshop_scheduler/
├── app/
│   ├── __init__.py              # アプリケーションファクトリ(create_app)
│   ├── services/
│   │   ├── excel_source.py      # Excel ブックの読み取り(キャッシュ付き・形式チェック)
│   │   └── schedule_service.py  # 画面用のデータに整形
│   ├── routes/
│   │   ├── views.py             # HTML画面("/")
│   │   └── api.py               # JSON API("/api/*")
│   ├── templates/schedule.html
│   └── static/
│       ├── css/style.css
│       └── js/schedule.js       # ガントチャートの描画
├── docs/                        # GitHub Pages 版(静的。上の「GitHub Pages 版」参照)
│   ├── index.html
│   ├── css/style.css
│   └── js/                      # config / onedrive / excel_source / schedule
├── tools/
│   ├── make_schedule_workbook.py  # Excel ブック(テンプレート)の生成
│   ├── upgrade_workbook.py        # 古い形式のブックから入力値を移す
│   └── js_parity/                 # Python 版と JS 版の読み取り結果の比較
├── data/
│   └── workshop_schedule.xlsx   # スケジュールの唯一の正(Git管理外)
├── tests/
│   ├── conftest.py              # テスト用ブックの作成
│   ├── test_excel_source.py     # 読み取りのテスト
│   └── test_api.py              # API・画面のテスト
├── config.py                    # 設定(読み込むファイルのパスなど)
├── run.py                       # 起動エントリーポイント
└── requirements.txt
```

## Webアプリの仕組み

- `excel_source.py` がブックを `openpyxl` で開き、**Excel が保存した数式の計算結果**(`data_only`)を読みます。
  読むのは「WS」「タスク一覧」「個別タスク」「メンバー」「タスクマスタ」。
- 読み込み結果はファイルの更新日時・サイズが変わるまでキャッシュします。
- API

| エンドポイント | 内容 |
|---|---|
| `GET /api/schedule` | `projects`(WS)・`tasks`(開始日が未定なら `start_date`/`end_date` が null)・`warnings`・`source`(ファイル名と更新日時) |
| `GET /api/members` | メンバー(メンバーシートに無い担当者名も含む) |
| `GET /api/projects` | WS 一覧 |
| `GET /api/task-masters` | タスクマスタ |

  ファイルが無い・形式が違う・計算結果が無い場合は、HTTP 503 と `{"error": "理由"}` を返し、画面にその理由を表示します。
- 画面
  - プロジェクト別/担当者別の表示切替。プロジェクト別は WS、担当者別は担当者で絞り込み(複数選択)。
  - バーの色はステータス(未着手=黄 / 着手中=青 / 完了=グレー)。日付決定有無=Yes のタスクは赤。バーの左端の線は WS の色。
  - 表示期間は「今週の1週前〜10週先」を基本に、それより前後に予定のあるタスクがあればその週まで広げる。
  - 未定タスクは WS 名の下に「未定 N件」、警告・メモ・ステータスはホバーで表示。
  - 読み込んだファイル名・更新日時と注意の一覧を画面上部に表示。

## テスト

```bash
python -m pytest tests/ -v
```

テストでは `tools/make_schedule_workbook.py` で実際にブックを作り、タスク一覧に計算結果に見立てた値を入れて、
読み取り(並び順・日付・未定・担当者・個別タスク・警告・ファイル更新時の読み直し・エラー)と API を検証します。

## 注意点

- **Webアプリは数式の計算結果を読む**ため、スクリプトで作っただけ(Excel で保存していない)のブックは読めません。
  Excel(ブラウザ版可)で一度開いて保存してください。画面にもその旨が表示されます。
- **WS・タスクマスタは必ず一番下に追加**し、途中の行の削除・挿入・並べ替えはしないこと。
  日程入力の入力済みの値は行の位置で対応しているため、ずれます(ずれると日程入力の「チェック」列に警告)。
  使わなくなったマスタは行を消さずに「WS自動生成=対象外」にします。
- タスクマスタは15行まで、ワークショップは40件まで。
- タスクマスタの値を変えると、過去のWSのタスクにも反映されます(作成時点の値を固定する仕組みはありません)。

## 今後の拡張

| やりたいこと | 方法 |
|---|---|
| OneDrive から直接読む(同期フォルダを使わない) | Microsoft Graph API でファイルをダウンロードし、`excel_source.load_book` に渡す。個人用アカウントは初回ログインが必要 |
| 上限(15マスタ・40WS)を増やす | `tools/make_schedule_workbook.py` の `SLOTS` / `MAX_WS_BLOCKS` を変えてブックを作り直す |
| 営業日計算(土日祝を除く) | タスク一覧の終了日・起点日の数式を `WORKDAY` 系に置き換える |
