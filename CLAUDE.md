# CLAUDE.md

Google スプレッドシートに紐付く Google Apps Script (GAS) プロジェクト群。各ディレクトリが 1 つの Apps Script プロジェクトで、`clasp` で push/pull する(`.clasp.json` の `scriptId` が対応先)。ビルド・テスト・lint の仕組みは無い。

## ディレクトリ構成

| パス | プロジェクト | 状態 |
|---|---|---|
| `stock-sheet/main/` | 株式情報(日本株 終値自動取得システム、スプレッドシート紐付き) | **現行** |
| `stock-sheet/db/` | 株式情報_DB(旧ライブラリ `StockLib`、`openById` でメインを操作) | **旧版**。2026/10 に main へ統合済み |
| `pl-sheet/` | 損益管理(月次シート自動作成・週間累計) | 現行 |

- `stock-sheet/db/コード.js` は統合前のコード。新規の修正は `stock-sheet/main/` に入れる。db 側は参照用で、触る場合はユーザーに確認する。
- 時間主導トリガーは main 側に設定済み(一覧は `00_設定・共通.js` 冒頭のコメント)。投資部門別(`updateJpxStatsForLatestWeek`)は毎日16時台に動かし、JPX の掲載日(毎週第4営業日15:30、祝日分は後ろ倒し)を計算して、掲載済みの未取得週があるときだけ取得する。設定用の関数は削除したので、作り直すときは Apps Script エディタの「トリガー」から手動で追加する。DB 側にトリガーが残っていると二重実行になる。

## stock-sheet/main のファイル構成

Apps Script エディタ上でこの番号順に並べる(`00` が先頭でないと定数が未定義になる)。

- `00_設定・共通.js` — シート構成の定数、`getMainSpreadsheet()`、祝日判定、PDF→テキスト変換などの共通処理
- `01_メニュー.js` — `onOpen` / 「機能」メニュー、ダイアログの呼び出し口
- `02_銘柄操作.js` — 銘柄の追加・削除・並び替え、色設定(`COLOR_SETTINGS` をスクリプトプロパティに保存)
- `03_業種.js` — 業種情報の取得・設定
- `04_終値・市場.js` — 終値・前日比・売買代金(Yahoo チャート API)、市場ブロック、来月シート作成
- `05_データベース・書式.js` — データベースシート、スパークライン、書式の再適用
- `06_信用残.js` — JPX 日次信用残 PDF(mtall.pdf)を自前の PDF パーサ(`mt*_` 関数)で解析
- `07_空売り集計.js` — 空売り集計シート
- `08_投資部門別.js` — 投資部門別売買状況(売買状況シート)
- `09_チャート.js` — `ChartDialog` 用データ組み立て
- `10_決算情報.js` — 「決算情報取得銘柄」ダイアログ(`KessanTargetDialog`)。選んだ銘柄を別ファイル `決算情報` の `決算情報取得銘柄` シートに書き出す(ファイル ID はスクリプトプロパティ `KESSAN_INFO_SPREADSHEET_ID`)。決算スケジュールや決算の評価は Claude の会話側で行い、このシートを対象銘柄として読む
- `*Dialog.html` — `HtmlService` のモーダルダイアログ。サーバー関数は `google.script.run` で呼ぶ

有効な高度なサービス: Sheets API v4, Drive API v3(`appsscript.json`)。

## シートのレイアウト(stock-sheet)

- 月次シート名は `{月}月{西暦下2桁}年`(例: `9月26年`)。`テンプレート` シートから作成する。
- 5 行目=見出し、6〜36 行目=日次(1〜31 日)の終値等、38 行目=信用残見出し、39〜69 行目=日次の信用残。
- C〜K 列はプライム/スタンダード/グロースの市場ブロック(3 列×3)。通常銘柄は L 列(`STOCK_START_COL = 12`)から 3 列単位。
- 行・列位置はすべて `00_設定・共通.js` の定数経由で扱う。直書きしない。
- JPX の日次信用残 PDF(mtall.pdf)は直近5営業日分しか公開されない。読み取った全銘柄分の結果はドライブの `株式情報_信用残データ` フォルダに1日1ファイル(`mtall_yyyyMMdd.json.gz`)で保存し、PDF が消えた後はそこから読む(`06_信用残.js`)。毎日の `updateMarginData` が直近分を保存する。

## pl-sheet

損益管理。詳細は `pl-sheet/CLAUDE.md`。

## チャット(セッション)の分担

- 株式情報(`stock-sheet/`)と損益管理(`pl-sheet/`)は別々のチャットで扱う。1 つのチャットでは担当のディレクトリだけを変更する。
- 共有のファイル(ルートの `CLAUDE.md`、`.github/workflows/clasp-push.yml`)を変える場合は、もう一方にも影響することをユーザーに伝える。
- 自動デプロイは、どちらのディレクトリの変更でも両方のプロジェクトを `clasp push` する(main の内容で上書きするだけなので、もう一方の中身は変わらない)。

## コーディング規約

- コメント・UI 文言・ログは日本語。ファイル名も日本語。
- V8 ランタイム、`const` / `let`、トップレベルの `function` 宣言。ファイルはすべてグローバルスコープを共有するので、関数名・定数名の衝突に注意。
- 内部用の関数は末尾に `_`(例: `parseDateStr_`)。末尾 `_` の関数はメニューや `google.script.run` から呼べない。
- 外部取得は `UrlFetchApp.fetch(url, { muteHttpExceptions: true })` でステータスコードを確認する。
- 長時間処理は `RUN_TIME_LIMIT_MS`(4.5 分)で打ち切り、進捗を `PropertiesService` に保存して再実行で続きから進める(GAS の 6 分制限対策)。
- メニュー実行のラッパーは `runLibFromMenu_(fn, 既定メッセージ, 失敗時の接頭辞)` を使う。
- 銘柄の列を動かす処理(追加・削除・並び替え)は `withStockColumnLock_` で囲み、同時に実行しない(銘柄追加はサイドバー、並び替えはモードレスダイアログなので、並行して操作できる)。
- タイムゾーンは `Asia/Tokyo`。祝日は Google の「日本の祝日」カレンダーから国民の祝日名だけを採用する(七夕などの行事を除外)。

## デプロイ

main ブランチに `stock-sheet/main/` か `pl-sheet/` の変更が入ると、GitHub Actions(`.github/workflows/clasp-push.yml`)が両プロジェクトを自動で `clasp push` する。認証情報はリポジトリの Actions Secrets `CLASPRC_JSON`。Apps Script 側はリポジトリの内容で丸ごと上書きされるので、エディタで直接編集しない。

手動で反映する場合:

```sh
cd stock-sheet/main && clasp push   # pl-sheet も同様
```

ローカルで実行・テストする手段は無い。動作確認は Apps Script エディタまたはスプレッドシート上で行う。
