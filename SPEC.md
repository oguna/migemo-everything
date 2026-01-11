# migemo-everything 仕様書

本書は Rust 製 Windows アプリ「migemo-everything」の再現手順と動作仕様をまとめたものです。実装詳細に依存しないレベルで、求められる画面構成・挙動・依存物を定義します。

## 対象環境
- OS: Windows 10 以降 (x64 を想定)
- ランタイム: Everything が起動済みで、同梱の `Everything64.dll` を利用できること
- フォント: 標準 UI フォント (Segoe UI) を使用
- DPI: Per-monitor DPI 対応。DPI 変更に応じてコントロールのサイズ・配置を再計算する

## ビルド・配置
1. Rust (stable) をインストール済みであること。
2. ルートで `cargo build --release` を実行し、`target/release/migemo-everything.exe` を得る。
3. 実行ファイルと同じディレクトリに以下を配置する。
   - `Everything64.dll` (配布物に同梱)
   - `migemo-compact-dict` (https://github.com/oguna/yet-another-migemo-dict から取得)
4. Everything 本体を起動した状態で `migemo-everything.exe` を実行する。

## 起動時初期化
- COM を STA で初期化し、終了時に Uninitialize する。
- プロセス DPI 認識を有効化 (`SetProcessDpiAwarenessContext(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2)` 相当)。
- Migemo 辞書を `migemo-compact-dict` から読み込み。読み込み失敗時も起動は続行する。
- ウィンドウタイトルは「Migemo Everything」。検索語入力時は「<検索語> - Migemo Everything」に更新される。

## UI レイアウト (DPI スケール適用後の論理サイズ)
- 上部バー (高さ 25 * scale)
  - 左: 検索テキストボックス (単一行、Auto-scroll)
  - 右端: 幅 40 * scale のボタンを 2 つ横並び  
    - 「RE」: 正規表現トグル
    - 「Mi」: Migemo トグル
- 中央: 仮想リストビュー (`LVS_REPORT | LVS_OWNERDATA`、小アイコン付き、垂直/水平スクロール)  
  カラムは左から「名前」「フォルダ」「サイズ (右寄せ)」「更新日時」。
- 下部ステータスバー (高さ 20 * scale)
  - 左: ステータス文字列 (`Ready` または `<件数> items found`)
  - 右: 幅 100 * scale のチェックボックス「Shell Menu」(シェルコンテキストメニューの有効/無効)

## キーボードショートカット
- `Ctrl+Q`: アプリ終了
- `Ctrl+R`: 正規表現検索トグル (ON 時は Migemo を自動で OFF)
- `Ctrl+Shift+R`: Migemo 検索トグル (ON 時は正規表現を自動で OFF)
- ウィンドウがフォーカスを得た際、検索ボックスへフォーカスを戻す。

## 検索挙動
- 入力ボックス変更時: 500ms のタイマー後に検索実行。連続入力時はタイマーをリセット。
- トグル操作時 (`RE` / `Mi`): 100ms の短いタイマーで検索を走らせる。
- 検索語が空の場合: 検索結果をクリアし、件数 0、タイトルを初期化、ステータスを `Ready` に戻す。
- Migemo が有効な場合: 辞書で検索語を展開し、展開後の文字列を Everything 検索に使用。
- Everything へのクエリ:
  - リクエストフラグ: ファイル名、パス、サイズ、更新日時、属性、ハイライト済みファイル名/パスを要求。
  - 検索モード: 正規表現は「正規表現 ON または Migemo ON」で有効。
  - 初回取得: 100 件を取得し総件数を保存。`page_size` は 100。
  - 仮想リスト: 要求インデックスが未ロードの場合、`offset` をインデックスに合わせて 100 件ずつ追加入手。
- ステータスバーには `<総件数> items found` を表示し、リストビューのアイテム数を総件数に設定。

## リストビュー表示
- アイコン: システムイメージリストの小アイコンを使用。フォルダかファイルかで属性を切替えて `SHGetFileInfoW` からインデックス取得。
- テキスト:
  - 「名前」および「フォルダ」カラムは Everything のハイライト情報（`*` で囲まれた範囲）をパースし、非選択時はハイライト部分の背景をシアン系で塗る。
  - サイズは 3 桁ごとにカンマ区切り、更新日時は `YYYY-MM-DD HH:MM` の 24 時間表記。
- ダブルクリック: 該当パスを `ShellExecuteW(..., "open")` で開く。

## コンテキストメニュー
- 右クリック時の動作は「Shell Menu」チェックボックスで切替。
  - OFF（既定）: カスタムメニュー  
    - `開く`: アイテムを開く  
    - `フォルダを開く`: エクスプローラで選択状態で開く  
    - `フルパスをコピー`: パス + ファイル名をクリップボードへ UTF-16 でコピー  
    - 既定選択は「開く」
  - ON: シェル提供のコンテキストメニューをそのまま表示し、選択コマンドを `IContextMenu::InvokeCommand` で実行。
- コンテキストメニュー用にアイテム情報を事前取得し、メニュー表示前にロックを解放してデッドロックを回避。

## クリップボード操作
- `CF_UNICODETEXT` でフルパス文字列をセット。Open/Empty/SetClipboardData の Win32 API を使用。

## DPI/リサイズ
- `WM_DPICHANGED` で新 DPI を取得しスケールを再計算。提示された矩形に合わせてウィンドウを再配置し、無効領域を再描画。
- `WM_SIZE` で現在サイズに応じてコントロールを再配置。

## 終了
- `Ctrl+Q` もしくはメニュー/アクセラレータ/ウィンドウクローズ操作で `DestroyWindow` を実行し、メッセージループ終了後に COM を解放して終了。

## 依存関係のサマリ
- Rust crates: windows 0.62.2 系、everything-sdk、rustmigemo、その他ロックファイルに準拠。
- 外部ファイル: `Everything64.dll`、`migemo-compact-dict`
- 外部プロセス: Everything (検索対象のインデックスを提供)
