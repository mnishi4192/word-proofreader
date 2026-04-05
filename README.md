# 文書校正アシスタント（Caddy ローカルホスト版）

**インターネット接続不要**で動作する Word アドインです。  
アドインの UI ファイルも Caddy がローカルで配信するため、GitHub Pages への依存がありません。  
校正処理には Ollama（ローカル LLM）を使用し、文書の内容が外部に送信されることはありません。

> **クラウド API（OpenAI / Gemini / Claude）も引き続き利用できます。**  
> ただし、それらを使う場合はインターネット接続と各サービスの API キーが必要です。

---

## 構成概要

```
Word
 └─ アドイン UI（taskpane.html 等）
      └─ https://localhost:11436  ← Caddy が配信（インターネット不要）
           ├─ /office.js          ← ローカルに保存した office.js
           ├─ /taskpane.html
           ├─ /taskpane.js
           └─ /taskpane.css

Caddy リバースプロキシ
 └─ https://localhost:11435  →  http://localhost:11434（Ollama）
```

---

## セットアップ手順

### Step 1: Caddy をインストールする

**Mac**
```bash
brew install caddy
```

**Windows（管理者権限の PowerShell）**
```powershell
winget install Caddy.Caddy
```

---

### Step 2: office.js をローカルに取得する

office.js は Microsoft の CDN から**一度だけ**ダウンロードします。  
以降はインターネット接続なしで動作します。

**Mac / Linux**
```bash
chmod +x fetch_office_js.sh
./fetch_office_js.sh
```

**Windows**
```powershell
.\fetch_office_js.ps1
```

---

### Step 3: セットアップスクリプトを実行する

セットアップスクリプトが以下を自動で行います。

- Caddy のローカル CA 証明書を OS に登録（`caddy trust`）
- `manifest_local.xml` を Word のアドインフォルダにコピー

**Mac**
```bash
chmod +x setup_manifest.sh
./setup_manifest.sh
```

**Windows（管理者権限の PowerShell）**
```powershell
.\setup_manifest.ps1
```

---

### Step 4: Mac のキーチェーン設定（Mac のみ）

`caddy trust` を実行しても、Mac の WKWebView（Word が使う WebView）が証明書を信頼するには追加の設定が必要です。

1. 「キーチェーンアクセス」を開く（Spotlight で検索）
2. 左側の「システム」キーチェーンを選択
3. 「Caddy Local Authority」を見つけてダブルクリック
4. 「信頼」セクションを開く
5. 「この証明書を使用するとき」を **「常に信頼」** に変更
6. ウィンドウを閉じる（パスワードを求められたら入力）

---

### Step 5: Ollama をインストールしてモデルを取得する

1. [https://ollama.com](https://ollama.com) から Ollama をインストール
2. ターミナルでモデルをダウンロード（例）:

```bash
ollama pull llama3.2        # 軽量・高速（推奨）
ollama pull qwen2.5:14b     # 日本語性能が高い
ollama pull gemma3:12b      # Google 製
```

---

### Step 6: Caddy を起動する

このフォルダで以下を実行します。

```bash
caddy run --config Caddyfile
```

**PC 起動時に自動起動する場合:**

```bash
# Mac
brew services start caddy

# Windows（管理者 PowerShell）
caddy service install
caddy service start
```

> **注意（Mac の自動起動）:** `brew services` で起動する場合、`Caddyfile` の `root` ディレクティブに絶対パスを指定してください。  
> 例: `root * /Users/yourname/word-proofreader-caddy`

---

### Step 7: Word でアドインを起動する

1. Word を再起動する
2. 「挿入」タブ →「アドイン」→「個人用アドイン」
3. 「文書校正アシスタント（ローカル）」を選択
4. サイドパネルが開いたら「Ollama」タブを選択
5. サーバー URL に `https://localhost:11435` を入力して「保存」
6. 「↻」ボタンでモデル一覧を取得し、使用するモデルを選択

---

## 使い方

1. Word で校正したい文書を開く
2. 「文書校正アシスタント（ローカル）」のサイドパネルを開く
3. 「文書を校正する」ボタンをクリック
4. 校正結果がサイドパネルに表示される

---

## ファイル構成

```
addin-caddy/
├── taskpane.html          # メインの UI
├── taskpane.js            # 校正ロジック（ストリーミング対応）
├── taskpane.css           # スタイルシート
├── commands.html          # リボンボタン用（空ページ）
├── office.js              # ★ fetch_office_js.sh/ps1 で取得（要実行）
├── Caddyfile              # Caddy 設定ファイル
├── manifest_local.xml     # Word アドインマニフェスト（ローカル版）
├── manifest_template.xml  # GitHub Pages 版マニフェストテンプレート（参考用）
├── fetch_office_js.sh     # office.js 取得スクリプト（Mac）
├── fetch_office_js.ps1    # office.js 取得スクリプト（Windows）
├── setup_manifest.sh      # セットアップスクリプト（Mac）
├── setup_manifest.ps1     # セットアップスクリプト（Windows）
└── assets/
    ├── icon-16.png
    ├── icon-32.png
    └── icon-64.png
```

> `office.js` はリポジトリに含まれていません。`fetch_office_js.sh`（Mac）または `fetch_office_js.ps1`（Windows）を実行して取得してください。

---

## よくある質問

**Q: 文書の内容がどこかに送信されますか？**  
A: Ollama を使用する場合、文書の内容は外部に送信されません。すべての処理がローカルで完結します。OpenAI / Gemini / Claude を使用する場合は、各サービスの API に送信されます。

**Q: インターネット接続は本当に不要ですか？**  
A: `office.js` を一度ダウンロードした後は、Ollama タブを使用する限りインターネット接続は不要です。OpenAI / Gemini / Claude タブを使用する場合はインターネット接続と API キーが必要です。

**Q: Caddy を起動し忘れるとどうなりますか？**  
A: Word でアドインを開こうとしても「ページを読み込めません」というエラーになります。`caddy run --config Caddyfile` を実行してから Word を再起動してください。自動起動の設定をお勧めします。

**Q: Mac でキーチェーンの設定をしたのに接続できません。**  
A: Word を完全に終了してから再起動してください。Word の WebView は起動時に証明書の信頼情報を読み込むため、設定変更後は再起動が必要です。

**Q: Ollama のモデルを追加したい。**  
A: ターミナルで `ollama pull モデル名` を実行してください。アドインの「↻」ボタンを押すと新しいモデルが一覧に表示されます。
