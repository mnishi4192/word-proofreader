# 文書校正アシスタント - Word アドイン

OpenAI API を使って、Word 文書をワンクリックで日本語校正するアドインです。  
**ローカルサーバー不要**。GitHub Pages でホストするため、一度セットアップすれば常時利用できます。

---

## セットアップ手順

### Step 1: GitHub リポジトリを作成する

1. [GitHub](https://github.com) にログインする
2. 右上の「+」→「New repository」をクリック
3. Repository name に任意の名前を入力（例: `word-proofreader`）
4. **Public** を選択
5. 「Create repository」をクリック

### Step 2: ファイルをアップロードする

1. 作成したリポジトリのページで「uploading an existing file」をクリック
2. 以下のファイルをすべてアップロードする：
   - `taskpane.html`
   - `taskpane.js`
   - `taskpane.css`
   - `commands.html`
   - `manifest_template.xml`
   - `setup_manifest.sh`（Mac）または `setup_manifest.ps1`（Windows）
   - `assets/` フォルダ内のアイコン画像 3 枚

   > **注意**: `assets/` フォルダごとアップロードする場合は、フォルダをドラッグ＆ドロップしてください。

3. 「Commit changes」をクリック

### Step 3: GitHub Pages を有効にする

1. リポジトリの「Settings」タブを開く
2. 左メニューの「Pages」をクリック
3. Source を「**Deploy from a branch**」に設定
4. Branch を「**main**」、フォルダを「**/ (root)**」に設定して「Save」
5. しばらく待つと「Your site is published at `https://ユーザー名.github.io/リポジトリ名/`」と表示される

この URL をメモしておく（例: `https://yamada.github.io/word-proofreader`）

### Step 4: manifest.xml を生成して Word に登録する

**Mac の場合:**

```bash
# ダウンロードしたフォルダに移動
cd ~/Downloads/addin-final

# スクリプトに実行権限を付与
chmod +x setup_manifest.sh

# セットアップを実行
./setup_manifest.sh
```

プロンプトが表示されたら、Step 3 でメモした GitHub Pages の URL を入力する。

**Windows の場合:**

1. `setup_manifest.ps1` を右クリック →「PowerShell で実行」
2. プロンプトが表示されたら GitHub Pages の URL を入力する
3. 表示された手順に従って Word のトラスト センターに登録する

### Step 5: Word でアドインを起動する

1. Word を起動（または再起動）する
2. 「挿入」タブ →「アドイン」→「個人用アドイン」
3. 「文書校正アシスタント」を選択してクリック
4. サイドパネルが開いたら、OpenAI API キーを入力して「保存」

---

## 使い方

1. Word で校正したい文書を開く
2. 「文書校正アシスタント」のサイドパネルを開く
3. 「文書を校正する」ボタンをクリック
4. 校正結果がサイドパネルに表示される

---

## ファイル構成

```
addin-final/
├── taskpane.html          # メインの UI
├── taskpane.js            # 校正ロジック（ストリーミング対応）
├── taskpane.css           # スタイルシート
├── commands.html          # リボンボタン用（空ページ）
├── manifest_template.xml  # マニフェストのテンプレート
├── setup_manifest.sh      # Mac 用セットアップスクリプト
├── setup_manifest.ps1     # Windows 用セットアップスクリプト
└── assets/
    ├── icon-16.png
    ├── icon-32.png
    └── icon-64.png
```

---

## よくある質問

**Q: 文書の内容が GitHub に保存されますか？**  
A: いいえ。GitHub には HTML/JS/CSS などのプログラムコードのみが置かれます。文書の内容は、ボタンを押した瞬間に Word から直接 OpenAI API に送信されます。GitHub のサーバーを経由することはありません。

**Q: API キーはどこに保存されますか？**  
A: お使いの PC のブラウザ内（localStorage）にのみ保存されます。GitHub には一切アップロードされません。

**Q: どのモデルを使えばいいですか？**  
A: `↻` ボタンを押すと、お使いの API キーで利用可能なモデルの一覧が自動取得されます。そこから選択してください。高精度な校正には `gpt-4o` または `gpt-4.1` をお勧めします。

**Q: 処理が遅いです。**  
A: 文書が長いほど処理に時間がかかります。処理中は「○○ 字受信済み」とリアルタイムで進捗が表示されます。表示が更新されている間は正常に処理中です。

**Q: 「このモデルは使用できません」というエラーが出ます。**  
A: お使いの OpenAI プロジェクトでそのモデルへのアクセスが許可されていません。`↻` ボタンで利用可能なモデルを取得し、別のモデルを選択してください。
