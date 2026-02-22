# Word アドイン：文書校正アシスタント

OpenAI API を利用して、Microsoft Word 上で直接、文章の校正を行うためのアドインです。**ローカルサーバーの起動は不要です。**

執筆中の文書をワンクリックで解析し、誤字脱字・表記揺れ・不自然な表現などをチェック。修正候補を Word のサイドパネルに分かりやすく表示します。

---

## 仕組みの概要

アドインのファイルは GitHub Pages（GitHub が提供する無料の HTTPS ホスティング）上に置かれます。Word はインターネット経由でこのファイルを読み込みます。文書の内容は OpenAI API にのみ送信され、GitHub には一切送られません。

```
Word ←→ GitHub Pages（アドインのコード）
Word ←→ OpenAI API（文書の内容を送信・校正結果を受信）
```

---

## 必要なもの

| 必要なもの | 説明 |
|-----------|------|
| **GitHub アカウント** | 無料の通常アカウントで可 |
| **OpenAI API キー** | `sk-` または `sk-proj-` で始まるキー。https://platform.openai.com/api-keys で発行 |

---

## セットアップ手順

### ステップ 1：GitHub にリポジトリを作成してファイルをアップロードする

1. https://github.com にログインします。

2. 右上の「**+**」→「**New repository**」をクリックします。

3. 以下のように設定します。

   | 項目 | 設定値 |
   |------|--------|
   | Repository name | `word-proofreader`（任意の名前でも可） |
   | Visibility | **Public** |
   | Initialize this repository | チェックしない |

4. 「**Create repository**」をクリックします。

5. 作成されたリポジトリのページで「**uploading an existing file**」をクリックします。

6. このフォルダ内の以下のファイルをすべてドラッグ＆ドロップしてアップロードします。

   ```
   taskpane.html
   taskpane.css
   taskpane.js
   commands.html
   assets/（フォルダごと）
   ```

   > `assets` フォルダはフォルダごとアップロードしてください。

7. 「**Commit changes**」をクリックしてアップロードを確定します。

---

### ステップ 2：GitHub Pages を有効にする

1. リポジトリのページで「**Settings**」タブをクリックします。

2. 左メニューの「**Pages**」をクリックします。

3. 「**Branch**」のドロップダウンで「**main**」を選択し、「**Save**」をクリックします。

4. しばらく待つと、以下のような URL が表示されます。

   ```
   https://あなたのユーザー名.github.io/word-proofreader/
   ```

---

### ステップ 3：manifest.xml を生成して Word に登録する

#### Mac の場合

ターミナルを開き、このフォルダに移動して以下を実行します。

```bash
chmod +x setup_manifest.sh
./setup_manifest.sh
```

#### Windows の場合

PowerShell を開き、このフォルダに移動して以下を実行します。

```powershell
Set-ExecutionPolicy Unrestricted -Scope Process
.\setup_manifest.ps1
```

GitHub ユーザー名とリポジトリ名を入力すると、`manifest.xml` が自動生成され、Word の設定フォルダにも自動でコピーされます。

---

### ステップ 4：Word でアドインを使う

1. Word を起動（または再起動）します。
2. 「**ホーム**」タブの右端に「**校正ツール**」グループと「**文書を校正する**」ボタンが表示されます。
3. ボタンをクリックすると、画面右側にサイドパネルが開きます。
4. 初回のみ、「設定」欄に OpenAI API キーを入力して「設定を保存」をクリックします。
5. 「**文書を校正する**」ボタンをクリックすると、校正が始まります。

---

## モデルが使えない場合（「does not have access to model」エラー）

OpenAI の **プロジェクト API キー**（`sk-proj-` で始まるキー）を使用している場合、プロジェクトの設定によっては特定のモデルが利用できないことがあります。

その場合は、サイドパネルのモデル選択欄の右にある **↻ ボタン**をクリックしてください。お使いの API キーで実際に利用可能なモデルの一覧が自動取得され、選択肢が更新されます。

---

## トラブルシューティング

### アドインのボタンが Word に表示されない

- `manifest.xml` が正しいフォルダに配置されているか確認してください。
- Word を完全に終了してから再起動してください。

### 「Not Found」エラーが表示される

GitHub Pages の URL が正しく設定されているか確認してください。`manifest.xml` 内の URL と、GitHub Pages で発行された URL が一致しているか確認します。

### API エラーが表示される

- API キーが正しいか確認してください。
- ↻ ボタンでモデル一覧を取得し、利用可能なモデルを選択してください。
- OpenAI のダッシュボード（https://platform.openai.com/usage）で利用枠が残っているか確認してください。

---

## アンインストール

1. Word を終了します。
2. `manifest.xml` を配置したフォルダから削除します。
   - **Mac**: `~/Library/Containers/com.microsoft.Word/Data/Documents/wef/manifest.xml`
   - **Windows**: `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\manifest.xml`
3. （任意）GitHub のリポジトリを削除します。
