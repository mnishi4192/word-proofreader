/* =========================================================
   文書校正アシスタント - taskpane.js
   ========================================================= */

'use strict';

// ===== システムプロンプト =====
const SYSTEM_PROMPT = `このGPTは、編集者の視点から日本語文書を推敲し、改善を提案します。
特に、誤字脱字、変換ミス、表記ミスに注目し、数字やアルファベットの表記が半角で統一されているかを確認します。文章表現の技巧にはあまり踏み込まず、かなやアルファベットで複数回出てくる固有名詞の表記揺れ、削除可能な指示代名詞、同じ文頭や文末表現が連続する部分、同音異義語の誤用を指摘します。読みやすさについても評価してください。
指摘は具体的に、何行目のどの部分に関するものかを、文章の形で示します。
文書が「■」で区切られている場合、区切りごとに内容を要約し、全体の構成が妥当で論理的かを評価します。指摘や提案は、アップロードされた文書の中に含まれる内容に限定して行います。
全ての返答は日本語で行います。
チェック後には、修正部分を【修正前→修正後】の形式で示してください。

返答は以下の構造で行ってください：

## 【校正結果】

### 1. 誤字・脱字・表記ミス
（該当箇所を「○行目：〜」の形式で列挙。なければ「指摘なし」）

### 2. 表記の統一（固有名詞・数字・アルファベット）
（該当箇所を列挙。なければ「指摘なし」）

### 3. 文章表現・読みやすさ
（指示代名詞・連続する文頭/文末・同音異義語などを列挙。なければ「指摘なし」）

### 4. 構成・論理性の評価
（「■」区切りがある場合は各セクションを要約し評価。ない場合は全体を評価）

### 5. 総評
（文書全体の品質について簡潔に総評）`;

// ===== DOM 要素の取得 =====
let apiKeyInput, modelSelect, saveSettingsBtn, settingsSavedMsg;
let toggleKeyBtn, proofreadBtn, btnText, btnSpinner;
let progressArea, progressText;
let resultsSection, resultsMeta, resultsContent;
let copyBtn, errorArea, errorMessage;

// ===== Office.js 初期化 =====
Office.onReady(function (info) {
  if (info.host === Office.HostType.Word) {
    initDOM();
    loadSettings();
    bindEvents();
    // APIキーが設定済みなら校正ボタンを有効化
    const savedKey = localStorage.getItem('proofreader_api_key');
    if (savedKey && savedKey.trim().startsWith('sk-')) {
      proofreadBtn.disabled = false;
    }
  }
});

// ===== DOM 初期化 =====
function initDOM() {
  apiKeyInput      = document.getElementById('api-key');
  modelSelect      = document.getElementById('model-select');
  saveSettingsBtn  = document.getElementById('save-settings');
  settingsSavedMsg = document.getElementById('settings-saved');
  toggleKeyBtn     = document.getElementById('toggle-key');
  proofreadBtn     = document.getElementById('proofread-btn');
  btnText          = document.getElementById('btn-text');
  btnSpinner       = document.getElementById('btn-spinner');
  progressArea     = document.getElementById('progress-area');
  progressText     = document.getElementById('progress-text');
  resultsSection   = document.getElementById('results-section');
  resultsMeta      = document.getElementById('results-meta');
  resultsContent   = document.getElementById('results-content');
  copyBtn          = document.getElementById('copy-btn');
  errorArea        = document.getElementById('error-area');
  errorMessage     = document.getElementById('error-message');
}

// ===== 設定の読み込み =====
function loadSettings() {
  const savedKey   = localStorage.getItem('proofreader_api_key') || '';
  const savedModel = localStorage.getItem('proofreader_model') || 'gpt-4o';
  apiKeyInput.value = savedKey;
  modelSelect.value = savedModel;
}

// ===== イベントバインド =====
function bindEvents() {
  // APIキー表示/非表示
  toggleKeyBtn.addEventListener('click', function () {
    apiKeyInput.type = apiKeyInput.type === 'password' ? 'text' : 'password';
  });

  // 設定保存
  saveSettingsBtn.addEventListener('click', function () {
    const key   = apiKeyInput.value.trim();
    const model = modelSelect.value;
    localStorage.setItem('proofreader_api_key', key);
    localStorage.setItem('proofreader_model', model);
    settingsSavedMsg.style.display = 'inline';
    setTimeout(() => { settingsSavedMsg.style.display = 'none'; }, 2000);
    // キーが有効そうなら校正ボタンを有効化
    proofreadBtn.disabled = !(key && key.startsWith('sk-'));
  });

  // APIキー入力時にリアルタイムでボタン状態を更新
  apiKeyInput.addEventListener('input', function () {
    const key = apiKeyInput.value.trim();
    proofreadBtn.disabled = !(key && key.startsWith('sk-'));
  });

  // 校正実行
  proofreadBtn.addEventListener('click', runProofread);

  // 結果コピー
  copyBtn.addEventListener('click', function () {
    const text = resultsContent.innerText;
    navigator.clipboard.writeText(text).then(() => {
      copyBtn.textContent = '✓ コピー済み';
      setTimeout(() => { copyBtn.textContent = 'コピー'; }, 2000);
    });
  });
}

// ===== 校正実行メイン処理 =====
async function runProofread() {
  const apiKey = localStorage.getItem('proofreader_api_key') || '';
  const model  = localStorage.getItem('proofreader_model') || 'gpt-4o';

  if (!apiKey || !apiKey.startsWith('sk-')) {
    showError('APIキーが設定されていません。上の「設定」欄にOpenAI APIキーを入力して保存してください。');
    return;
  }

  // UI: 実行中状態へ
  setLoading(true);
  hideError();
  resultsSection.style.display = 'none';
  progressArea.style.display = 'block';
  setProgress('文書のテキストを取得中...');

  try {
    // Step 1: Word文書のテキストを取得
    const documentText = await getDocumentText();

    if (!documentText || documentText.trim().length === 0) {
      throw new Error('文書にテキストが見つかりませんでした。文書にテキストを入力してから再試行してください。');
    }

    setProgress('OpenAI APIに送信中...');

    // Step 2: OpenAI API で校正
    const result = await callOpenAI(apiKey, model, documentText);

    setProgress('結果を表示中...');

    // Step 3: 結果を表示
    displayResults(result, documentText, model);

  } catch (err) {
    showError(formatError(err));
  } finally {
    setLoading(false);
    progressArea.style.display = 'none';
  }
}

// ===== Word 文書テキスト取得 =====
async function getDocumentText() {
  return new Promise((resolve, reject) => {
    Word.run(async function (context) {
      try {
        const body = context.document.body;
        body.load('text');
        await context.sync();
        resolve(body.text);
      } catch (e) {
        reject(new Error('文書の読み取りに失敗しました: ' + e.message));
      }
    }).catch(reject);
  });
}

// ===== OpenAI API 呼び出し =====
async function callOpenAI(apiKey, model, documentText) {
  const userMessage = `以下の文書を校正してください。\n\n---\n${documentText}\n---`;

  const response = await fetch('https://api.openai.com/v1/chat/completions', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${apiKey}`,
    },
    body: JSON.stringify({
      model: model,
      messages: [
        { role: 'system', content: SYSTEM_PROMPT },
        { role: 'user',   content: userMessage },
      ],
      temperature: 0.2,
      max_tokens: 4096,
    }),
  });

  if (!response.ok) {
    const errData = await response.json().catch(() => ({}));
    const errMsg = errData.error?.message || `HTTPエラー ${response.status}`;
    throw new Error(`OpenAI APIエラー: ${errMsg}`);
  }

  const data = await response.json();
  return data.choices?.[0]?.message?.content || '（結果を取得できませんでした）';
}

// ===== 結果表示 =====
function displayResults(rawText, documentText, model) {
  // メタ情報
  const charCount = documentText.length;
  const lineCount = documentText.split('\n').length;
  const now = new Date().toLocaleString('ja-JP');
  resultsMeta.innerHTML =
    `使用モデル: <strong>${model}</strong> ／ ` +
    `文字数: <strong>${charCount.toLocaleString()}</strong> 字 ／ ` +
    `行数: <strong>${lineCount}</strong> 行 ／ ` +
    `実行日時: ${now}`;

  // Markdown を HTML に変換して表示
  resultsContent.innerHTML = markdownToHTML(rawText);

  resultsSection.style.display = 'block';
  // 結果セクションまでスクロール
  resultsSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

// ===== 簡易 Markdown → HTML 変換 =====
function markdownToHTML(text) {
  // 【修正前→修正後】パターンを赤字表示
  text = text.replace(/【(.+?)→(.+?)】/g,
    '<span class="correction">【<del>$1</del> → $2】</span>');

  // 見出し
  text = text.replace(/^### (.+)$/gm, '<div class="section-header">$1</div>');
  text = text.replace(/^## (.+)$/gm,  '<div class="section-header" style="font-size:14px;margin-top:8px;">$1</div>');
  text = text.replace(/^# (.+)$/gm,   '<div class="section-header" style="font-size:15px;margin-top:8px;">$1</div>');

  // 太字
  text = text.replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>');
  text = text.replace(/__(.+?)__/g,     '<strong>$1</strong>');

  // 箇条書き
  text = text.replace(/^[-*] (.+)$/gm, '<div class="issue-item">• $1</div>');

  // 番号付きリスト
  text = text.replace(/^\d+\. (.+)$/gm, '<div class="issue-item info">$1</div>');

  // 改行
  text = text.replace(/\n{2,}/g, '<br><br>');
  text = text.replace(/\n/g, '<br>');

  return text;
}

// ===== エラーフォーマット =====
function formatError(err) {
  if (err.message.includes('401')) {
    return 'APIキーが無効です。正しいOpenAI APIキーを設定してください。';
  }
  if (err.message.includes('429')) {
    return 'APIのレート制限に達しました。しばらく待ってから再試行してください。';
  }
  if (err.message.includes('insufficient_quota')) {
    return 'OpenAI APIの利用枠が不足しています。OpenAIのダッシュボードで残高を確認してください。';
  }
  return err.message || '不明なエラーが発生しました。';
}

// ===== UI ヘルパー =====
function setLoading(isLoading) {
  proofreadBtn.disabled = isLoading;
  btnText.textContent   = isLoading ? '校正中...' : '文書を校正する';
  btnSpinner.style.display = isLoading ? 'inline-block' : 'none';
}

function setProgress(msg) {
  progressText.textContent = msg;
}

function showError(msg) {
  errorMessage.textContent = msg;
  errorArea.style.display  = 'block';
}

function hideError() {
  errorArea.style.display = 'none';
}
