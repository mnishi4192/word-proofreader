/* =========================================================
   文書校正アシスタント - taskpane.js  v4
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
let toggleKeyBtn, fetchModelsBtn, modelHint;
let proofreadBtn, btnText, btnSpinner;
let progressArea, progressText;
let resultsSection, resultsMeta, resultsContent;
let copyBtn, errorArea, errorMessage;

// ===== Office.js 初期化 =====
Office.onReady(function (info) {
  if (info.host === Office.HostType.Word) {
    initDOM();
    loadSettings();
    bindEvents();
    const savedKey = localStorage.getItem('proofreader_api_key');
    if (savedKey && savedKey.trim().length > 10) {
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
  fetchModelsBtn   = document.getElementById('fetch-models-btn');
  modelHint        = document.getElementById('model-hint');
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

  // 保存済みモデルが選択肢にあればそれを選択、なければ追加して選択
  const existingOption = Array.from(modelSelect.options).find(o => o.value === savedModel);
  if (existingOption) {
    modelSelect.value = savedModel;
  } else if (savedModel) {
    const opt = document.createElement('option');
    opt.value = savedModel;
    opt.textContent = savedModel + '（保存済み）';
    modelSelect.appendChild(opt);
    modelSelect.value = savedModel;
  }
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
    proofreadBtn.disabled = !(key && key.length > 10);
  });

  // APIキー入力時にリアルタイムでボタン状態を更新
  apiKeyInput.addEventListener('input', function () {
    const key = apiKeyInput.value.trim();
    proofreadBtn.disabled = !(key && key.length > 10);
  });

  // 利用可能なモデルを取得
  fetchModelsBtn.addEventListener('click', fetchAvailableModels);

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

// ===== 利用可能なモデルを API から取得 =====
async function fetchAvailableModels() {
  const apiKey = apiKeyInput.value.trim() || localStorage.getItem('proofreader_api_key') || '';
  if (!apiKey || apiKey.length < 10) {
    modelHint.textContent = '先に API キーを入力してください。';
    modelHint.style.color = '#c0392b';
    return;
  }

  fetchModelsBtn.disabled = true;
  fetchModelsBtn.textContent = '…';
  modelHint.textContent = 'モデルを取得中...';
  modelHint.style.color = '#666';

  try {
    const response = await fetch('https://api.openai.com/v1/models', {
      headers: { 'Authorization': `Bearer ${apiKey}` }
    });

    if (!response.ok) {
      const err = await response.json().catch(() => ({}));
      throw new Error(err.error?.message || `HTTP ${response.status}`);
    }

    const data = await response.json();

    // GPT 系のチャットモデルのみ抽出してソート
    const chatModels = data.data
      .filter(m => m.id.startsWith('gpt-'))
      .map(m => m.id)
      .sort((a, b) => b.localeCompare(a));

    if (chatModels.length === 0) {
      throw new Error('利用可能な GPT モデルが見つかりませんでした。');
    }

    // 現在の選択値を保持
    const currentValue = modelSelect.value;

    // 選択肢を再構築
    modelSelect.innerHTML = '';
    const group = document.createElement('optgroup');
    group.label = `利用可能なモデル（${chatModels.length} 件）`;
    chatModels.forEach(id => {
      const opt = document.createElement('option');
      opt.value = id;
      opt.textContent = id;
      group.appendChild(opt);
    });
    modelSelect.appendChild(group);

    // 以前の選択値を復元（なければ最初の項目）
    if (chatModels.includes(currentValue)) {
      modelSelect.value = currentValue;
    } else {
      // gpt-4o があれば優先、なければ先頭
      const preferred = chatModels.find(m => m === 'gpt-4o') || chatModels[0];
      modelSelect.value = preferred;
    }

    modelHint.textContent = `✓ ${chatModels.length} 件のモデルを取得しました。お使いのプロジェクトで利用可能なモデルが表示されています。`;
    modelHint.style.color = '#27ae60';

  } catch (err) {
    modelHint.textContent = `取得失敗: ${err.message}`;
    modelHint.style.color = '#c0392b';
  } finally {
    fetchModelsBtn.disabled = false;
    fetchModelsBtn.textContent = '↻';
  }
}

// ===== 校正実行メイン処理 =====
async function runProofread() {
  const apiKey = localStorage.getItem('proofreader_api_key') || '';
  const model  = modelSelect.value || localStorage.getItem('proofreader_model') || 'gpt-4o';

  if (!apiKey || apiKey.length < 10) {
    showError('API キーが設定されていません。上の「設定」欄に OpenAI API キーを入力して保存してください。');
    return;
  }

  // UI: 実行中状態へ
  setLoading(true);
  hideError();
  resultsSection.style.display = 'none';
  progressArea.style.display = 'block';
  setProgress('文書のテキストを取得中...');

  try {
    // Step 1: Word 文書のテキストを取得
    const documentText = await getDocumentText();

    if (!documentText || documentText.trim().length === 0) {
      throw new Error('文書にテキストが見つかりませんでした。文書にテキストを入力してから再試行してください。');
    }

    setProgress(`OpenAI API（${model}）に送信中...`);

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
// GPT-5 系は Responses API（/v1/responses）を使用する
// gpt-4o / gpt-4.1 系は Chat Completions API（/v1/chat/completions）を使用する
async function callOpenAI(apiKey, model, documentText) {
  const userMessage = `以下の文書を校正してください。\n\n---\n${documentText}\n---`;

  // GPT-5 系モデルは Responses API を使用
  const isGpt5 = /^gpt-5/.test(model);

  if (isGpt5) {
    return await callResponsesAPI(apiKey, model, userMessage);
  } else {
    return await callChatCompletionsAPI(apiKey, model, userMessage);
  }
}

// GPT-5 系: Responses API
async function callResponsesAPI(apiKey, model, userMessage) {
  const requestBody = {
    model: model,
    instructions: SYSTEM_PROMPT,
    input: userMessage,
    max_output_tokens: 4096,
  };

  const response = await fetch('https://api.openai.com/v1/responses', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${apiKey}`,
    },
    body: JSON.stringify(requestBody),
  });

  if (!response.ok) {
    const errData = await response.json().catch(() => ({}));
    const errMsg = errData.error?.message || `HTTP エラー ${response.status}`;
    throw new Error(`OpenAI API エラー: ${errMsg}`);
  }

  const data = await response.json();

  // Responses API のレスポンス構造: output[0].content[0].text
  let text = '';
  if (data.output && Array.isArray(data.output)) {
    for (const item of data.output) {
      if (item.content && Array.isArray(item.content)) {
        for (const c of item.content) {
          if (c.text) text += c.text;
        }
      }
    }
  }
  // フォールバック: output_text フィールド
  if (!text && data.output_text) text = data.output_text;

  return text || '（結果を取得できませんでした）';
}

// 旧来モデル: Chat Completions API
async function callChatCompletionsAPI(apiKey, model, userMessage) {
  // o1 / o3 系は max_completion_tokens、それ以外は max_tokens
  const usesMaxCompletionTokens = /^(o1|o3|o4)/.test(model);
  const tokenParam = usesMaxCompletionTokens
    ? { max_completion_tokens: 4096 }
    : { max_tokens: 4096 };

  const requestBody = {
    model: model,
    messages: [
      { role: 'system', content: SYSTEM_PROMPT },
      { role: 'user',   content: userMessage },
    ],
    ...tokenParam,
  };

  // o1 / o3 系は temperature 非対応
  if (!usesMaxCompletionTokens) {
    requestBody.temperature = 0.2;
  }

  const response = await fetch('https://api.openai.com/v1/chat/completions', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${apiKey}`,
    },
    body: JSON.stringify(requestBody),
  });

  if (!response.ok) {
    const errData = await response.json().catch(() => ({}));
    const errMsg = errData.error?.message || `HTTP エラー ${response.status}`;
    throw new Error(`OpenAI API エラー: ${errMsg}`);
  }

  const data = await response.json();
  return data.choices?.[0]?.message?.content || '（結果を取得できませんでした）';
}

// ===== 結果表示 =====
function displayResults(rawText, documentText, model) {
  const charCount = documentText.length;
  const lineCount = documentText.split('\n').length;
  const now = new Date().toLocaleString('ja-JP');
  resultsMeta.innerHTML =
    `使用モデル: <strong>${model}</strong> ／ ` +
    `文字数: <strong>${charCount.toLocaleString()}</strong> 字 ／ ` +
    `行数: <strong>${lineCount}</strong> 行 ／ ` +
    `実行日時: ${now}`;

  resultsContent.innerHTML = markdownToHTML(rawText);
  resultsSection.style.display = 'block';
  resultsSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

// ===== 簡易 Markdown → HTML 変換 =====
function markdownToHTML(text) {
  text = text.replace(/【(.+?)→(.+?)】/g,
    '<span class="correction">【<del>$1</del> → $2】</span>');
  text = text.replace(/^### (.+)$/gm, '<div class="section-header">$1</div>');
  text = text.replace(/^## (.+)$/gm,  '<div class="section-header" style="font-size:14px;margin-top:8px;">$1</div>');
  text = text.replace(/^# (.+)$/gm,   '<div class="section-header" style="font-size:15px;margin-top:8px;">$1</div>');
  text = text.replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>');
  text = text.replace(/__(.+?)__/g,     '<strong>$1</strong>');
  text = text.replace(/^[-*] (.+)$/gm, '<div class="issue-item">• $1</div>');
  text = text.replace(/^\d+\. (.+)$/gm, '<div class="issue-item info">$1</div>');
  text = text.replace(/\n{2,}/g, '<br><br>');
  text = text.replace(/\n/g, '<br>');
  return text;
}

// ===== エラーフォーマット =====
function formatError(err) {
  const msg = err.message || '';
  if (msg.includes('401')) {
    return 'API キーが無効です。正しい OpenAI API キーを設定してください。';
  }
  if (msg.includes('429')) {
    return 'API のレート制限に達しました。しばらく待ってから再試行してください。';
  }
  if (msg.includes('insufficient_quota')) {
    return 'OpenAI API の利用枠が不足しています。OpenAI のダッシュボードで残高を確認してください。';
  }
  if (msg.includes('does not have access to model') || msg.includes('model_not_found')) {
    return `選択中のモデルはお使いのプロジェクトでは利用できません。\n↻ ボタンを押して利用可能なモデルを取得し、別のモデルを選択してください。\n\n詳細: ${msg}`;
  }
  return msg || '不明なエラーが発生しました。';
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
