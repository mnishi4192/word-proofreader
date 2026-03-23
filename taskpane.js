/* =========================================================
   文書校正アシスタント - taskpane.js (multi-service)
   対応サービス: OpenAI / Gemini / Claude
   - 各サービスの API キーはサービスごとに localStorage に保存
   - OpenAI: Chat Completions API + stream:true
   - Gemini: generateContent API (gemini-2.5-flash 固定可)
   - Claude: Messages API + stream:true
   ========================================================= */

'use strict';

// ===== 校正システムプロンプト =====
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

// ===== 状態管理 =====
let currentService = 'openai'; // 'openai' | 'gemini' | 'claude'

// ===== DOM 要素キャッシュ =====
let proofreadBtn, btnText, btnSpinner;
let progressArea, progressText;
let resultsSection, resultsMeta, resultsContent;
let copyBtn, errorArea, errorMessage;
let saveSettingsBtn, settingsSavedMsg;

// ===== 初期化 =====
Office.onReady(function (info) {
  if (info.host === Office.HostType.Word) {
    initDOM();
    loadSettings();
    bindEvents();
    updateProofreadBtnState();
  }
});

function initDOM() {
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
  saveSettingsBtn  = document.getElementById('save-settings');
  settingsSavedMsg = document.getElementById('settings-saved');
}

// ===== 設定の読み込み =====
function loadSettings() {
  // 最後に使ったサービスを復元
  const lastService = localStorage.getItem('proofreader_service') || 'openai';
  switchService(lastService);

  // 各サービスの API キーとモデルを復元
  ['openai', 'gemini', 'claude'].forEach(svc => {
    const keyEl   = document.getElementById(`apikey-${svc}`);
    const modelEl = document.getElementById(`model-${svc}`);
    if (keyEl)   keyEl.value   = localStorage.getItem(`proofreader_apikey_${svc}`) || '';
    if (modelEl) {
      const saved = localStorage.getItem(`proofreader_model_${svc}`);
      if (saved) {
        // 保存済みモデルが選択肢にない場合は追加する
        if (!Array.from(modelEl.options).find(o => o.value === saved)) {
          const opt = document.createElement('option');
          opt.value = saved;
          opt.textContent = saved + '（保存済み）';
          modelEl.appendChild(opt);
        }
        modelEl.value = saved;
      }
    }
  });
}

// ===== サービス切り替え =====
function switchService(svc) {
  currentService = svc;

  // タブのアクティブ状態を更新
  document.querySelectorAll('.service-tab').forEach(tab => {
    tab.classList.toggle('active', tab.dataset.service === svc);
  });

  // パネルの表示を切り替え
  ['openai', 'gemini', 'claude'].forEach(s => {
    const panel = document.getElementById(`panel-${s}`);
    if (panel) panel.style.display = s === svc ? 'block' : 'none';
  });

  updateProofreadBtnState();
}

// ===== 校正ボタンの有効/無効 =====
function updateProofreadBtnState() {
  const key = getActiveApiKey();
  proofreadBtn.disabled = !key || key.length < 10;
}

// ===== 現在のサービスの API キーを取得 =====
function getActiveApiKey() {
  return localStorage.getItem(`proofreader_apikey_${currentService}`) ||
         (document.getElementById(`apikey-${currentService}`)?.value.trim() || '');
}

// ===== 現在のサービスのモデルを取得 =====
function getActiveModel() {
  return document.getElementById(`model-${currentService}`)?.value || '';
}

// ===== イベントバインド =====
function bindEvents() {
  // サービスタブ切り替え
  document.querySelectorAll('.service-tab').forEach(tab => {
    tab.addEventListener('click', () => switchService(tab.dataset.service));
  });

  // API キー表示/非表示トグル
  document.querySelectorAll('.toggle-key-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const target = document.getElementById(btn.dataset.target);
      if (target) target.type = target.type === 'password' ? 'text' : 'password';
    });
  });

  // API キー入力時にボタン状態を更新
  ['openai', 'gemini', 'claude'].forEach(svc => {
    const keyEl = document.getElementById(`apikey-${svc}`);
    if (keyEl) keyEl.addEventListener('input', updateProofreadBtnState);
  });

  // 設定保存
  saveSettingsBtn.addEventListener('click', saveSettings);

  // モデル取得ボタン
  document.getElementById('fetch-models-openai').addEventListener('click', () => fetchModels('openai'));
  document.getElementById('fetch-models-gemini').addEventListener('click', () => fetchModels('gemini'));
  document.getElementById('fetch-models-claude').addEventListener('click', () => fetchModels('claude'));

  // 校正実行
  proofreadBtn.addEventListener('click', runProofread);

  // コピー
  copyBtn.addEventListener('click', () => {
    navigator.clipboard.writeText(resultsContent.innerText).then(() => {
      copyBtn.textContent = '✓ コピー済み';
      setTimeout(() => { copyBtn.textContent = 'コピー'; }, 2000);
    });
  });
}

// ===== 設定保存 =====
function saveSettings() {
  ['openai', 'gemini', 'claude'].forEach(svc => {
    const keyEl   = document.getElementById(`apikey-${svc}`);
    const modelEl = document.getElementById(`model-${svc}`);
    if (keyEl)   localStorage.setItem(`proofreader_apikey_${svc}`, keyEl.value.trim());
    if (modelEl) localStorage.setItem(`proofreader_model_${svc}`, modelEl.value);
  });
  localStorage.setItem('proofreader_service', currentService);
  settingsSavedMsg.style.display = 'inline';
  setTimeout(() => { settingsSavedMsg.style.display = 'none'; }, 2000);
  updateProofreadBtnState();
}

// ===== モデル一覧取得 =====
async function fetchModels(svc) {
  const keyEl    = document.getElementById(`apikey-${svc}`);
  const modelEl  = document.getElementById(`model-${svc}`);
  const hintEl   = document.getElementById(`hint-${svc}`);
  const fetchBtn = document.getElementById(`fetch-models-${svc}`);
  const apiKey   = keyEl?.value.trim() || localStorage.getItem(`proofreader_apikey_${svc}`) || '';

  if (apiKey.length < 10) {
    hintEl.textContent = '先に API キーを入力してください。';
    hintEl.style.color = '#c0392b';
    return;
  }

  fetchBtn.disabled = true;
  fetchBtn.textContent = '…';
  hintEl.textContent = 'モデルを取得中...';
  hintEl.style.color = '#666';

  try {
    let models = [];

    if (svc === 'openai') {
      const res = await fetch('https://api.openai.com/v1/models', {
        headers: { 'Authorization': `Bearer ${apiKey}` }
      });
      if (!res.ok) { const e = await res.json().catch(() => ({})); throw new Error(e.error?.message || `HTTP ${res.status}`); }
      const data = await res.json();
      models = data.data
        .filter(m => /^(gpt-|o1|o3|o4)/.test(m.id))
        .map(m => m.id)
        .sort((a, b) => b.localeCompare(a));

    } else if (svc === 'gemini') {
      const res = await fetch(`https://generativelanguage.googleapis.com/v1beta/models?key=${apiKey}`);
      if (!res.ok) { const e = await res.json().catch(() => ({})); throw new Error(e.error?.message || `HTTP ${res.status}`); }
      const data = await res.json();
      models = (data.models || [])
        .filter(m => m.name.includes('gemini') && m.supportedGenerationMethods?.includes('generateContent'))
        .map(m => m.name.replace('models/', ''))
        .sort((a, b) => b.localeCompare(a));

    } else if (svc === 'claude') {
      const res = await fetch('https://api.anthropic.com/v1/models', {
        headers: {
          'x-api-key': apiKey,
          'anthropic-version': '2023-06-01',
          'anthropic-dangerous-direct-browser-access': 'true',
        }
      });
      if (!res.ok) { const e = await res.json().catch(() => ({})); throw new Error(e.error?.message || `HTTP ${res.status}`); }
      const data = await res.json();
      models = (data.data || []).map(m => m.id).sort((a, b) => b.localeCompare(a));
    }

    if (!models.length) throw new Error('利用可能なモデルが見つかりませんでした。');

    const cur = modelEl.value;
    modelEl.innerHTML = '';
    const grp = document.createElement('optgroup');
    grp.label = `利用可能なモデル（${models.length} 件）`;
    models.forEach(id => {
      const opt = document.createElement('option');
      opt.value = id; opt.textContent = id;
      grp.appendChild(opt);
    });
    modelEl.appendChild(grp);
    modelEl.value = models.includes(cur) ? cur : models[0];

    hintEl.textContent = `✓ ${models.length} 件取得しました。`;
    hintEl.style.color = '#27ae60';
  } catch (err) {
    hintEl.textContent = `取得失敗: ${err.message}`;
    hintEl.style.color = '#c0392b';
  } finally {
    fetchBtn.disabled = false;
    fetchBtn.textContent = '↻';
  }
}

// ===== 校正実行 =====
async function runProofread() {
  const apiKey = getActiveApiKey();
  const model  = getActiveModel();

  if (!apiKey || apiKey.length < 10) {
    showError('API キーが設定されていません。設定欄に API キーを入力して保存してください。');
    return;
  }

  setLoading(true);
  hideError();
  resultsSection.style.display = 'none';
  progressArea.style.display = 'block';
  setProgress('文書のテキストを取得中...');

  try {
    const docText = await getDocumentText();
    if (!docText || !docText.trim()) throw new Error('文書にテキストが見つかりませんでした。');

    const serviceLabel = { openai: 'OpenAI', gemini: 'Gemini', claude: 'Claude' }[currentService];
    setProgress(`${serviceLabel} (${model}) で校正中... お待ちください`);

    let result = '';
    if (currentService === 'openai') {
      result = await callOpenAIStream(apiKey, model, docText);
    } else if (currentService === 'gemini') {
      result = await callGemini(apiKey, model, docText);
    } else if (currentService === 'claude') {
      result = await callClaudeStream(apiKey, model, docText);
    }

    displayResults(result, docText, model);

  } catch (err) {
    showError(formatError(err, currentService));
  } finally {
    setLoading(false);
    progressArea.style.display = 'none';
  }
}

// ===== Word 文書テキスト取得 =====
function getDocumentText() {
  return new Promise((resolve, reject) => {
    Word.run(async ctx => {
      try {
        const body = ctx.document.body;
        body.load('text');
        await ctx.sync();
        resolve(body.text);
      } catch (e) {
        reject(new Error('文書の読み取りに失敗しました: ' + e.message));
      }
    }).catch(reject);
  });
}

// ===== OpenAI: Chat Completions API (stream) =====
function usesMaxCompletionTokens(model) {
  return /^(o1|o3|o4|gpt-5)/.test(model);
}

async function callOpenAIStream(apiKey, model, docText) {
  const userMsg = `以下の文書を校正してください。\n\n---\n${docText}\n---`;
  const useCompletionTokens = usesMaxCompletionTokens(model);

  const requestBody = {
    model,
    messages: [
      { role: 'system', content: SYSTEM_PROMPT },
      { role: 'user',   content: userMsg },
    ],
    stream: true,
  };
  if (useCompletionTokens) {
    requestBody.max_completion_tokens = 16384;
  } else {
    requestBody.max_tokens = 16384;
    requestBody.temperature = 0.2;
  }

  const res = await fetch('https://api.openai.com/v1/chat/completions', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${apiKey}`,
    },
    body: JSON.stringify(requestBody),
  });

  if (!res.ok) {
    const errData = await res.json().catch(() => ({}));
    throw new Error('OpenAI API エラー: ' + (errData.error?.message || `HTTP ${res.status}`));
  }

  return readSSEStream(res, chunk => chunk.choices?.[0]?.delta?.content || '');
}

// ===== Gemini: generateContent API =====
async function callGemini(apiKey, model, docText) {
  const userMsg = `${SYSTEM_PROMPT}\n\n以下の文書を校正してください。\n\n---\n${docText}\n---`;

  // Gemini は 2.5 Flash を推奨。ストリーミングは streamGenerateContent を使用
  const endpoint = `https://generativelanguage.googleapis.com/v1beta/models/${model}:streamGenerateContent?alt=sse&key=${apiKey}`;

  const requestBody = {
    contents: [{ role: 'user', parts: [{ text: userMsg }] }],
    generationConfig: {
      maxOutputTokens: 16384,
      temperature: 0.2,
    },
  };

  const res = await fetch(endpoint, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(requestBody),
  });

  if (!res.ok) {
    const errData = await res.json().catch(() => ({}));
    throw new Error('Gemini API エラー: ' + (errData.error?.message || `HTTP ${res.status}`));
  }

  return readSSEStream(res, chunk => {
    // Gemini SSE レスポンス: candidates[0].content.parts[0].text
    return chunk.candidates?.[0]?.content?.parts?.[0]?.text || '';
  });
}

// ===== Claude: Messages API (stream) =====
async function callClaudeStream(apiKey, model, docText) {
  const userMsg = `以下の文書を校正してください。\n\n---\n${docText}\n---`;

  const requestBody = {
    model,
    max_tokens: 16384,
    system: SYSTEM_PROMPT,
    messages: [{ role: 'user', content: userMsg }],
    stream: true,
  };

  const res = await fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01',
      'anthropic-dangerous-direct-browser-access': 'true',
    },
    body: JSON.stringify(requestBody),
  });

  if (!res.ok) {
    const errData = await res.json().catch(() => ({}));
    throw new Error('Claude API エラー: ' + (errData.error?.message || `HTTP ${res.status}`));
  }

  // Claude SSE イベント: content_block_delta の delta.text
  return readSSEStream(res, chunk => {
    if (chunk.type === 'content_block_delta' && chunk.delta?.type === 'text_delta') {
      return chunk.delta.text || '';
    }
    return '';
  });
}

// ===== 共通 SSE ストリーム読み取り =====
async function readSSEStream(res, extractText) {
  const reader  = res.body.getReader();
  const decoder = new TextDecoder();
  let fullText = '';
  let buf = '';
  let received = 0;

  while (true) {
    const { done, value } = await reader.read();
    if (done) break;

    buf += decoder.decode(value, { stream: true });
    const lines = buf.split('\n');
    buf = lines.pop(); // 末尾の不完全行はバッファへ

    for (const line of lines) {
      const t = line.trim();
      if (!t || t === 'data: [DONE]') continue;
      if (!t.startsWith('data: ')) continue;
      try {
        const chunk = JSON.parse(t.slice(6));
        const delta = extractText(chunk);
        if (delta) {
          fullText += delta;
          received += delta.length;
          setProgress(`校正中... （${received} 字受信済み）`);
        }
      } catch (_) { /* 不完全チャンクは無視 */ }
    }
  }

  if (!fullText.trim()) {
    throw new Error('API からの応答が空でした。モデルを変更するか、しばらく待ってから再試行してください。');
  }
  return fullText;
}

// ===== 結果表示 =====
function displayResults(rawText, docText, model) {
  const now = new Date().toLocaleString('ja-JP');
  const svcLabel = { openai: 'OpenAI', gemini: 'Gemini', claude: 'Claude' }[currentService];
  const badgeClass = `badge badge-${currentService}`;

  resultsMeta.innerHTML =
    `<span class="${badgeClass}">${svcLabel}</span>` +
    `モデル: <strong>${model}</strong> ／ ` +
    `文字数: <strong>${docText.length.toLocaleString()}</strong> 字 ／ ` +
    `行数: <strong>${docText.split('\n').length}</strong> 行 ／ ` +
    `実行日時: ${now}`;

  resultsContent.innerHTML = md2html(rawText);
  resultsSection.style.display = 'block';
  resultsSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

// ===== Markdown → HTML =====
function md2html(t) {
  t = t.replace(/【(.+?)→(.+?)】/g, '<span class="correction">【<del>$1</del> → $2】</span>');
  t = t.replace(/^### (.+)$/gm, '<div class="section-header">$1</div>');
  t = t.replace(/^## (.+)$/gm,  '<div class="section-header" style="font-size:14px;margin-top:8px;">$1</div>');
  t = t.replace(/^# (.+)$/gm,   '<div class="section-header" style="font-size:15px;margin-top:8px;">$1</div>');
  t = t.replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>');
  t = t.replace(/^[-*] (.+)$/gm, '<div class="issue-item">• $1</div>');
  t = t.replace(/^\d+\. (.+)$/gm, '<div class="issue-item info">$1</div>');
  t = t.replace(/\n{2,}/g, '<br><br>');
  t = t.replace(/\n/g, '<br>');
  return t;
}

// ===== エラーフォーマット =====
function formatError(err, svc) {
  const m = err.message || '';
  if (m.includes('401') || m.includes('invalid_api_key') || m.includes('API_KEY_INVALID')) {
    return `API キーが無効です。正しい ${svc === 'openai' ? 'OpenAI' : svc === 'gemini' ? 'Gemini' : 'Claude'} API キーを設定してください。`;
  }
  if (m.includes('429')) return 'API のレート制限に達しました。しばらく待ってから再試行してください。';
  if (m.includes('insufficient_quota')) return 'API の利用枠が不足しています。残高を確認してください。';
  if (m.includes('does not have access') || m.includes('model_not_found')) {
    return `このモデルはご利用のプランで使用できません。↻ ボタンで利用可能なモデルを取得してください。\n\n詳細: ${m}`;
  }
  if (m.includes('max_tokens')) {
    return `トークンパラメータのエラーです。↻ ボタンでモデル一覧を再取得し、別のモデルを試してください。\n\n詳細: ${m}`;
  }
  return m || '不明なエラーが発生しました。';
}

// ===== UI ヘルパー =====
function setLoading(on) {
  proofreadBtn.disabled    = on;
  btnText.textContent      = on ? '校正中...' : '文書を校正する';
  btnSpinner.style.display = on ? 'inline-block' : 'none';
}
function setProgress(msg) { progressText.textContent = msg; }
function showError(msg)    { errorMessage.textContent = msg; errorArea.style.display = 'block'; }
function hideError()       { errorArea.style.display = 'none'; }
