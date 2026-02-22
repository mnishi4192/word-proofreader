/* =========================================================
   文書校正アシスタント - taskpane.js
   - 全モデルを Chat Completions API + stream:true で統一
   - ブロック分割なし。文書全体を1回のリクエストで送信
   - ストリーミングにより長文書でもタイムアウトしない
   ========================================================= */

'use strict';

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

// DOM 要素
let apiKeyInput, modelSelect, saveSettingsBtn, settingsSavedMsg;
let toggleKeyBtn, fetchModelsBtn, modelHint;
let proofreadBtn, btnText, btnSpinner;
let progressArea, progressText;
let resultsSection, resultsMeta, resultsContent;
let copyBtn, errorArea, errorMessage;

// ===== 初期化 =====
Office.onReady(function (info) {
  if (info.host === Office.HostType.Word) {
    initDOM();
    loadSettings();
    bindEvents();
    if ((localStorage.getItem('proofreader_api_key') || '').length > 10) {
      proofreadBtn.disabled = false;
    }
  }
});

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

function loadSettings() {
  const savedKey   = localStorage.getItem('proofreader_api_key') || '';
  const savedModel = localStorage.getItem('proofreader_model') || 'gpt-4o';
  apiKeyInput.value = savedKey;

  // 保存済みモデルが選択肢になければ追加
  if (!Array.from(modelSelect.options).find(o => o.value === savedModel)) {
    const opt = document.createElement('option');
    opt.value = savedModel;
    opt.textContent = savedModel + '（保存済み）';
    modelSelect.appendChild(opt);
  }
  modelSelect.value = savedModel;
}

function bindEvents() {
  toggleKeyBtn.addEventListener('click', () => {
    apiKeyInput.type = apiKeyInput.type === 'password' ? 'text' : 'password';
  });

  saveSettingsBtn.addEventListener('click', () => {
    const key = apiKeyInput.value.trim();
    localStorage.setItem('proofreader_api_key', key);
    localStorage.setItem('proofreader_model', modelSelect.value);
    settingsSavedMsg.style.display = 'inline';
    setTimeout(() => { settingsSavedMsg.style.display = 'none'; }, 2000);
    proofreadBtn.disabled = key.length <= 10;
  });

  apiKeyInput.addEventListener('input', () => {
    proofreadBtn.disabled = apiKeyInput.value.trim().length <= 10;
  });

  fetchModelsBtn.addEventListener('click', fetchAvailableModels);
  proofreadBtn.addEventListener('click', runProofread);

  copyBtn.addEventListener('click', () => {
    navigator.clipboard.writeText(resultsContent.innerText).then(() => {
      copyBtn.textContent = '✓ コピー済み';
      setTimeout(() => { copyBtn.textContent = 'コピー'; }, 2000);
    });
  });
}

// ===== モデル一覧取得 =====
async function fetchAvailableModels() {
  const apiKey = apiKeyInput.value.trim() || localStorage.getItem('proofreader_api_key') || '';
  if (apiKey.length < 10) {
    modelHint.textContent = '先に API キーを入力してください。';
    modelHint.style.color = '#c0392b';
    return;
  }
  fetchModelsBtn.disabled = true;
  fetchModelsBtn.textContent = '…';
  modelHint.textContent = 'モデルを取得中...';
  modelHint.style.color = '#666';
  try {
    const res = await fetch('https://api.openai.com/v1/models', {
      headers: { 'Authorization': `Bearer ${apiKey}` }
    });
    if (!res.ok) {
      const e = await res.json().catch(() => ({}));
      throw new Error(e.error?.message || `HTTP ${res.status}`);
    }
    const data = await res.json();
    const models = data.data
      .filter(m => m.id.startsWith('gpt-'))
      .map(m => m.id)
      .sort((a, b) => b.localeCompare(a));
    if (!models.length) throw new Error('利用可能な GPT モデルが見つかりませんでした。');

    const cur = modelSelect.value;
    modelSelect.innerHTML = '';
    const grp = document.createElement('optgroup');
    grp.label = `利用可能なモデル（${models.length} 件）`;
    models.forEach(id => {
      const opt = document.createElement('option');
      opt.value = id;
      opt.textContent = id;
      grp.appendChild(opt);
    });
    modelSelect.appendChild(grp);
    modelSelect.value = models.includes(cur) ? cur : (models.find(m => m === 'gpt-4o') || models[0]);

    modelHint.textContent = `✓ ${models.length} 件取得しました。`;
    modelHint.style.color = '#27ae60';
  } catch (err) {
    modelHint.textContent = `取得失敗: ${err.message}`;
    modelHint.style.color = '#c0392b';
  } finally {
    fetchModelsBtn.disabled = false;
    fetchModelsBtn.textContent = '↻';
  }
}

// ===== 校正実行 =====
async function runProofread() {
  const apiKey = localStorage.getItem('proofreader_api_key') || '';
  const model  = modelSelect.value || 'gpt-4o';

  if (apiKey.length < 10) {
    showError('API キーが設定されていません。設定欄に OpenAI API キーを入力して保存してください。');
    return;
  }

  setLoading(true);
  hideError();
  resultsSection.style.display = 'none';
  progressArea.style.display = 'block';
  setProgress('文書のテキストを取得中...');

  try {
    const docText = await getDocumentText();
    if (!docText || !docText.trim()) {
      throw new Error('文書にテキストが見つかりませんでした。');
    }

    setProgress(`${model} で校正中... お待ちください`);

    // ストリーミングで API 呼び出し
    const result = await callOpenAIStream(apiKey, model, docText);

    displayResults(result, docText, model);

  } catch (err) {
    showError(formatError(err));
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

// ===== ストリーミング API 呼び出し =====
// stream:true を使い、応答を少しずつ受信することでタイムアウトを回避する。
// ブロック分割は行わず、文書全体を1回のリクエストで送信する。
async function callOpenAIStream(apiKey, model, docText) {
  const userMsg = `以下の文書を校正してください。\n\n---\n${docText}\n---`;

  // o1/o3/o4 系は max_completion_tokens、それ以外は max_tokens
  const isReasoning = /^(o1|o3|o4)/.test(model);
  const body = {
    model,
    messages: [
      { role: 'system', content: SYSTEM_PROMPT },
      { role: 'user',   content: userMsg },
    ],
    stream: true,
    ...(isReasoning ? { max_completion_tokens: 16384 } : { max_tokens: 16384, temperature: 0.2 }),
  };

  const res = await fetch('https://api.openai.com/v1/chat/completions', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${apiKey}`,
    },
    body: JSON.stringify(body),
  });

  if (!res.ok) {
    const errData = await res.json().catch(() => ({}));
    throw new Error('OpenAI API エラー: ' + (errData.error?.message || `HTTP ${res.status}`));
  }

  // SSE ストリームを読み取る
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
        const delta = chunk.choices?.[0]?.delta?.content;
        if (delta) {
          fullText += delta;
          received += delta.length;
          setProgress(`${model} で校正中... （${received} 字受信済み）`);
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
  resultsMeta.innerHTML =
    `使用モデル: <strong>${model}</strong> ／ ` +
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
function formatError(err) {
  const m = err.message || '';
  if (m.includes('401'))                    return 'API キーが無効です。正しい OpenAI API キーを設定してください。';
  if (m.includes('429'))                    return 'API のレート制限に達しました。しばらく待ってから再試行してください。';
  if (m.includes('insufficient_quota'))     return 'OpenAI API の利用枠が不足しています。残高を確認してください。';
  if (m.includes('does not have access') || m.includes('model_not_found'))
    return `このモデルはご利用のプロジェクトで使用できません。↻ ボタンで利用可能なモデルを取得してください。\n\n詳細: ${m}`;
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
