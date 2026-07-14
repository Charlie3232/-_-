const SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbyG79Hu01QQv2Bieo5yaaydb5TTWDUOWqoi7ieQwP87PvlXef5JhQ-GYAB8Eta7f-0u/exec';
let allIssues = [];
let dataConfig = {};
let isMutating = false;
let userList = [];
let currentUser = { id: "", name: "", role: "" };
let currentModalType = 'TS';

// 重工資料結構
// kwData[catKey] = { subCols: ['廠房類型','空間性質',...], data: { '廠房類型': [...values] } }
let kwData = {};

// 重工分頁：已選關鍵字
// selectedKw[catKey] = Set of "小項::值"
let selectedKw = { industry: new Set(), system: new Set(), hardware: new Set(), spec: new Set() };

// 重工資料庫分頁：各大類當前顯示的小項
let kwActiveTab = { industry: null, system: null, hardware: null, spec: null };

// Modal 關鍵詞彙：各大類當前選中小項
let kwModalActiveTab = { industry: null, system: null, hardware: null, spec: null };
// Modal 關鍵詞彙：已選值 ( catKey -> "小項::值" | null )
let kwModalSelected = { industry: null, system: null, hardware: null, spec: null };

/* ══ 大類定義 ══ */
const KW_CATS = {
  industry: { label: '產業範疇', cols: ['廠房類型','空間性質','環境等級'] },
  system:   { label: '系統類別', cols: ['水系統','氣體系統','化學品系統','電力系統','空調系統'] },
  hardware: { label: '硬件組成', cols: ['管路與材料','機構組成','客製化開發'] },
  spec:     { label: '技術規範', cols: ['工程文件','測試驗證','規範標準'] }
};

const getToday2026 = () => {
  const d = new Date();
  return `2026-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
};

/* ════════════════════════════════════════
   Toast 通知
   ════════════════════════════════════════ */
function showToast(msg, duration = 3000) {
  const toast = document.getElementById('validation-toast');
  toast.textContent = msg;
  toast.classList.toggle('tag-toast', msg.includes('[ TAG REGISTERED ]'));
  toast.style.display = 'block';
  clearTimeout(toast._timer);
  toast._timer = setTimeout(() => {
    toast.style.display = 'none';
    toast.classList.remove('tag-toast');
  }, duration);
}

function escapeHtml(value) {
  return (value || '').toString()
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function getIssueKeywordParts(issue) {
  return ['kw_industry', 'kw_system', 'kw_hardware', 'kw_spec']
    .flatMap(field => String(issue[field] || '').split('||'))
    .map(item => item.trim())
    .filter(Boolean)
    .map(item => {
      const parts = item.split('::');
      return {
        column: parts[0] || '',
        value: parts.slice(1).join('::') || parts[0] || ''
      };
    });
}

function renderIssueKeywordChips(issue) {
  const parts = getIssueKeywordParts(issue).slice(0, 4);
  if (!parts.length) return '';
  return `<div class="issue-kw-chips">${parts.map(p =>
    `<span class="issue-kw-chip" title="${escapeHtml(p.column)}">${escapeHtml(p.value)}</span>`
  ).join('')}</div>`;
}

function renderIssueCategorySummary(issue) {
  const parts = getIssueKeywordParts(issue);
  if (!parts.length) return '';
  const summary = parts.slice(0, 3).map(p => p.value).join(' / ');
  return `<div class="issue-kw-summary">分類：${escapeHtml(summary)}</div>`;
}

const DRAFT_KEY = 'issueFormDraftV1';

function isModalOpen() {
  const overlay = document.getElementById('modal-overlay');
  return overlay && overlay.style.display === 'flex';
}

function collectFormDraft() {
  return {
    modalType: window.currentModalType || 'TS',
    issue: document.getElementById('input-issue')?.value || '',
    owner: document.getElementById('input-owner')?.value || '',
    status: document.getElementById('input-status')?.value || '',
    customer: document.getElementById('input-customer')?.value || '',
    customerText: document.getElementById('ss-customer-input')?.value || '',
    product: document.getElementById('input-product')?.value || '',
    productText: document.getElementById('ss-product-input')?.value || '',
    project: document.getElementById('input-project')?.value || '',
    projectText: document.getElementById('ss-project-input')?.value || '',
    deadline: document.getElementById('input-deadline')?.value || '',
    priority: document.getElementById('input-priority')?.value || '',
    description: document.getElementById('input-description')?.value || '',
    links: Array.from(document.querySelectorAll('#link-group .link-entry')).map(el => el.value),
    records: Array.from(document.querySelectorAll('.record-item-row')).map(row => ({
      checked: !!row.querySelector('.record-chk')?.checked,
      title: row.querySelector('.record-title-input')?.value || '',
      content: row.querySelector('.record-content-input')?.value || ''
    })),
    keywords: { ...kwModalSelected },
    savedAt: Date.now()
  };
}

function saveFormDraft() {
  if (!isModalOpen() || document.getElementById('edit-id')?.value) return;
  localStorage.setItem(DRAFT_KEY, JSON.stringify(collectFormDraft()));
}

function clearFormDraft() {
  localStorage.removeItem(DRAFT_KEY);
}

function hasDraftContent(draft) {
  if (!draft) return false;
  return [
    draft.issue, draft.customerText, draft.productText, draft.projectText,
    draft.description, draft.deadline, draft.priority
  ].some(v => normalizeDictValue(v)) ||
    (draft.records || []).some(r => normalizeDictValue(r.title) || normalizeDictValue(r.content)) ||
    (draft.links || []).some(normalizeDictValue) ||
    Object.values(draft.keywords || {}).some(Boolean);
}

function tryRestoreFormDraft() {
  const raw = localStorage.getItem(DRAFT_KEY);
  if (!raw) return;
  let draft = null;
  try { draft = JSON.parse(raw); } catch(e) { clearFormDraft(); return; }
  if (!hasDraftContent(draft)) return;
  if (!confirm("偵測到未完成草稿，是否恢復？")) return;

  document.getElementById('input-issue').value = draft.issue || '';
  document.getElementById('input-owner').value = draft.owner || '';
  document.getElementById('input-status').value = draft.status || '';
  loadSearchableSelectValue('ss-customer', draft.customerText || draft.customer || '');
  loadSearchableSelectValue('ss-product', draft.productText || draft.product || '');
  loadSearchableSelectValue('ss-project', draft.projectText || draft.project || '');
  document.getElementById('input-deadline').value = draft.deadline || '';
  document.getElementById('input-priority').value = draft.priority || '';
  document.getElementById('input-description').value = draft.description || '';

  document.getElementById('records-container').innerHTML = '';
  (draft.records || []).forEach(r => addRecordItem(`${r.title || ''}::${r.content || ''}`, !!r.checked));
  if (!document.querySelector('.record-item-row')) addRecordItem();

  document.getElementById('link-group').innerHTML = '';
  (draft.links || []).filter(v => v !== undefined).forEach(url => addLinkItem(url));
  if (!document.querySelector('#link-group .link-entry')) addLinkItem();

  initModalKeywords(draft.keywords || null);
  showToast("草稿已恢復", 2500);
}

document.addEventListener('input', saveFormDraft);
document.addEventListener('change', saveFormDraft);

/* ════════════════════════════════════════
   一般 select 填充
   ════════════════════════════════════════ */
function fillFormSelect(id, list) {
  const el = document.getElementById(id);
  if (el && dataConfig[list]) {
    el.innerHTML = dataConfig[list].map(t => `<option value="${t}">${t}</option>`).join('');
    el.insertAdjacentHTML('afterbegin', '<option value="" disabled selected>請選擇...</option>');
  }
}

/* ════════════════════════════════════════
   可搜尋下拉元件（Modal 用）
   ════════════════════════════════════════ */
function initSearchableSelect(wrapperId, listKey) {
  const dropdown = document.getElementById(`${wrapperId}-dropdown`);
  if (!dropdown || !dataConfig[listKey]) return;
  dropdown.dataset.list = listKey;
  renderSearchableOptions(wrapperId, dataConfig[listKey]);
}

function renderSearchableOptions(wrapperId, options) {
  const dropdown = document.getElementById(`${wrapperId}-dropdown`);
  if (!dropdown) return;
  dropdown.innerHTML = options.map(opt =>
    `<div class="ss-option" onclick="selectSearchableOption('${wrapperId}', '${opt}', event)">${opt}</div>`
  ).join('');
}

function filterSearchableSelect(wrapperId) {
  const input = document.getElementById(`${wrapperId}-input`);
  const hiddenInput = document.getElementById(`input-${wrapperId.replace('ss-', '')}`);
  if (hiddenInput) hiddenInput.value = input.value.trim();
  const dropdown = document.getElementById(`${wrapperId}-dropdown`);
  const listKey = dropdown.dataset.list;
  if (!listKey || !dataConfig[listKey]) return;
  const query = input.value.toLowerCase();
  const filtered = dataConfig[listKey].filter(opt => opt.toLowerCase().includes(query));
  renderSearchableOptions(wrapperId, filtered);
  const typed = input.value.trim();
  const exists = dataConfig[listKey].some(opt => opt.toLowerCase() === typed.toLowerCase());
  if (typed && !exists) {
    dropdown.insertAdjacentHTML('beforeend',
      `<div class="ss-create-option" onclick="selectSearchableOption('${wrapperId}', '${typed.replace(/'/g,"\\'")}', event)">＋ 使用並新增「${typed}」</div>`
    );
  }
  dropdown.style.display = 'block';
}

function openSearchableSelect(wrapperId) {
  document.querySelectorAll('.ss-dropdown').forEach(d => {
    if (d.id !== `${wrapperId}-dropdown`) d.style.display = 'none';
  });
  const dropdown = document.getElementById(`${wrapperId}-dropdown`);
  if (dropdown) dropdown.style.display = 'block';
}

function selectSearchableOption(wrapperId, value, event) {
  event.stopPropagation();
  document.getElementById(`${wrapperId}-input`).value = value;
  document.getElementById(`input-${wrapperId.replace('ss-', '')}`).value = value;
  document.getElementById(`${wrapperId}-dropdown`).style.display = 'none';
  // 清除錯誤狀態
  const wrapper = document.getElementById(wrapperId);
  if (wrapper) wrapper.closest('.field-item')?.classList.remove('field-error');
}

function loadSearchableSelectValue(wrapperId, value) {
  const textInput = document.getElementById(`${wrapperId}-input`);
  const hiddenInput = document.getElementById(`input-${wrapperId.replace('ss-', '')}`);
  if (textInput) textInput.value = value || '';
  if (hiddenInput) hiddenInput.value = value || '';
}

/* ════════════════════════════════════════
   登入
   ════════════════════════════════════════ */
async function handleLogin() {
  const idInput = document.getElementById('login-user').value.trim();
  const pwdInput = document.getElementById('login-pwd').value.trim();
  if (!idInput || !pwdInput) { alert("請輸入 ID 與密碼"); return; }
  if (userList.length === 0) await fetchDataOnLoad();
  const user = userList.find(u => u.id === idInput && u.pwd === pwdInput);
  if (user) {
    currentUser = { id: user.id, name: user.name, role: (user.id === "G0006" ? "MANAGER" : "USER") };
    document.getElementById('current-username-display').innerText = currentUser.name;
    document.getElementById('login-overlay').style.display = 'none';
    document.getElementById('main-ui').style.display = 'block';
    if (currentUser.role === "MANAGER") {
      document.getElementById('btn-tab-main').style.display = 'block';
      document.getElementById('btn-tab-manager').style.display = 'block';
    }
    initUI();
  } else { alert("驗證失敗：人員識別碼或密碼錯誤"); }
}

async function fetchDataOnLoad() {
  try {
    const resp = await fetch(SCRIPT_URL + '?action=getData');
    const data = await resp.json();
    allIssues = data.issues || [];
    dataConfig = data.config || {};
    userList = data.users || [];
    kwData = data.kwData || {};
    if (document.getElementById('login-status')) document.getElementById('login-status').innerText = "系統就緒。";
  } catch(e) { console.error("Initial load failed", e); }
}
window.onload = fetchDataOnLoad;

function initUI() {
  fillUIConfigs();
  renderIssues();
  renderManagerIssues();
  renderStats();
  initKeywordDB();
  setInterval(silentSync, 15000);
  updateHeaderGif('tab-issues');
}

async function silentSync() {
  if (isMutating || document.getElementById('main-ui').style.display === 'none') return;
  try {
    const resp = await fetch(SCRIPT_URL + '?action=getData');
    const data = await resp.json();
    if (isMutating) return;
    allIssues = data.issues || [];
    kwData = data.kwData || {};
    renderIssues(); renderManagerIssues(); renderStats();
    renderKeywordResults();
  } catch(e) {}
}

/* ════════════════════════════════════════
   GIF 顯示控制
   ════════════════════════════════════════ */
function updateHeaderGif(tabId) {
  const gifIssues  = document.getElementById('header-gif-issues');
  const gifManager = document.getElementById('header-gif-manager');
  if (gifIssues)  gifIssues.style.display  = (tabId === 'tab-issues')  ? 'flex' : 'none';
  if (gifManager) gifManager.style.display = (tabId === 'tab-manager') ? 'flex' : 'none';
}

/* ════════════════════════════════════════
   篩選下拉搜尋
   ════════════════════════════════════════ */
function filterDropdownSearch(inputEl, dropdownId) {
  const query = inputEl.value.toLowerCase();
  const listEl = document.getElementById(`${dropdownId}-list`);
  if (!listEl) return;
  Array.from(listEl.querySelectorAll('.checkbox-label')).forEach(label => {
    label.style.display = label.textContent.toLowerCase().includes(query) ? '' : 'none';
  });
}

function selectAllFilter(dropdownId, onChangeCode, event) {
  event.stopPropagation();
  const listEl = document.getElementById(`${dropdownId}-list`);
  if (!listEl) return;
  listEl.querySelectorAll('input[type="checkbox"]').forEach(cb => cb.checked = true);
  eval(onChangeCode);
}

function clearAllFilter(dropdownId, onChangeCode, event) {
  event.stopPropagation();
  const listEl = document.getElementById(`${dropdownId}-list`);
  if (!listEl) return;
  listEl.querySelectorAll('input[type="checkbox"]').forEach(cb => cb.checked = false);
  const parent = document.getElementById(dropdownId);
  if (parent) {
    const searchInput = parent.querySelector('.filter-search-input');
    if (searchInput) { searchInput.value = ''; filterDropdownSearch(searchInput, dropdownId); }
  }
  eval(onChangeCode);
}

const fillCheckboxes = (dropdownId, listKey, onChangeCode) => {
  const listEl = document.getElementById(`${dropdownId}-list`);
  if (listEl && dataConfig[listKey]) {
    listEl.innerHTML = dataConfig[listKey].map(t =>
      `<label class="checkbox-label" onclick="event.stopPropagation()">
        <input type="checkbox" value="${t}" onchange="${onChangeCode}"> ${t}
      </label>`
    ).join('');
  }
};

function fillUIConfigs() {
  fillCheckboxes('items-owner',    'owners',     'renderIssues()');
  fillCheckboxes('items-status',   'statusList', 'renderIssues()');
  fillCheckboxes('items-product',  'products',   'renderIssues()');
  fillCheckboxes('items-customer', 'customers',  'renderIssues()');
  fillCheckboxes('mgr-owner',    'owners',     'renderManagerIssues()');
  fillCheckboxes('mgr-status',   'statusList', 'renderManagerIssues()');
  fillCheckboxes('mgr-product',  'products',   'renderManagerIssues()');
  fillCheckboxes('mgr-customer', 'customers',  'renderManagerIssues()');
  fillFormSelect('input-owner',  'owners');
  fillFormSelect('input-status', 'statusList');
  initSearchableSelect('ss-customer', 'customers');
  initSearchableSelect('ss-product',  'products');
  initSearchableSelect('ss-project',  'projects');
}

const getCheckedValues = (id) => {
  const listEl = document.getElementById(`${id}-list`);
  const source = listEl || document.getElementById(id);
  if (!source) return [];
  return Array.from(source.querySelectorAll('input[type="checkbox"]:checked')).map(cb => cb.value);
};

const isTaskUrgent = (deadlineStr, status) => {
  if (!deadlineStr || status === "已解決" || status === "Done") return false;
  const today = new Date(); today.setHours(0, 0, 0, 0);
  let dParts = String(deadlineStr).split(/[-/T ]/);
  if (dParts.length < 3) return false;
  const deadline = new Date(dParts[0], dParts[1] - 1, dParts[2]);
  deadline.setHours(0, 0, 0, 0);
  return Math.ceil((deadline.getTime() - today.getTime()) / (1000 * 60 * 60 * 24)) <= 2;
};

/* ════════════════════════════════════════
   統計圖（無變動）
   ════════════════════════════════════════ */
function renderStats() {
  const ownerStart = document.getElementById('stats-owner-start').value;
  const ownerEnd   = document.getElementById('stats-owner-end').value;
  const prodStart  = document.getElementById('stats-prod-start').value;
  const prodEnd    = document.getElementById('stats-prod-end').value;
  const colors = ['#0f0', '#ffeb3b', '#ff0055', '#a020f0', '#ff9800', '#00bcd4'];

  const ownerCounts = {}; let ownerTotal = 0;
  allIssues.filter(i => {
    if (i.id && String(i.id).startsWith('MGR-')) return false;
    let iDate = i.date ? i.date.replace(/\//g, '-') : "";
    if (ownerStart && iDate < ownerStart) return false;
    if (ownerEnd && iDate > ownerEnd) return false;
    return true;
  }).forEach(i => { ownerCounts[i.owner] = (ownerCounts[i.owner] || 0) + 1; ownerTotal++; });

  const ownerBars = document.getElementById('stats-bars');
  if (ownerBars) {
    ownerBars.innerHTML = Object.keys(ownerCounts).sort((a, b) => ownerCounts[b] - ownerCounts[a]).map((o, idx) => {
      const pct = ownerTotal ? Math.round(ownerCounts[o] / ownerTotal * 100) : 0;
      return `<div style="margin-bottom:12px; display:flex; align-items:center; gap:10px;">
                <div style="width:110px; font-size:12px; color:var(--pixel-green);">${o}</div>
                <div style="flex:1; background:#000; border:1px solid #285428; height:26px; position:relative;">
                  <div style="width:${pct}%; background:${colors[idx % colors.length]}; height:100%;"></div>
                  <div style="position:absolute; right:8px; top:5px; font-size:10px; color:#fff;">${ownerCounts[o]}件 (${pct}%)</div>
                </div>
              </div>`;
    }).join('') || "無人員數據";
  }

  const prodCounts = {}; let prodTotal = 0;
  allIssues.filter(i => {
    if (i.id && String(i.id).startsWith('MGR-')) return false;
    let iDate = i.date ? i.date.replace(/\//g, '-') : "";
    if (prodStart && iDate < prodStart) return false;
    if (prodEnd && iDate > prodEnd) return false;
    return true;
  }).forEach(i => { prodCounts[i.product] = (prodCounts[i.product] || 0) + 1; prodTotal++; });

  const prodBars = document.getElementById('product-stats-bars');
  if (prodBars) {
    prodBars.innerHTML = Object.keys(prodCounts).sort((a, b) => prodCounts[b] - prodCounts[a]).map((p, idx) => {
      const pct = prodTotal ? Math.round(prodCounts[p] / prodTotal * 100) : 0;
      return `<div style="margin-bottom:12px; display:flex; align-items:center; gap:10px;">
                <div style="width:110px; font-size:12px; color:var(--pixel-green);">${p}</div>
                <div style="flex:1; background:#000; border:1px solid #285428; height:26px; position:relative;">
                  <div style="width:${pct}%; background:${colors[(idx + 2) % colors.length]}; height:100%;"></div>
                  <div style="position:absolute; right:8px; top:5px; font-size:10px; color:#fff;">${prodCounts[p]}件 (${pct}%)</div>
                </div>
              </div>`;
    }).join('') || "無產品數據";
  }
}

function normalizeDictValue(value) {
  return (value || '').toString().trim();
}

function listHasValue(list, value) {
  const target = normalizeDictValue(value).toLowerCase();
  return (list || []).some(item => normalizeDictValue(item).toLowerCase() === target);
}

function compactCompareValue(value) {
  return normalizeDictValue(value).toLowerCase().replace(/[\s　_\-\/\\｜|:：,，.。()（）\[\]【】]/g, '');
}

function findSimilarValue(list, value) {
  const target = compactCompareValue(value);
  if (!target || target.length < 2) return '';
  return (list || []).find(item => {
    const source = compactCompareValue(item);
    if (!source || source === target) return false;
    return source.includes(target) || target.includes(source);
  }) || '';
}

function confirmSimilarDictionaryValue(list, value, label) {
  const similar = findSimilarValue(list, value);
  if (!similar) return true;
  return confirm(`資料庫中已有相似${label}：「${similar}」\n仍要新增「${value}」嗎？`);
}

function addLocalDictionaryValue(type, value, columnName) {
  const clean = normalizeDictValue(value);
  if (!clean) return;
  if (type === 'keyword') {
    if (!kwData[columnName]) kwData[columnName] = [];
    if (!listHasValue(kwData[columnName], clean)) kwData[columnName].push(clean);
    return;
  }
  const configKeyMap = { customer: 'customers', product: 'products', project: 'projects' };
  const configKey = configKeyMap[type];
  if (!configKey) return;
  if (!dataConfig[configKey]) dataConfig[configKey] = [];
  if (!listHasValue(dataConfig[configKey], clean)) dataConfig[configKey].push(clean);
}

async function saveDictionaryItem(type, value, columnName = '') {
  const clean = normalizeDictValue(value);
  if (!clean) return;
  await fetch(SCRIPT_URL, {
    method: 'POST',
    mode: 'no-cors',
    body: JSON.stringify({
      action: 'addDictionaryItem',
      type,
      value: clean,
      columnName
    })
  });
}

async function saveNewModalDictionaryValues() {
  const entries = [
    { type: 'customer', configKey: 'customers', value: document.getElementById('input-customer').value },
    { type: 'product',  configKey: 'products',  value: document.getElementById('input-product').value },
    { type: 'project',  configKey: 'projects',  value: document.getElementById('input-project').value }
  ];
  for (const item of entries) {
    const clean = normalizeDictValue(item.value);
    if (clean && !listHasValue(dataConfig[item.configKey], clean)) {
      if (!confirmSimilarDictionaryValue(dataConfig[item.configKey], clean, item.type === 'customer' ? '客戶別' : item.type === 'product' ? '產品別' : '專案別')) {
        throw new Error('使用者取消新增相似字典項目');
      }
      addLocalDictionaryValue(item.type, clean);
      await saveDictionaryItem(item.type, clean);
    }
  }
}

/* ════════════════════════════════════════
   renderIssues（無變動）
   ════════════════════════════════════════ */
function renderIssues() {
  const container = document.getElementById('issue-display');
  const search  = document.getElementById('search-input').value.toLowerCase();
  const fOwners = getCheckedValues('items-owner');
  const fStats  = getCheckedValues('items-status');
  const fProds  = getCheckedValues('items-product');
  const fCusts  = getCheckedValues('items-customer');

  let filtered = allIssues.filter(i =>
    (!i.id || !String(i.id).startsWith('MGR-')) &&
    String(i.issue).toLowerCase().includes(search) &&
    (fOwners.length === 0 || fOwners.includes(i.owner)) &&
    (fStats.length === 0 ? (i.status !== "已解決" && i.status !== "Done") : fStats.includes(i.status)) &&
    (fProds.length === 0 || fProds.includes(i.product)) &&
    (fCusts.length === 0 || fCusts.includes(i.customer))
  ).sort((a, b) => {
    const isDoneA = (a.status === "已解決" || a.status === "Done");
    const isDoneB = (b.status === "已解決" || b.status === "Done");
    if (isDoneA !== isDoneB) return isDoneA ? 1 : -1;
    const urgentA = isTaskUrgent(a.deadline, a.status);
    const urgentB = isTaskUrgent(b.deadline, b.status);
    if (urgentA !== urgentB) return urgentA ? -1 : 1;
    return new Date(b.date) - new Date(a.date);
  });

  container.innerHTML = filtered.map(i => {
    const stat = String(i.status);
    const isDone = (stat === "已解決" || stat === "Done");
    const isUrgent = (!isDone && isTaskUrgent(i.deadline, stat));
    return `<div class="pebble ${isDone ? 'resolved-card' : ''} ${isUrgent ? 'urgent-card' : ''}" onclick="openEdit('${i.id}')">
      <div style="font-size:11px; color:${isUrgent ? '#ff0055' : 'var(--pixel-green)'};">[ ${escapeHtml(stat)} ]</div>
      <div style="font-size:20px; margin:10px 0; line-height:1.3;">${escapeHtml(i.issue)}</div>
      <div style="font-size:12px; opacity:0.6;">${escapeHtml(i.product)} | ${escapeHtml(i.owner)}</div>
      ${renderIssueKeywordChips(i)}
      ${renderIssueCategorySummary(i)}
    </div>`;
  }).join('');
}

/* ════════════════════════════════════════
   renderManagerIssues（無變動）
   ════════════════════════════════════════ */
function renderManagerIssues() {
  const container = document.getElementById('manager-issue-display');
  const search  = document.getElementById('search-input-mgr').value.toLowerCase();
  const fOwners = getCheckedValues('mgr-owner');
  const fStats  = getCheckedValues('mgr-status');
  const fProds  = getCheckedValues('mgr-product');
  const fCusts  = getCheckedValues('mgr-customer');

  let filtered = allIssues.filter(i =>
    i.id && String(i.id).startsWith('MGR-') &&
    String(i.issue).toLowerCase().includes(search) &&
    (fOwners.length === 0 || fOwners.includes(i.owner)) &&
    (fStats.length === 0 ? (i.status !== "已解決") : fStats.includes(i.status)) &&
    (fProds.length === 0 || fProds.includes(i.product)) &&
    (fCusts.length === 0 || fCusts.includes(i.customer))
  ).sort((a, b) => new Date(b.date) - new Date(a.date));

  container.innerHTML = filtered.map(i =>
    `<div class="pebble" onclick="openEdit('${i.id}')">
      <div style="font-size:11px; color:#ffb7c5; margin-bottom:8px;">[ ${i.status} ]</div>
      <div style="font-size:20px; margin:10px 0; color:var(--mgr-pink);">${i.issue}</div>
      <div style="font-size:12px; opacity:0.6;">${i.owner} | ${i.product}</div>
    </div>`
  ).join('');
}

/* ════════════════════════════════════════
   switchTab
   ════════════════════════════════════════ */
function switchTab(tabId) {
  document.querySelectorAll('.tab-section').forEach(el => {
    el.style.display = 'none';
    el.classList.remove('tab-active');
  });
  document.querySelectorAll('.nav-btn').forEach(btn => btn.classList.remove('active'));
  document.getElementById(tabId).style.display = 'block';
  document.getElementById(tabId).classList.add('tab-active');
  document.getElementById('btn-' + tabId).classList.add('active');
  updateHeaderGif(tabId);
  if (tabId === 'tab-main') renderStats();
  if (tabId === 'tab-keyword') renderKeywordResults();
}

function toggleDropdown(id, event) {
  event.stopPropagation();
  document.querySelectorAll('.select-items').forEach(el => { if (el.id !== id) el.style.display = 'none'; });
  const el = document.getElementById(id);
  el.style.display = el.style.display === 'block' ? 'none' : 'block';
}

document.addEventListener('click', () => {
  document.querySelectorAll('.select-items').forEach(el => el.style.display = 'none');
  document.querySelectorAll('.ss-dropdown').forEach(el => el.style.display = 'none');
});

/* ════════════════════════════════════════
   ★ 重工資料庫 - 初始化
   ════════════════════════════════════════ */
function initKeywordDB() {
  Object.entries(KW_CATS).forEach(([catKey, catDef]) => {
    // 建小項 tabs
    const tabsEl = document.getElementById(`kwtabs-${catKey}`);
    if (!tabsEl) return;
    tabsEl.innerHTML = catDef.cols.map(col =>
      `<div class="kw-subcat-tab" id="kwtab-${catKey}-${col}" onclick="setKwTab('${catKey}','${col}')">${col}</div>`
    ).join('');
    // 預設選第一個小項
    if (catDef.cols.length > 0) setKwTab(catKey, catDef.cols[0]);
  });
  renderSelectedBar();
}

function toggleKwCategory(catKey) {
  const body = document.getElementById(`kwbody-${catKey}`);
  const header = document.querySelector(`#kwcat-${catKey} .kw-cat-header`);
  if (!body) return;
  const isOpen = body.style.display !== 'none';
  body.style.display = isOpen ? 'none' : 'block';
  const arrow = isOpen ? '▶' : '▼';
  // update arrow in header text
  const badge = header.querySelector('.kw-cat-badge');
  header.childNodes[0].textContent = `${arrow} ${KW_CATS[catKey].label} `;
}

function setKwTab(catKey, colName) {
  kwActiveTab[catKey] = colName;
  // update tab active state
  KW_CATS[catKey].cols.forEach(col => {
    const tab = document.getElementById(`kwtab-${catKey}-${col}`);
    if (!tab) return;
    tab.className = 'kw-subcat-tab';
    if (col === colName) tab.classList.add(`active-${catKey}`);
  });
  renderKwPool(catKey);
}

function renderKwPool(catKey) {
  const colName = kwActiveTab[catKey];
  const poolEl = document.getElementById(`kwpool-${catKey}`);
  if (!poolEl || !colName) return;
  const values = (kwData[colName] || []).filter(v => v && v.toString().trim());
  poolEl.innerHTML = values.map(val => {
    const key = `${colName}::${val}`;
    const isSelected = selectedKw[catKey].has(key);
    return `<div class="kw-tag ${isSelected ? `selected-${catKey}` : ''}"
      onclick="toggleKwTag('${catKey}','${colName}','${val.replace(/'/g,"\\'")}')">
      ${val}
    </div>`;
  }).join('') || '<span style="color:#444; font-size:11px;">（無資料）</span>';
}

function toggleKwTag(catKey, colName, val) {
  const key = `${colName}::${val}`;
  if (selectedKw[catKey].has(key)) {
    selectedKw[catKey].delete(key);
  } else {
    selectedKw[catKey].add(key);
  }
  renderKwPool(catKey);
  updateKwBadge(catKey);
  renderSelectedBar();
  renderKeywordResults();
}

function updateKwBadge(catKey) {
  const badge = document.getElementById(`kwbadge-${catKey}`);
  if (!badge) return;
  const count = selectedKw[catKey].size;
  badge.textContent = count;
  badge.className = count > 0 ? 'kw-cat-badge has-sel' : 'kw-cat-badge';
}

function renderSelectedBar() {
  const bar = document.getElementById('kw-selected-bar');
  const tagsEl = document.getElementById('kw-selected-tags');
  if (!bar || !tagsEl) return;
  const allSel = [];
  Object.entries(selectedKw).forEach(([catKey, set]) => {
    set.forEach(key => allSel.push({ catKey, key }));
  });
  if (allSel.length === 0) {
    bar.style.display = 'none';
    return;
  }
  bar.style.display = 'block';
  tagsEl.innerHTML = allSel.map(({ catKey, key }) =>
    `<div class="kw-sel-chip">
      <span>${key.split('::')[1]}</span>
      <span class="chip-x" onclick="removeKwTag('${catKey}','${key.replace(/'/g,"\\'")}')">✕</span>
    </div>`
  ).join('');
}

function removeKwTag(catKey, key) {
  selectedKw[catKey].delete(key);
  renderKwPool(catKey);
  updateKwBadge(catKey);
  renderSelectedBar();
  renderKeywordResults();
}

function clearAllKeywords() {
  Object.keys(selectedKw).forEach(k => selectedKw[k].clear());
  Object.keys(KW_CATS).forEach(catKey => {
    renderKwPool(catKey);
    updateKwBadge(catKey);
  });
  renderSelectedBar();
  renderKeywordResults();
}

/* ════════════════════════════════════════
   ★ 重工資料庫 - 渲染結果
   ════════════════════════════════════════ */
function renderKeywordResults() {
  const container = document.getElementById('keyword-issue-display');
  const countEl   = document.getElementById('kw-result-count');
  if (!container) return;

  const totalSelected = Object.values(selectedKw).reduce((s, set) => s + set.size, 0);
  if (totalSelected === 0) {
    container.innerHTML = `<div style="grid-column:1/-1; text-align:center; color:#285428; padding:60px; font-size:14px;">← 從左側選擇關鍵字來篩選 Issue</div>`;
    if (countEl) countEl.textContent = '[ 請選擇關鍵字 ]';
    return;
  }

  // 每個大類：只要 Issue 的該欄包含任意一個已選關鍵字即符合
  // 四大類之間是 AND（都要符合有選的那些類），同類內多選是 OR
  const filtered = allIssues.filter(i => {
    if (i.id && String(i.id).startsWith('MGR-')) return false;
    return Object.entries(selectedKw).every(([catKey, set]) => {
      if (set.size === 0) return true; // 該類未選任何 → 不限制
      const fieldVal = getIssueKwField(i, catKey) || '';
      // fieldVal 格式: "小項::值||小項::值||..."
      const parts = fieldVal.split('||').map(s => s.trim()).filter(Boolean);
      return Array.from(set).some(key => parts.includes(key));
    });
  });

  if (countEl) countEl.textContent = `[ 找到 ${filtered.length} 筆 Issue ]`;

  container.innerHTML = filtered.map(i => {
    const stat = String(i.status);
    const isDone = (stat === "已解決" || stat === "Done");
    const isUrgent = (!isDone && isTaskUrgent(i.deadline, stat));
    return `<div class="pebble ${isDone ? 'resolved-card' : ''} ${isUrgent ? 'urgent-card' : ''}" onclick="openEdit('${i.id}')">
      <div style="font-size:11px; color:${isUrgent ? '#ff0055' : 'var(--pixel-green)'};">[ ${stat} ]</div>
      <div style="font-size:20px; margin:10px 0; line-height:1.3;">${i.issue}</div>
      <div style="font-size:12px; opacity:0.6;">${i.product} | ${i.owner}</div>
    </div>`;
  }).join('') || `<div style="grid-column:1/-1; text-align:center; color:#444; padding:60px; font-size:14px;">無符合的 Issue</div>`;
}

// 取得 Issue 對應大類欄位
function getIssueKwField(issue, catKey) {
  const fieldMap = { industry: 'kw_industry', system: 'kw_system', hardware: 'kw_hardware', spec: 'kw_spec' };
  return issue[fieldMap[catKey]] || '';
}

/* ════════════════════════════════════════
   ★ Modal 關鍵詞彙 - 初始化
   ════════════════════════════════════════ */
function initModalKeywords(savedValues) {
  // savedValues = { industry: "小項::值", system: "小項::值", ... } or all null
  kwModalSelected = { industry: null, system: null, hardware: null, spec: null };
  kwModalActiveTab = { industry: null, system: null, hardware: null, spec: null };

  Object.entries(KW_CATS).forEach(([catKey, catDef]) => {
    const subsEl = document.getElementById(`kw-mc-${catKey}-subs`);
    const poolEl = document.getElementById(`kw-mc-${catKey}-pool`);
    if (!subsEl || !poolEl) return;

    // 如果有已存值，恢復
    if (savedValues && savedValues[catKey]) {
      kwModalSelected[catKey] = savedValues[catKey];
      const [col] = savedValues[catKey].split('::');
      kwModalActiveTab[catKey] = col;
    } else {
      kwModalActiveTab[catKey] = catDef.cols[0] || null;
    }

    // 小項按鈕
    subsEl.innerHTML = catDef.cols.map(col =>
      `<button type="button" class="kw-mc-sub-btn" id="kw-mc-sub-${catKey}-${col.replace(/\s/g,'_')}"
        onclick="setModalKwTab('${catKey}','${col}')">${col}</button>`
    ).join('');

    renderModalKwPool(catKey);
    updateModalSubTabs(catKey);
  });
}

function setModalKwTab(catKey, colName) {
  kwModalActiveTab[catKey] = colName;
  updateModalSubTabs(catKey);
  renderModalKwPool(catKey);
}

function updateModalSubTabs(catKey) {
  KW_CATS[catKey].cols.forEach(col => {
    const btn = document.getElementById(`kw-mc-sub-${catKey}-${col.replace(/\s/g,'_')}`);
    if (!btn) return;
    btn.className = 'kw-mc-sub-btn';
    if (col === kwModalActiveTab[catKey]) btn.classList.add(`mc-active-${catKey}`);
  });
}

function renderModalKwPool(catKey) {
  const colName = kwModalActiveTab[catKey];
  const poolEl = document.getElementById(`kw-mc-${catKey}-pool`);
  if (!poolEl || !colName) return;
  const values = (kwData[colName] || []).filter(v => v && v.toString().trim());
  const tagHtml = values.map(val => {
    const key = `${colName}::${val}`;
    const isSelected = kwModalSelected[catKey] === key;
    return `<div class="kw-mc-tag ${isSelected ? `mc-sel-${catKey}` : ''}"
      onclick="toggleModalKwTag('${catKey}','${colName}','${val.replace(/'/g,"\\'")}')">
      ${val}
    </div>`;
  }).join('') || '<span style="color:#444; font-size:10px;">（無資料）</span>';
  poolEl.innerHTML = `${tagHtml}
    <div class="kw-inline-add" style="width:100%;">
      <input type="text" class="pixel-input" id="kw-add-${catKey}" placeholder="新增到「${colName}」...">
      <button type="button" onclick="submitInlineKeyword('${catKey}')">新增</button>
    </div>`;
}

function toggleModalKwTag(catKey, colName, val) {
  const key = `${colName}::${val}`;
  if (kwModalSelected[catKey] === key) {
    kwModalSelected[catKey] = null; // 再點一次取消
  } else {
    kwModalSelected[catKey] = key;
  }
  renderModalKwPool(catKey);
}

async function submitInlineKeyword(catKey) {
  const input = document.getElementById(`kw-add-${catKey}`);
  const colName = kwModalActiveTab[catKey];
  if (!input || !colName) return;
  const clean = normalizeDictValue(input.value);
  if (!clean) {
    showToast("請先輸入要新增的關鍵詞彙");
    return;
  }
  if (listHasValue(kwData[colName], clean)) {
    kwModalSelected[catKey] = `${colName}::${clean}`;
    renderModalKwPool(catKey);
    showToast("此標籤已存在，已幫你選取");
    return;
  }
  if (!confirmSimilarDictionaryValue(kwData[colName], clean, '關鍵詞彙')) return;
  try {
    addLocalDictionaryValue('keyword', clean, colName);
    kwModalSelected[catKey] = `${colName}::${clean}`;
    renderModalKwPool(catKey);
    renderKwPool(catKey);
    input.value = '';
    await saveDictionaryItem('keyword', clean, colName);
    showToast("[ TAG REGISTERED ]\n標籤已寫入重工資料庫", 3500);
    saveFormDraft();
  } catch(e) {
    showToast("新增失敗：請確認 Apps Script 已加入 addDictionaryItem");
  }
}

/* ════════════════════════════════════════
   ★ 表單驗證
   ════════════════════════════════════════ */
function validateForm() {
  let valid = true;
  const errors = [];

  // 清除舊錯誤
  document.querySelectorAll('.field-error').forEach(el => el.classList.remove('field-error'));
  document.querySelectorAll('.field-error-msg').forEach(el => el.remove());

  const requiredFields = [
    { id: 'input-issue',       label: 'Issue' },
    { id: 'input-owner',       label: 'Owner' },
    { id: 'input-status',      label: 'Status' },
    { id: 'input-customer',    label: '客戶別',  ssId: 'ss-customer-input' },
    { id: 'input-product',     label: '產品別',  ssId: 'ss-product-input' },
    { id: 'input-project',     label: '專案別',  ssId: 'ss-project-input' },
    { id: 'input-deadline',    label: '預計完成時間' },
    { id: 'input-priority',    label: '優先級' },
    { id: 'input-description', label: '任務描述' },
  ];

  requiredFields.forEach(f => {
    const el = document.getElementById(f.id);
    if (!el) return;
    const val = el.value ? el.value.trim() : '';
    if (!val) {
      valid = false;
      errors.push(f.label);
      const fieldItem = el.closest('.field-item') || el.parentElement;
      if (fieldItem) {
        fieldItem.classList.add('field-error');
        const errMsg = document.createElement('div');
        errMsg.className = 'field-error-msg';
        errMsg.textContent = `${f.label} 為必填項目`;
        fieldItem.appendChild(errMsg);
      }
    }
  });

  // ★ 關鍵詞彙：四大類中至少選一個即可
  const hasAnyKeyword = Object.values(kwModalSelected).some(Boolean);
  if (!hasAnyKeyword) {
    valid = false;
    errors.push('關鍵詞彙');
    const kwBlock = document.getElementById('field-kw-block');
    if (kwBlock) {
      kwBlock.classList.add('field-error');
      // 移除舊的錯誤訊息再加新的
      kwBlock.querySelectorAll('.field-error-msg').forEach(el => el.remove());
      const errMsg = document.createElement('div');
      errMsg.className = 'field-error-msg';
      errMsg.textContent = '關鍵詞彙必填：四大類中至少選擇一個即可';
      kwBlock.appendChild(errMsg);
    }
  }

  if (!valid) {
    showToast(`⚠ 請填寫必填項目：${errors.join('、')}`, 4000);
  }
  return valid;
}

/* ════════════════════════════════════════
   事項紀錄
   ════════════════════════════════════════ */
function addRecordItem(text = "", checked = false) {
  const container = document.getElementById('records-container');
  const div = document.createElement('div');
  div.className = 'record-item-row';

  const sepIdx     = text.indexOf('::');
  const titleVal   = sepIdx >= 0 ? text.substring(0, sepIdx) : text;
  const contentVal = sepIdx >= 0 ? text.substring(sepIdx + 2) : '';
  const hasContent = contentVal.trim() !== '';
  const titlePreview = titleVal.trim() !== '' ? titleVal : '（點擊編輯）';

  div.innerHTML = `
    <input type="checkbox" class="record-chk" ${checked ? 'checked' : ''} onchange="this.parentElement.querySelector('.record-preview').classList.toggle('record-done', this.checked); saveFormDraft();">
    <input type="hidden" class="record-title-input"   value="${titleVal.replace(/"/g, '&quot;')}">
    <input type="hidden" class="record-content-input" value="${contentVal.replace(/"/g, '&quot;')}">
    <div class="record-node" aria-hidden="true"></div>
    <div class="pixel-input record-preview ${checked ? 'record-done' : ''}"
      onclick="openRecordSub(this)">
      <span class="record-title-text">${escapeHtml(titlePreview)}</span>
      <span class="record-content-badge" style="color:${hasContent ? '#aaffaa' : '#444'};">
        ${hasContent ? '[有內容]' : '[無內容]'}
      </span>
    </div>
    <button type="button" class="pixel-btn record-remove-btn"
      onclick="this.parentElement.remove(); saveFormDraft();">X</button>
  `;
  container.appendChild(div);
  saveFormDraft();
}

let _currentRecordRow = null;

function openRecordSub(previewEl) {
  _currentRecordRow = previewEl.closest('.record-item-row');
  const titleHidden   = _currentRecordRow.querySelector('.record-title-input');
  const contentHidden = _currentRecordRow.querySelector('.record-content-input');
  document.getElementById('record-sub-title-input').value = titleHidden   ? titleHidden.value   : '';
  document.getElementById('record-sub-text').value        = contentHidden ? contentHidden.value : '';
  document.getElementById('record-sub-overlay').style.display = 'flex';
}

function closeRecordSub() {
  document.getElementById('record-sub-overlay').style.display = 'none';
  _currentRecordRow = null;
}

function saveRecordSub() {
  const newTitle   = document.getElementById('record-sub-title-input').value;
  const newContent = document.getElementById('record-sub-text').value;
  if (_currentRecordRow) {
    _currentRecordRow.querySelector('.record-title-input').value   = newTitle;
    _currentRecordRow.querySelector('.record-content-input').value = newContent;
    const preview    = _currentRecordRow.querySelector('.record-preview');
    const titleSpan  = preview.querySelector('.record-title-text');
    const badgeSpan  = preview.querySelector('.record-content-badge');
    const hasContent = newContent.trim() !== '';
    titleSpan.textContent = newTitle.trim() !== '' ? newTitle : '（點擊編輯）';
    badgeSpan.textContent = hasContent ? '[有內容]' : '[無內容]';
    badgeSpan.style.color = hasContent ? '#aaffaa' : '#444';
    preview.classList.toggle('record-empty', !newTitle.trim());
    saveFormDraft();
  }
  closeRecordSub();
}

/* ════════════════════════════════════════
   相關檔案連結
   ════════════════════════════════════════ */
function addLinkItem(url = "") {
  const group = document.getElementById('link-group');
  const div = document.createElement('div');
  div.style.cssText = "display:flex; gap:8px; margin-bottom:8px;";
  div.innerHTML = `
    <input type="text" class="pixel-input link-entry" style="flex:1;"
      value="${url}" placeholder="https://...">
    <button type="button" class="pixel-btn"
      style="background:#444; padding:5px 10px; border:none; color:#fff; flex-shrink:0; cursor:pointer;"
      onclick="this.parentElement.remove(); saveFormDraft();">X</button>
  `;
  group.appendChild(div);
  saveFormDraft();
}

/* ════════════════════════════════════════
   Status 變動
   ════════════════════════════════════════ */
function handleStatusChange() {
  const status = document.getElementById('input-status').value;
  const actualField = document.getElementById('input-actual-closed');
  if (status === "已解決" || status === "Done") {
    const closedDate = prompt("此項目已完成！請確認實際結案日期 (YYYY-MM-DD):", getToday2026());
    if (closedDate) {
      window.currentClosedDate = closedDate;
      actualField.value = closedDate;
      actualField.style.opacity = "1";
    } else { document.getElementById('input-status').value = ""; actualField.value = ""; }
  } else { actualField.value = ""; actualField.style.opacity = "0.5"; }
}

/* ════════════════════════════════════════
   openModal
   ════════════════════════════════════════ */
function openModal(type) {
  window.currentModalType = type;
  fillUIConfigs();
  const modal = document.getElementById('modal-container');
  modal.className = type === 'MGR' ? 'pixel-modal navy-theme modal-manager-pink' : 'pixel-modal navy-theme';
  document.getElementById('issueForm').reset();
  document.getElementById('records-container').innerHTML = '';
  document.getElementById('edit-id').value = "";
  document.getElementById('input-creator').value = currentUser.name;
  document.getElementById('input-deadline').value = "";
  document.getElementById('input-priority').value = "";
  document.getElementById('input-actual-closed').value = "";
  document.getElementById('input-created-date').value = new Date().toLocaleDateString('zh-TW');
  window.currentClosedDate = "";
  loadSearchableSelectValue('ss-customer', '');
  loadSearchableSelectValue('ss-product', '');
  loadSearchableSelectValue('ss-project', '');
  addRecordItem();
  document.getElementById('link-group').innerHTML = '';
  addLinkItem();
  initModalKeywords(null);
  document.getElementById('modal-overlay').style.display = 'flex';
  document.getElementById('submit-btn').innerText = "建立完成";
  document.getElementById('btn-delete').style.display = 'none';
  tryRestoreFormDraft();
  // 清除驗證錯誤
  document.querySelectorAll('.field-error').forEach(el => el.classList.remove('field-error'));
  document.querySelectorAll('.field-error-msg').forEach(el => el.remove());
}

/* ════════════════════════════════════════
   openEdit
   ════════════════════════════════════════ */
function openEdit(id) {
  fillUIConfigs();
  const i = allIssues.find(x => x.id === id);
  if (!i) return;
  openModal(id.startsWith('MGR-') ? 'MGR' : 'TS');
  document.getElementById('edit-id').value = i.id;
  document.getElementById('input-issue').value = i.issue;
  document.getElementById('input-owner').value = i.owner;
  document.getElementById('input-status').value = i.status;
  document.getElementById('input-priority').value = i.priority;
  document.getElementById('input-deadline').value = i.deadline ? i.deadline.replace(/\//g, '-') : "";
  document.getElementById('input-description').value = i.description || "";
  document.getElementById('input-creator').value = i.creator || currentUser.name;
  document.getElementById('input-created-date').value = i.date;
  const actual = i.closedDate ? i.closedDate.replace(/\//g, '-') : "";
  document.getElementById('input-actual-closed').value = actual;
  window.currentClosedDate = actual;
  loadSearchableSelectValue('ss-customer', i.customer);
  loadSearchableSelectValue('ss-product',  i.product);
  loadSearchableSelectValue('ss-project',  i.project);
  // 事項紀錄
  const container = document.getElementById('records-container');
  container.innerHTML = '';
  (i.records || "").split('||').forEach(item => {
    if (item) addRecordItem(item.substring(3), item.startsWith('[v]'));
  });
  if (container.innerHTML === '') addRecordItem();
  // 相關連結
  const linkGroup = document.getElementById('link-group');
  linkGroup.innerHTML = "";
  const links = (i.link || "").split(' | ').filter(u => u.trim());
  if (links.length > 0) { links.forEach(url => addLinkItem(url)); }
  else { addLinkItem(); }
  // 關鍵詞彙恢復
  initModalKeywords({
    industry: i.kw_industry || null,
    system:   i.kw_system   || null,
    hardware: i.kw_hardware || null,
    spec:     i.kw_spec     || null
  });
  document.getElementById('submit-btn').innerText = "編輯完成";
  document.getElementById('btn-delete').style.display = 'inline-block';
}

/* ════════════════════════════════════════
   submitIssue
   ════════════════════════════════════════ */
async function submitIssue() {
  // 先執行驗證
  if (!validateForm()) return;

  const btn = document.getElementById('submit-btn');
  btn.disabled = true; btn.innerText = "同步中...";

  const recs = Array.from(document.querySelectorAll('.record-item-row')).map(row => {
    const chk     = row.querySelector('.record-chk').checked ? '[v]' : '[ ]';
    const title   = row.querySelector('.record-title-input').value;
    const content = row.querySelector('.record-content-input').value;
    return chk + title + '::' + content;
  }).join('||');

  const issueId = document.getElementById('edit-id').value ||
    (window.currentModalType === 'MGR' ? 'MGR-' : 'TS-') + Date.now();

  // 收集關鍵詞彙（每大類最多一個 "小項::值"，用 || 分隔不同類的多選不適用此格式，直接存單值）
  const kw_industry = kwModalSelected.industry || '';
  const kw_system   = kwModalSelected.system   || '';
  const kw_hardware = kwModalSelected.hardware || '';
  const kw_spec     = kwModalSelected.spec     || '';

  const payload = {
    action: document.getElementById('edit-id').value ? "edit" : "add",
    sheetName: (issueId.startsWith('MGR-') ? "主管事務" : "Issues"),
    id: issueId,
    issue:    document.getElementById('input-issue').value,
    owner:    document.getElementById('input-owner').value,
    status:   document.getElementById('input-status').value,
    customer: document.getElementById('input-customer').value,
    product:  document.getElementById('input-product').value,
    project:  document.getElementById('input-project').value,
    date:     document.getElementById('input-created-date').value,
    deadline: document.getElementById('input-deadline').value,
    priority: document.getElementById('input-priority').value,
    description: document.getElementById('input-description').value,
    records:  recs,
    link: Array.from(document.querySelectorAll('#link-group .link-entry'))
            .map(el => el.value).filter(v => v.trim()).join(' | '),
    closedDate: document.getElementById('input-actual-closed').value || "",
    creator:  document.getElementById('input-creator').value,
    kw_industry, kw_system, kw_hardware, kw_spec
  };
  try {
    isMutating = true;
    await saveNewModalDictionaryValues();
    await fetch(SCRIPT_URL, { method: 'POST', mode: 'no-cors', body: JSON.stringify(payload) });
    isMutating = false;
    clearFormDraft();
    alert("同步成功!");
    closeModal();
    await fetchDataOnLoad();
    renderIssues(); renderManagerIssues(); renderStats();
  } catch(e) {
    isMutating = false;
    alert(e.message && e.message.includes('取消新增') ? "已取消同步" : "同步失敗");
  }
  btn.disabled = false; btn.innerText = "完成";
}

/* ════════════════════════════════════════
   closeModal / deleteIssue
   ════════════════════════════════════════ */
function closeModal() {
  document.getElementById('modal-overlay').style.display = 'none';
}

async function deleteIssue() {
  const pwd = prompt("確認密碼:");
  if (pwd !== "59075364" && pwd !== "13091309") return;
  const id = document.getElementById('edit-id').value;
  isMutating = true;
  await fetch(SCRIPT_URL, {
    method: 'POST', mode: 'no-cors',
    body: JSON.stringify({ action: "delete", id: id, sheetName: (window.currentModalType === 'MGR' ? "主管事務" : "Issues") })
  });
  isMutating = false;
  closeModal();
  await fetchDataOnLoad();
  renderIssues(); renderManagerIssues();
}
