const SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbwR7WpnX-GBwkikW0YuPMIwoY0evLu8fDFGaLY-4Af0itgauxOcpl6sjoeHvRLB6e8m/exec';
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
  toast.style.display = 'block';
  clearTimeout(toast._timer);
  toast._timer = setTimeout(() => { toast.style.display = 'none'; }, duration);
}

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
  const listKey = document.getElementById(`${wrapperId}-dropdown`).dataset.list;
  if (!listKey || !dataConfig[listKey]) return;
  const query = input.value.toLowerCase();
  const filtered = dataConfig[listKey].filter(opt => opt.toLowerCase().includes(query));
  renderSearchableOptions(wrapperId, filtered);
  document.getElementById(`${wrapperId}-dropdown`).style.display = 'block';
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
      <div style="font-size:11px; color:${isUrgent ? '#ff0055' : 'var(--pixel-green)'};">[ ${stat} ]</div>
      <div style="font-size:20px; margin:10px 0; line-height:1.3;">${i.issue}</div>
      <div style="font-size:12px; opacity:0.6;">${i.product} | ${i.owner}</div>
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
  poolEl.innerHTML = values.map(val => {
    const key = `${colName}::${val}`;
    const isSelected = kwModalSelected[catKey] === key;
    return `<div class="kw-mc-tag ${isSelected ? `mc-sel-${catKey}` : ''}"
      onclick="toggleModalKwTag('${catKey}','${colName}','${val.replace(/'/g,"\\'")}')">
      ${val}
    </div>`;
  }).join('') || '<span style="color:#444; font-size:10px;">（無資料）</span>';
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

  // ★ 關鍵詞彙：四大類各至少選一個
  const kwCatLabels = { industry: '產業範疇', system: '系統類別', hardware: '硬件組成', spec: '技術規範' };
  const kwMissing = [];
  Object.entries(kwCatLabels).forEach(([catKey, catLabel]) => {
    if (!kwModalSelected[catKey]) {
      kwMissing.push(catLabel);
    }
  });
  if (kwMissing.length > 0) {
    valid = false;
    errors.push(`關鍵詞彙（${kwMissing.join('、')}）`);
    const kwBlock = document.getElementById('field-kw-block');
    if (kwBlock) {
      kwBlock.classList.add('field-error');
      // 移除舊的錯誤訊息再加新的
      kwBlock.querySelectorAll('.field-error-msg').forEach(el => el.remove());
      const errMsg = document.createElement('div');
      errMsg.className = 'field-error-msg';
      errMsg.textContent = `關鍵詞彙必填：${kwMissing.join('、')} 尚未選擇`;
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
    <input type="checkbox" class="record-chk" ${checked ? 'checked' : ''}>
    <input type="hidden" class="record-title-input"   value="${titleVal.replace(/"/g, '&quot;')}">
    <input type="hidden" class="record-content-input" value="${contentVal.replace(/"/g, '&quot;')}">
    <div class="pixel-input record-preview"
      style="flex:1; border-width:2px; font-size:14px; cursor:pointer; text-align:left;
             min-height:42px; display:flex; align-items:center; gap:8px; overflow:hidden;
             color:${titleVal.trim() ? 'var(--pixel-green)' : '#555'};"
      onclick="openRecordSub(this)">
      <span style="flex:1; overflow:hidden; white-space:nowrap; text-overflow:ellipsis;">${titlePreview}</span>
      <span style="font-size:11px; flex-shrink:0; color:${hasContent ? '#aaffaa' : '#444'};">
        ${hasContent ? '[有內容]' : '[無內容]'}
      </span>
    </div>
    <button type="button" class="pixel-btn"
      style="background:#444; padding:5px 10px; border:none; color:#fff; flex-shrink:0;"
      onclick="this.parentElement.remove()">X</button>
  `;
  container.appendChild(div);
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
    const titleSpan  = preview.querySelector('span:first-child');
    const badgeSpan  = preview.querySelector('span:last-child');
    const hasContent = newContent.trim() !== '';
    titleSpan.textContent = newTitle.trim() !== '' ? newTitle : '（點擊編輯）';
    badgeSpan.textContent = hasContent ? '[有內容]' : '[無內容]';
    badgeSpan.style.color = hasContent ? '#aaffaa' : '#444';
    preview.style.color   = newTitle.trim() ? 'var(--pixel-green)' : '#555';
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
      onclick="this.parentElement.remove()">X</button>
  `;
  group.appendChild(div);
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
    await fetch(SCRIPT_URL, { method: 'POST', mode: 'no-cors', body: JSON.stringify(payload) });
    isMutating = false;
    alert("同步成功!");
    closeModal();
    await fetchDataOnLoad();
    renderIssues(); renderManagerIssues(); renderStats();
  } catch(e) { alert("同步失敗"); }
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
