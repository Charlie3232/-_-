const SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbzXA_jatbThi0ZXvHSBweJGTINLfmAD-TP5q_AM1-8ZD2Kd9QHLD7gEL6SLLQc9QOmf/exec';
let allIssues = [];
let dataConfig = {};
let isMutating = false;
let userList = [];
let currentUser = { id: "", name: "", role: "" };
let currentModalType = 'TS';

const getToday2026 = () => {
  const d = new Date();
  return `2026-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
};

/* ════════════════════════════════════════
   一般 select 填充（Owner / Status，無變動）
   ════════════════════════════════════════ */
function fillFormSelect(id, list) {
  const el = document.getElementById(id);
  if (el && dataConfig[list]) {
    el.innerHTML = dataConfig[list].map(t => `<option value="${t}">${t}</option>`).join('');
    el.insertAdjacentHTML('afterbegin', '<option value="" disabled selected>請選擇...</option>');
  }
}

/* ════════════════════════════════════════
   ★ 新增：可搜尋下拉元件（Modal 用）
   ════════════════════════════════════════ */
/**
 * 初始化一個可搜尋下拉
 * @param {string} wrapperId  - searchable-select 容器 id（例如 'ss-customer'）
 * @param {string} listKey    - dataConfig 中的清單 key（例如 'customers'）
 */
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
  // 關閉其他所有 ss-dropdown
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
}

/** 從 hidden input 讀值，填回 text input 顯示 */
function loadSearchableSelectValue(wrapperId, value) {
  const textInput = document.getElementById(`${wrapperId}-input`);
  const hiddenInput = document.getElementById(`input-${wrapperId.replace('ss-', '')}`);
  if (textInput) textInput.value = value || '';
  if (hiddenInput) hiddenInput.value = value || '';
}

/* ════════════════════════════════════════
   登入（無變動）
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
    if (document.getElementById('login-status')) document.getElementById('login-status').innerText = "系統就緒。";
  } catch(e) { console.error("Initial load failed", e); }
}
window.onload = fetchDataOnLoad;

function initUI() {
  fillUIConfigs(); renderIssues(); renderManagerIssues(); renderStats();
  setInterval(silentSync, 15000);
}

async function silentSync() {
  if (isMutating || document.getElementById('main-ui').style.display === 'none') return;
  try {
    const resp = await fetch(SCRIPT_URL + '?action=getData');
    const data = await resp.json();
    if (isMutating) return;
    allIssues = data.issues || [];
    renderIssues(); renderManagerIssues(); renderStats();
  } catch(e) {}
}

/* ════════════════════════════════════════
   ★ 新增：篩選下拉搜尋框
   ════════════════════════════════════════ */
/**
 * 過濾篩選下拉的 checkbox 選項
 * @param {HTMLInputElement} inputEl  - 搜尋框元素
 * @param {string} dropdownId         - 下拉容器 id（例如 'items-owner'）
 */
function filterDropdownSearch(inputEl, dropdownId) {
  const query = inputEl.value.toLowerCase();
  const listEl = document.getElementById(`${dropdownId}-list`);
  if (!listEl) return;
  Array.from(listEl.querySelectorAll('.checkbox-label')).forEach(label => {
    const text = label.textContent.toLowerCase();
    label.style.display = text.includes(query) ? '' : 'none';
  });
}

/* ════════════════════════════════════════
   ★ 新增：全選 / 全清
   ════════════════════════════════════════ */
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
  // 也清空搜尋框
  const parent = document.getElementById(dropdownId);
  if (parent) {
    const searchInput = parent.querySelector('.filter-search-input');
    if (searchInput) {
      searchInput.value = '';
      filterDropdownSearch(searchInput, dropdownId);
    }
  }
  eval(onChangeCode);
}

/* ════════════════════════════════════════
   填充 checkbox 篩選（★ 變動：寫入 -list 子容器）
   ════════════════════════════════════════ */
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

/* ════════════════════════════════════════
   ★ 變動：fillUIConfigs 新增客戶別篩選
   ════════════════════════════════════════ */
function fillUIConfigs() {
  // 篩選 checkboxes（TS tab）
  fillCheckboxes('items-owner',    'owners',     'renderIssues()');
  fillCheckboxes('items-status',   'statusList', 'renderIssues()');
  fillCheckboxes('items-product',  'products',   'renderIssues()');
  fillCheckboxes('items-customer', 'customers',  'renderIssues()');  // 新增

  // 篩選 checkboxes（MGR tab）
  fillCheckboxes('mgr-owner',    'owners',     'renderManagerIssues()');
  fillCheckboxes('mgr-status',   'statusList', 'renderManagerIssues()');
  fillCheckboxes('mgr-product',  'products',   'renderManagerIssues()');
  fillCheckboxes('mgr-customer', 'customers',  'renderManagerIssues()');  // 新增

  // Modal 一般 select
  fillFormSelect('input-owner',  'owners');
  fillFormSelect('input-status', 'statusList');

  // Modal 可搜尋下拉（★ 變動：取代原本的 fillFormSelect）
  initSearchableSelect('ss-customer', 'customers');
  initSearchableSelect('ss-product',  'products');
  initSearchableSelect('ss-project',  'projects');
}

/* ════════════════════════════════════════
   取得已勾選篩選值（無變動）
   ════════════════════════════════════════ */
const getCheckedValues = (id) => {
  // 從 -list 子容器取值
  const listEl = document.getElementById(`${id}-list`);
  const source = listEl || document.getElementById(id);
  if (!source) return [];
  return Array.from(source.querySelectorAll('input[type="checkbox"]:checked')).map(cb => cb.value);
};

/* ════════════════════════════════════════
   緊急判斷（無變動）
   ════════════════════════════════════════ */
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
      return `<div class="stat-row" style="margin-bottom:12px; display:flex; align-items:center; gap:10px;">
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
      return `<div class="stat-row" style="margin-bottom:12px; display:flex; align-items:center; gap:10px;">
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
   ★ 變動：renderIssues 加入客戶別篩選
   ════════════════════════════════════════ */
function renderIssues() {
  const container = document.getElementById('issue-display');
  const search   = document.getElementById('search-input').value.toLowerCase();
  const fOwners  = getCheckedValues('items-owner');
  const fStats   = getCheckedValues('items-status');
  const fProds   = getCheckedValues('items-product');
  const fCusts   = getCheckedValues('items-customer');  // 新增

  let filtered = allIssues.filter(i =>
    (!i.id || !String(i.id).startsWith('MGR-')) &&
    String(i.issue).toLowerCase().includes(search) &&
    (fOwners.length === 0 || fOwners.includes(i.owner)) &&
    (fStats.length === 0 ? (i.status !== "已解決" && i.status !== "Done") : fStats.includes(i.status)) &&
    (fProds.length === 0 || fProds.includes(i.product)) &&
    (fCusts.length === 0 || fCusts.includes(i.customer))  // 新增
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
   ★ 變動：renderManagerIssues 加入客戶別篩選
   ════════════════════════════════════════ */
function renderManagerIssues() {
  const container = document.getElementById('manager-issue-display');
  const search   = document.getElementById('search-input-mgr').value.toLowerCase();
  const fOwners  = getCheckedValues('mgr-owner');
  const fStats   = getCheckedValues('mgr-status');
  const fProds   = getCheckedValues('mgr-product');
  const fCusts   = getCheckedValues('mgr-customer');  // 新增

  let filtered = allIssues.filter(i =>
    i.id && String(i.id).startsWith('MGR-') &&
    String(i.issue).toLowerCase().includes(search) &&
    (fOwners.length === 0 || fOwners.includes(i.owner)) &&
    (fStats.length === 0 ? (i.status !== "已解決") : fStats.includes(i.status)) &&
    (fProds.length === 0 || fProds.includes(i.product)) &&
    (fCusts.length === 0 || fCusts.includes(i.customer))  // 新增
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
   toggleDropdown（無變動）
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

  // ★ GIF 只在主管事務頁顯示
  const gifBox = document.getElementById('header-gif-box');
  if (gifBox) gifBox.style.display = (tabId === 'tab-manager') ? 'flex' : 'none';

  if (tabId === 'tab-main') renderStats();
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
   事項紀錄（無變動）
   ════════════════════════════════════════ */
function addRecordItem(text = "", checked = false) {
  const container = document.getElementById('records-container');
  const div = document.createElement('div');
  div.className = 'record-item-row';
  div.innerHTML = `
    <input type="checkbox" class="record-chk" ${checked ? 'checked' : ''}>
    <input type="text" class="pixel-input record-txt-input" style="flex:1; border-width:2px; font-size:14px;" value="${text}" required>
    <button type="button" class="pixel-btn" style="background:#444; padding:5px 10px; border:none; color:#fff;" onclick="this.parentElement.remove()">X</button>
  `;
  container.appendChild(div);
}

/* ════════════════════════════════════════
   Status 變動處理（無變動）
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
   ★ 變動：openModal 初始化可搜尋下拉
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

  // 清空可搜尋下拉
  loadSearchableSelectValue('ss-customer', '');
  loadSearchableSelectValue('ss-product', '');
  loadSearchableSelectValue('ss-project', '');

  addRecordItem();
  document.getElementById('modal-overlay').style.display = 'flex';
  document.getElementById('submit-btn').innerText = "建立完成";
  document.getElementById('btn-delete').style.display = 'none';
}

/* ════════════════════════════════════════
   ★ 變動：openEdit 填入可搜尋下拉值
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

  // 可搜尋下拉填值
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

  // 連結
  const linkGroup = document.getElementById('link-group');
  linkGroup.innerHTML = "";
  (i.link || "").split(' | ').forEach(url => {
    if (!url) return;
    const input = document.createElement('input');
    input.className = "pixel-input wide link-entry";
    input.value = url;
    linkGroup.appendChild(input);
  });
  if (!linkGroup.innerHTML) {
    const input = document.createElement('input');
    input.className = "pixel-input wide link-entry";
    input.placeholder = "https://...";
    linkGroup.appendChild(input);
  }

  document.getElementById('submit-btn').innerText = "編輯完成";
  document.getElementById('btn-delete').style.display = 'inline-block';
}

/* ════════════════════════════════════════
   ★ 變動：submitIssue 從 hidden input 取客戶/產品/專案值
   ════════════════════════════════════════ */
async function submitIssue() {
  const form = document.getElementById('issueForm');
  if (!form.checkValidity()) { form.reportValidity(); return; }
  const btn = document.getElementById('submit-btn');
  btn.disabled = true; btn.innerText = "同步中...";

  const recs = Array.from(document.querySelectorAll('.record-item-row')).map(row => {
    const chk = row.querySelector('.record-chk').checked ? '[v]' : '[ ]';
    return chk + row.querySelector('.record-txt-input').value;
  }).join('||');

  const issueId = document.getElementById('edit-id').value ||
    (window.currentModalType === 'MGR' ? 'MGR-' : 'TS-') + Date.now();

  const payload = {
    action: document.getElementById('edit-id').value ? "edit" : "add",
    sheetName: (issueId.startsWith('MGR-') ? "主管事務" : "Issues"),
    id: issueId,
    issue:    document.getElementById('input-issue').value,
    owner:    document.getElementById('input-owner').value,
    status:   document.getElementById('input-status').value,
    customer: document.getElementById('input-customer').value,   // hidden input
    product:  document.getElementById('input-product').value,    // hidden input
    project:  document.getElementById('input-project').value,    // hidden input
    date:     document.getElementById('input-created-date').value,
    deadline: document.getElementById('input-deadline').value,
    priority: document.getElementById('input-priority').value,
    description: document.getElementById('input-description').value,
    records:  recs,
    link:     document.querySelector('.link-entry').value || "",
    closedDate: document.getElementById('input-actual-closed').value || "",
    creator:  document.getElementById('input-creator').value
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
   關閉 Modal / 刪除（無變動）
   ════════════════════════════════════════ */
function closeModal() {
  document.getElementById('modal-overlay').style.display = 'none';
}

async function deleteIssue() {
  const pwd = prompt("確認密碼:");
  if (pwd !== "13091309" && pwd !== "13321332") return;
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
