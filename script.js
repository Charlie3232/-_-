const SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbzByMLLdHFcDi3tPNQ3tfeNGpoKF0XtDTkj0MLxWrNpjweB-zRbjSG2KxdunwAzADiE/exec'; 
let allIssues = [];
let dataConfig = {};
let userList = []; 
let currentUser = { id: "", name: "", role: "" };
let isMutating = false;

async function handleLogin() {
  const idInput = document.getElementById('login-user').value.trim();
  const pwdInput = document.getElementById('login-pwd').value.trim();
  
  // 優先抓取資料
  if (userList.length === 0) await fetchData();

  const user = userList.find(u => u.id === idInput && u.pwd === pwdInput);
  
  if (user) {
    currentUser = { id: user.id, name: user.name, role: (user.id === "G0006" ? "MANAGER" : "USER") };
    document.getElementById('current-username').innerText = currentUser.name;
    document.getElementById('login-overlay').style.display = 'none';
    document.getElementById('main-ui').style.display = 'block';
    
    if (currentUser.role === "MANAGER") {
      document.getElementById('btn-tab-main').style.display = 'block';
      document.getElementById('btn-tab-manager').style.display = 'block';
    }
    init();
  } else {
    alert("驗證失敗：人員識別碼或密碼不正確");
  }
}

async function fetchData() {
  const resp = await fetch(SCRIPT_URL + '?action=getData');
  const data = await resp.json();
  allIssues = data.issues || [];
  dataConfig = data.config || {};
  userList = data.users || [];

  fillCheckboxes('items-owner', 'owners', data, 'renderIssues()');
  fillCheckboxes('items-status', 'statusList', data, 'renderIssues()');
  fillFormSelect('input-owner', 'owners');
  fillFormSelect('input-status', 'statusList');
  fillFormSelect('input-customer', 'customers');
  fillFormSelect('input-product', 'products'); // D 欄
  fillFormSelect('input-project', 'projects'); // E 欄
}

function init() {
  renderIssues();
  renderManagerIssues();
  renderStats();
}

function switchTab(tabId) {
  document.querySelectorAll('.tab-section').forEach(el => el.style.display = 'none');
  document.querySelectorAll('.nav-btn').forEach(btn => btn.classList.remove('active'));
  document.getElementById(tabId).style.display = 'block';
  document.getElementById('btn-' + tabId).classList.add('active');
}

function openModal(type = 'TS') {
  document.getElementById('issueForm').reset();
  document.getElementById('edit-id').value = "";
  document.getElementById('input-creator').value = currentUser.name;
  document.getElementById('modal-overlay').style.display = 'flex';
  document.getElementById('btn-delete').style.display = 'none';
}

function fillFormSelect(id, listKey) {
  const el = document.getElementById(id);
  if(el && dataConfig[listKey]) {
    el.innerHTML = '<option value="" disabled selected>請選擇...</option>' + 
      dataConfig[listKey].map(t => `<option value="${t}">${t}</option>`).join('');
  }
}

// ...其餘 renderIssues, submitIssue 等邏輯與原版相同，僅需確保欄位 ID 對應...