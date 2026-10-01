// ---------- Constants ----------
const DEFAULT_CATEGORIES = {
  expense: ["Food/Groceries","Gifts","Health/medical","Home","Transportation","Miscellaneous","Utilities","Travel","Debt","Credit Card Bill","Shopping","Custom category 1","Custom category 2","Custom category 3"],
  income: ["Gift","Paycheck","Dividend","Interest","Other","Refund"],
  investment: ["Investment","Share","Bonds","Mutual Fund","Fixed Deposit","New Home TDS","Other"]
};

let store = null;
let currentMonth = null;
let editState = null; // { type, monthKey, id } while editing a row
let auth = null;
let db = null;
let unsubscribeSnapshot = null;
let suppressNextSnapshotRender = false;

// ---------- Utilities ----------
function fmtMoney(n) {
  n = Number(n) || 0;
  return "₹" + n.toLocaleString("en-IN", { maximumFractionDigits: 2, minimumFractionDigits: 0 });
}
// Short form for chart labels, e.g. ₹2.28L or ₹45K
function fmtCompactINR(n) {
  n = Number(n) || 0;
  const abs = Math.abs(n);
  if (abs >= 100000) return "₹" + (n / 100000).toFixed(2).replace(/\.?0+$/, "") + "L";
  if (abs >= 1000) return "₹" + (n / 1000).toFixed(1).replace(/\.0$/, "") + "K";
  return "₹" + Math.round(n);
}
function monthKeyFromDate(dateStr) {
  return dateStr.slice(0, 7); // "YYYY-MM"
}
function monthLabelFromKey(key) {
  const [y, m] = key.split("-").map(Number);
  return new Date(y, m - 1, 1).toLocaleString("en-US", { month: "long", year: "numeric" });
}
function sortedMonthKeys() {
  return Object.keys(store.months).sort();
}
function newId() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
}
function uniquePush(arr, val) {
  if (val && !arr.includes(val)) arr.push(val);
}

// ---------- Data load / save ----------
function blankMonth(label, startingBalance, budget) {
  return {
    label,
    startingBalance: startingBalance || 0,
    budget: {
      expense: Object.assign({}, budget ? budget.expense : {}),
      income: Object.assign({}, budget ? budget.income : {})
    },
    transactions: { expense: [], income: [] },
    investments: []
  };
}

function seedFromExcelData() {
  const months = {};
  const categories = JSON.parse(JSON.stringify(DEFAULT_CATEGORIES));
  const seed = window.SEED_DATA || {};
  Object.keys(seed).forEach((key) => {
    const m = seed[key];
    const expense = (m.transactions.expense || []).map((t) => Object.assign({ id: newId() }, t));
    const income = (m.transactions.income || []).map((t) => Object.assign({ id: newId() }, t));
    const investments = (m.investments || []).map((t) => Object.assign({ id: newId() }, t));
    expense.forEach((t) => uniquePush(categories.expense, t.category));
    income.forEach((t) => uniquePush(categories.income, t.category));
    investments.forEach((t) => uniquePush(categories.investment, t.category));
    Object.keys(m.budget.expense || {}).forEach((c) => uniquePush(categories.expense, c));
    Object.keys(m.budget.income || {}).forEach((c) => uniquePush(categories.income, c));
    months[key] = {
      label: m.label,
      startingBalance: m.startingBalance || 0,
      budget: { expense: Object.assign({}, m.budget.expense), income: Object.assign({}, m.budget.income) },
      transactions: { expense, income },
      investments
    };
  });
  store = { months, categories };
}

function saveData() {
  if (!auth || !auth.currentUser) return;
  suppressNextSnapshotRender = true;
  docRefFor(auth.currentUser.uid).set(store).catch((err) => {
    console.error(err);
    alert("Could not save to the cloud: " + err.message);
  });
}

// ---------- Firebase auth + Firestore sync ----------
function docRefFor(uid) {
  return db.collection("users").doc(uid).collection("appData").doc("store");
}

function initFirebaseAuth() {
  firebase.initializeApp(window.FIREBASE_CONFIG);
  auth = firebase.auth();
  db = firebase.firestore();
  db.enablePersistence().catch(() => {}); // best-effort offline cache
  auth.onAuthStateChanged(handleAuthChange);
}

function handleAuthChange(user) {
  if (unsubscribeSnapshot) { unsubscribeSnapshot(); unsubscribeSnapshot = null; }
  if (user) {
    document.getElementById("loginOverlay").style.display = "none";
    document.getElementById("appRoot").style.display = "";
    showUserName(user);
    subscribeToData(user.uid);
  } else {
    store = null;
    currentMonth = null;
    document.getElementById("loginOverlay").style.display = "flex";
    document.getElementById("appRoot").style.display = "none";
  }
}

function showUserName(user) {
  const label = document.getElementById("userEmailLabel");
  if (!user.displayName) {
    const name = prompt("What should we call you? (shown here instead of your email)");
    if (name && name.trim()) {
      user.updateProfile({ displayName: name.trim() }).then(() => {
        label.textContent = name.trim();
      });
      return;
    }
  }
  label.textContent = user.displayName || user.email;
}

function renameUser() {
  const user = auth.currentUser;
  if (!user) return;
  const name = prompt("Enter your name:", user.displayName || "");
  if (name === null || !name.trim()) return;
  user.updateProfile({ displayName: name.trim() }).then(() => {
    document.getElementById("userEmailLabel").textContent = name.trim();
  });
}

function subscribeToData(uid) {
  unsubscribeSnapshot = docRefFor(uid).onSnapshot((snap) => {
    if (snap.exists) {
      if (suppressNextSnapshotRender) { suppressNextSnapshotRender = false; return; }
      store = snap.data();
      mergeNewSeedMonths();
      renderAll();
    } else {
      seedFromExcelData();
      saveData();
      renderAll();
    }
  }, (err) => {
    console.error(err);
    alert("Could not load data from the cloud: " + err.message);
  });
}

// Adds any months present in the local seed-data.js but missing from the cloud store (e.g. newly added historical Excel files), without touching existing months.
function mergeNewSeedMonths() {
  const seed = window.SEED_DATA || {};
  const missingKeys = Object.keys(seed).filter((key) => !store.months[key]);
  if (!missingKeys.length) return;

  missingKeys.forEach((key) => {
    const m = seed[key];
    const expense = (m.transactions.expense || []).map((t) => Object.assign({ id: newId() }, t));
    const income = (m.transactions.income || []).map((t) => Object.assign({ id: newId() }, t));
    const investments = (m.investments || []).map((t) => Object.assign({ id: newId() }, t));
    expense.forEach((t) => uniquePush(store.categories.expense, t.category));
    income.forEach((t) => uniquePush(store.categories.income, t.category));
    investments.forEach((t) => uniquePush(store.categories.investment, t.category));
    Object.keys(m.budget.expense || {}).forEach((c) => uniquePush(store.categories.expense, c));
    Object.keys(m.budget.income || {}).forEach((c) => uniquePush(store.categories.income, c));
    store.months[key] = {
      label: m.label,
      startingBalance: m.startingBalance || 0,
      budget: { expense: Object.assign({}, m.budget.expense), income: Object.assign({}, m.budget.income) },
      transactions: { expense, income },
      investments
    };
  });
  saveData();
}

// ---------- Month helpers ----------
function ensureMonthForDate(dateStr) {
  const key = monthKeyFromDate(dateStr);
  if (!store.months[key]) {
    const keys = sortedMonthKeys();
    let prevBudget = null, prevStartBalance = 0;
    if (keys.length) {
      const lastKey = keys[keys.length - 1];
      const lastMonth = store.months[lastKey];
      prevBudget = lastMonth.budget;
      prevStartBalance = computeMonthTotals(lastKey).endingBalance;
    }
    store.months[key] = blankMonth(monthLabelFromKey(key), prevStartBalance, prevBudget);
  }
  return key;
}

function addNewMonth() {
  const today = new Date();
  let candidate = new Date(today.getFullYear(), today.getMonth(), 1);
  const keys = sortedMonthKeys();
  if (keys.length) {
    const [y, m] = keys[keys.length - 1].split("-").map(Number);
    candidate = new Date(y, m, 1); // next month after the latest one
  }
  const key = `${candidate.getFullYear()}-${String(candidate.getMonth() + 1).padStart(2, "0")}`;
  if (store.months[key]) {
    alert(monthLabelFromKey(key) + " already exists.");
    return;
  }
  const lastKey = keys[keys.length - 1];
  const prevBudget = lastKey ? store.months[lastKey].budget : null;
  const prevStartBalance = lastKey ? computeMonthTotals(lastKey).endingBalance : 0;
  store.months[key] = blankMonth(monthLabelFromKey(key), prevStartBalance, prevBudget);
  saveData();
  currentMonth = key;
  renderAll();
}

// ---------- Totals ----------
function sumAmounts(list) {
  return list.reduce((s, t) => s + (Number(t.amount) || 0), 0);
}
function categoryTotals(list) {
  const totals = {};
  list.forEach((t) => {
    totals[t.category] = (totals[t.category] || 0) + (Number(t.amount) || 0);
  });
  return totals;
}
function computeMonthTotals(key) {
  const m = store.months[key];
  if (!m) return { plannedExpense: 0, actualExpense: 0, plannedIncome: 0, actualIncome: 0, invested: 0, netSavings: 0, endingBalance: 0 };
  const plannedExpense = Object.values(m.budget.expense || {}).reduce((a, b) => a + (Number(b) || 0), 0);
  const plannedIncome = Object.values(m.budget.income || {}).reduce((a, b) => a + (Number(b) || 0), 0);
  const actualExpense = sumAmounts(m.transactions.expense);
  const actualIncome = sumAmounts(m.transactions.income);
  const invested = sumAmounts(m.investments);
  const netSavings = actualIncome - actualExpense;
  const endingBalance = (Number(m.startingBalance) || 0) + netSavings;
  return { plannedExpense, actualExpense, plannedIncome, actualIncome, invested, netSavings, endingBalance };
}

// ---------- Rendering: header / tabs ----------
function renderMonthSelect() {
  const sel = document.getElementById("monthSelect");
  const keys = sortedMonthKeys();
  sel.innerHTML = keys.map((k) => `<option value="${k}">${store.months[k].label}</option>`).join("");
  if (!currentMonth || !store.months[currentMonth]) currentMonth = keys[keys.length - 1];
  sel.value = currentMonth;
}

function setupTabs() {
  document.querySelectorAll(".tab-btn").forEach((btn) => {
    btn.addEventListener("click", () => {
      document.querySelectorAll(".tab-btn").forEach((b) => b.classList.remove("active"));
      document.querySelectorAll(".tab-panel").forEach((p) => p.classList.remove("active"));
      btn.classList.add("active");
      document.getElementById("tab-" + btn.dataset.tab).classList.add("active");
    });
  });
}

// ---------- Dashboard ----------
function renderDashboard() {
  const totals = computeMonthTotals(currentMonth);
  const cards = document.getElementById("summaryCards");
  const diffExpense = totals.plannedExpense - totals.actualExpense;
  const diffIncome = totals.actualIncome - totals.plannedIncome;
  cards.innerHTML = `
    <div class="card"><div class="label">Starting Balance</div><div class="value">${fmtMoney(store.months[currentMonth].startingBalance)}</div></div>
    <div class="card"><div class="label">Planned Expense</div><div class="value">${fmtMoney(totals.plannedExpense)}</div></div>
    <div class="card"><div class="label">Actual Expense</div><div class="value ${diffExpense < 0 ? "negative" : "positive"}">${fmtMoney(totals.actualExpense)}</div></div>
    <div class="card"><div class="label">Planned Income</div><div class="value">${fmtMoney(totals.plannedIncome)}</div></div>
    <div class="card"><div class="label">Actual Income</div><div class="value ${diffIncome < 0 ? "negative" : "positive"}">${fmtMoney(totals.actualIncome)}</div></div>
    <div class="card"><div class="label">Net Savings this month</div><div class="value ${totals.netSavings < 0 ? "negative" : "positive"}">${fmtMoney(totals.netSavings)}</div></div>
    <div class="card"><div class="label">Ending Balance</div><div class="value">${fmtMoney(totals.endingBalance)}</div></div>
    <div class="card"><div class="label">Invested / Saved</div><div class="value">${fmtMoney(totals.invested)}</div></div>
  `;

  renderBreakdown("expenseBreakdown", "expense");
  renderBreakdown("incomeBreakdown", "income");
  renderInvestmentSummary();
}

function renderBreakdown(containerId, type) {
  const m = store.months[currentMonth];
  const planned = m.budget[type] || {};
  const actuals = categoryTotals(m.transactions[type]);
  const categories = Array.from(new Set([...Object.keys(planned), ...Object.keys(actuals)]));
  const maxVal = Math.max(1, ...categories.map((c) => Math.max(planned[c] || 0, actuals[c] || 0)));
  const el = document.getElementById(containerId);
  if (!categories.length) { el.innerHTML = "<p>No categories yet.</p>"; return; }
  el.innerHTML = categories.map((c) => {
    const p = planned[c] || 0, a = actuals[c] || 0;
    const pWidth = Math.min(100, (p / maxVal) * 100);
    const aWidth = Math.min(100, (a / maxVal) * 100);
    const actualClass = type === "income" ? "income-actual" : "actual";
    return `<div class="bar-item">
      <div class="bar-label"><span>${c}</span><span>Planned ${fmtMoney(p)} · Actual ${fmtMoney(a)}</span></div>
      <div class="bar-track">
        <div class="bar-fill planned" style="width:${pWidth}%"></div>
        <div class="bar-fill ${actualClass}" style="width:${aWidth}%"></div>
      </div>
    </div>`;
  }).join("");
}

function renderInvestmentSummary() {
  const m = store.months[currentMonth];
  const totals = categoryTotals(m.investments);
  const el = document.getElementById("investmentSummary");
  const categories = Object.keys(totals);
  if (!categories.length) { el.innerHTML = "<p>No investments logged for this month.</p>"; return; }
  const maxVal = Math.max(1, ...categories.map((c) => totals[c]));
  el.innerHTML = categories.map((c) => {
    const w = Math.min(100, (totals[c] / maxVal) * 100);
    return `<div class="bar-item">
      <div class="bar-label"><span>${c}</span><span>${fmtMoney(totals[c])}</span></div>
      <div class="bar-track"><div class="bar-fill actual" style="width:${w}%; height:100%; top:0;"></div></div>
    </div>`;
  }).join("");
}

// ---------- Add Entry ----------
function populateCategorySelect() {
  if (!store) return;
  const type = document.querySelector('input[name="entryType"]:checked').value;
  const sel = document.getElementById("entryCategory");
  const list = store.categories[type] || [];
  sel.innerHTML = list.map((c) => `<option value="${c}">${c}</option>`).join("") + `<option value="__new__">+ Add new category...</option>`;
  document.getElementById("newCategoryRow").style.display = "none";
}

function setupAddEntryForm() {
  document.querySelectorAll('input[name="entryType"]').forEach((r) => r.addEventListener("change", populateCategorySelect));
  document.getElementById("entryCategory").addEventListener("change", (e) => {
    document.getElementById("newCategoryRow").style.display = e.target.value === "__new__" ? "flex" : "none";
  });
  document.getElementById("entryDate").value = new Date().toISOString().slice(0, 10);
  populateCategorySelect();

  document.getElementById("entryForm").addEventListener("submit", (e) => {
    e.preventDefault();
    const type = document.querySelector('input[name="entryType"]:checked').value;
    const date = document.getElementById("entryDate").value;
    const amount = parseFloat(document.getElementById("entryAmount").value);
    let category = document.getElementById("entryCategory").value;
    const description = document.getElementById("entryDescription").value.trim();

    if (category === "__new__") {
      category = document.getElementById("newCategoryInput").value.trim();
      if (!category) { alert("Please enter a category name."); return; }
      uniquePush(store.categories[type], category);
    }
    if (!date || isNaN(amount)) { alert("Please fill in date and amount."); return; }

    const monthKey = ensureMonthForDate(date);
    const entry = { id: newId(), date, amount, description, category };
    const m = store.months[monthKey];
    if (type === "investment") m.investments.push(entry);
    else m.transactions[type].push(entry);

    saveData();
    document.getElementById("entryForm").reset();
    document.getElementById("entryDate").value = new Date().toISOString().slice(0, 10);
    populateCategorySelect();
    const msg = document.getElementById("addEntryMsg");
    msg.textContent = `Added to ${store.months[monthKey].label}.`;
    setTimeout(() => (msg.textContent = ""), 3000);
    renderAll();
  });
}

// ---------- Transactions table ----------
function allEntriesFlat() {
  const rows = [];
  sortedMonthKeys().forEach((key) => {
    const m = store.months[key];
    m.transactions.expense.forEach((t) => rows.push(Object.assign({ monthKey: key, type: "expense" }, t)));
    m.transactions.income.forEach((t) => rows.push(Object.assign({ monthKey: key, type: "income" }, t)));
    m.investments.forEach((t) => rows.push(Object.assign({ monthKey: key, type: "investment" }, t)));
  });
  rows.sort((a, b) => (a.date < b.date ? 1 : -1));
  return rows;
}

function populateFilters() {
  const monthSel = document.getElementById("filterMonth");
  const keys = sortedMonthKeys();
  monthSel.innerHTML = `<option value="__all__">All months</option>` + keys.map((k) => `<option value="${k}">${store.months[k].label}</option>`).join("");

  const catSel = document.getElementById("filterCategory");
  const allCats = Array.from(new Set([...store.categories.expense, ...store.categories.income, ...store.categories.investment]));
  catSel.innerHTML = `<option value="__all__">All</option>` + allCats.map((c) => `<option value="${c}">${c}</option>`).join("");
}

function renderTransactionsTable() {
  const monthFilter = document.getElementById("filterMonth").value;
  const typeFilter = document.getElementById("filterType").value;
  const catFilter = document.getElementById("filterCategory").value;
  const dateFrom = document.getElementById("filterDateFrom").value;
  const dateTo = document.getElementById("filterDateTo").value;
  const search = document.getElementById("filterSearch").value.toLowerCase();

  let rows = allEntriesFlat();
  if (monthFilter !== "__all__") rows = rows.filter((r) => r.monthKey === monthFilter);
  if (typeFilter !== "__all__") rows = rows.filter((r) => r.type === typeFilter);
  if (catFilter !== "__all__") rows = rows.filter((r) => r.category === catFilter);
  if (dateFrom) rows = rows.filter((r) => r.date >= dateFrom);
  if (dateTo) rows = rows.filter((r) => r.date <= dateTo);
  if (search) rows = rows.filter((r) => (r.description || "").toLowerCase().includes(search));

  const tbody = document.querySelector("#transactionsTable tbody");
  if (!rows.length) { tbody.innerHTML = `<tr><td colspan="6">No entries found.</td></tr>`; return; }
  tbody.innerHTML = rows.map((r) => `
    <tr>
      <td>${r.date}</td>
      <td><span class="tag ${r.type}">${r.type}</span></td>
      <td>${r.category}</td>
      <td>${r.description || ""}</td>
      <td>${fmtMoney(r.amount)}</td>
      <td class="row-actions">
        <button class="edit-btn" data-id="${r.id}" data-month="${r.monthKey}" data-type="${r.type}">Edit</button>
        <button class="delete-btn" data-id="${r.id}" data-month="${r.monthKey}" data-type="${r.type}">Delete</button>
      </td>
    </tr>`).join("");

  tbody.querySelectorAll(".delete-btn").forEach((btn) => btn.addEventListener("click", () => {
    if (!confirm("Delete this entry?")) return;
    deleteEntry(btn.dataset.type, btn.dataset.month, btn.dataset.id);
  }));
  tbody.querySelectorAll(".edit-btn").forEach((btn) => btn.addEventListener("click", () => {
    openEditPrompt(btn.dataset.type, btn.dataset.month, btn.dataset.id);
  }));
}

function getEntryList(type, monthKey) {
  const m = store.months[monthKey];
  if (type === "investment") return m.investments;
  return m.transactions[type];
}

function deleteEntry(type, monthKey, id) {
  const list = getEntryList(type, monthKey);
  const idx = list.findIndex((e) => e.id === id);
  if (idx > -1) list.splice(idx, 1);
  saveData();
  renderAll();
}

function openEditPrompt(type, monthKey, id) {
  const list = getEntryList(type, monthKey);
  const entry = list.find((e) => e.id === id);
  if (!entry) return;
  const newAmount = prompt("Amount (₹):", entry.amount);
  if (newAmount === null) return;
  const newDesc = prompt("Description:", entry.description || "");
  if (newDesc === null) return;
  const newCategory = prompt("Category:", entry.category);
  if (newCategory === null) return;
  const newDate = prompt("Date (YYYY-MM-DD):", entry.date);
  if (newDate === null) return;

  entry.amount = parseFloat(newAmount) || 0;
  entry.description = newDesc;
  entry.category = newCategory;

  if (newDate !== entry.date && monthKeyFromDate(newDate) !== monthKey) {
    list.splice(list.indexOf(entry), 1);
    const newKey = ensureMonthForDate(newDate);
    entry.date = newDate;
    if (type === "investment") store.months[newKey].investments.push(entry);
    else store.months[newKey].transactions[type].push(entry);
  } else {
    entry.date = newDate;
  }
  uniquePush(store.categories[type], newCategory);
  saveData();
  renderAll();
}

// ---------- Budget ----------
function renderBudgetTab() {
  const m = store.months[currentMonth];
  document.getElementById("budgetMonthLabel").textContent = m.label;
  document.getElementById("startingBalanceInput").value = m.startingBalance;

  const renderList = (containerId, type) => {
    const cats = store.categories[type];
    document.getElementById(containerId).innerHTML = cats.map((c) => `
      <div class="budget-item">
        <label>${c}</label>
        <input type="number" step="0.01" min="0" data-cat="${c}" data-type="${type}" value="${m.budget[type][c] || 0}" />
      </div>`).join("");
  };
  renderList("budgetExpenseList", "expense");
  renderList("budgetIncomeList", "income");
}

function setupBudgetTab() {
  document.getElementById("copyPrevBudgetBtn").addEventListener("click", () => {
    const keys = sortedMonthKeys();
    const idx = keys.indexOf(currentMonth);
    if (idx <= 0) { alert("No previous month to copy from."); return; }
    const prev = store.months[keys[idx - 1]];
    const m = store.months[currentMonth];
    m.budget.expense = Object.assign({}, prev.budget.expense);
    m.budget.income = Object.assign({}, prev.budget.income);
    renderBudgetTab();
  });

  document.getElementById("saveBudgetBtn").addEventListener("click", () => {
    const m = store.months[currentMonth];
    m.startingBalance = parseFloat(document.getElementById("startingBalanceInput").value) || 0;
    document.querySelectorAll("#budgetExpenseList input, #budgetIncomeList input").forEach((inp) => {
      m.budget[inp.dataset.type][inp.dataset.cat] = parseFloat(inp.value) || 0;
    });
    saveData();
    const msg = document.getElementById("saveBudgetMsg");
    msg.textContent = "Saved!";
    setTimeout(() => (msg.textContent = ""), 2000);
    renderAll();
  });

  document.getElementById("deleteMonthBtn").addEventListener("click", deleteCurrentMonth);
}

function deleteCurrentMonth() {
  const keys = sortedMonthKeys();
  if (keys.length <= 1) { alert("You can't delete the only remaining month."); return; }
  const m = store.months[currentMonth];
  const entryCount = m.transactions.expense.length + m.transactions.income.length + m.investments.length;
  if (entryCount > 0) {
    alert(`Can't delete ${m.label}: it still has ${entryCount} entr${entryCount === 1 ? "y" : "ies"} (expenses/income/investments). Remove or move those first in the Transactions tab.`);
    return;
  }
  if (!confirm(`Delete ${m.label}? It has no entries, only its empty budget will be removed.`)) return;
  delete store.months[currentMonth];
  currentMonth = sortedMonthKeys().pop();
  saveData();
  renderAll();
}

// ---------- History ----------
function exportHistoryPdf() {
  const dateEl = document.getElementById("historyPrintDate");
  if (dateEl) dateEl.textContent = "Generated on " + new Date().toLocaleString("en-IN");
  const cleanup = () => window.removeEventListener("afterprint", cleanup);
  window.addEventListener("afterprint", cleanup);
  window.print();
}

function renderHistory() {
  const keys = sortedMonthKeys();
  const data = keys.map((k) => ({ key: k, label: store.months[k].label, totals: computeMonthTotals(k) }));

  const tbody = document.querySelector("#historyTable tbody");
  tbody.innerHTML = data.map((d) => `
    <tr>
      <td>${d.label}</td>
      <td>${fmtMoney(d.totals.plannedExpense)}</td>
      <td>${fmtMoney(d.totals.actualExpense)}</td>
      <td>${fmtMoney(d.totals.plannedIncome)}</td>
      <td>${fmtMoney(d.totals.actualIncome)}</td>
      <td>${fmtMoney(d.totals.netSavings)}</td>
      <td>${fmtMoney(d.totals.invested)}</td>
    </tr>`).join("");

  const grand = data.reduce((acc, d) => {
    acc.plannedExpense += d.totals.plannedExpense;
    acc.actualExpense += d.totals.actualExpense;
    acc.plannedIncome += d.totals.plannedIncome;
    acc.actualIncome += d.totals.actualIncome;
    acc.netSavings += d.totals.netSavings;
    acc.invested += d.totals.invested;
    return acc;
  }, { plannedExpense: 0, actualExpense: 0, plannedIncome: 0, actualIncome: 0, netSavings: 0, invested: 0 });

  tbody.innerHTML += `
    <tr class="totals-row">
      <td>Total</td>
      <td>${fmtMoney(grand.plannedExpense)}</td>
      <td>${fmtMoney(grand.actualExpense)}</td>
      <td>${fmtMoney(grand.plannedIncome)}</td>
      <td>${fmtMoney(grand.actualIncome)}</td>
      <td>${fmtMoney(grand.netSavings)}</td>
      <td>${fmtMoney(grand.invested)}</td>
    </tr>`;

  drawHistoryChart(data);
}

function drawHistoryChart(data) {
  const canvas = document.getElementById("historyChart");
  const ctx = canvas.getContext("2d");
  const W = canvas.width, H = canvas.height;
  ctx.clearRect(0, 0, W, H);
  if (!data.length) { ctx.fillText("No data yet.", 20, 30); return; }

  const padding = { left: 60, right: 20, top: 32, bottom: 40 };
  const chartW = W - padding.left - padding.right;
  const chartH = H - padding.top - padding.bottom;
  const maxVal = Math.max(1, ...data.flatMap((d) => [d.totals.actualExpense, d.totals.actualIncome, d.totals.invested]));
  const groupW = chartW / data.length;
  const barW = Math.min(22, groupW / 5);
  const series = [
    { key: "actualExpense", color: "#c0504d" },
    { key: "actualIncome", color: "#2563eb" },
    { key: "invested", color: "#2f6f4e" }
  ];

  // axes
  ctx.strokeStyle = "#ccc";
  ctx.beginPath();
  ctx.moveTo(padding.left, padding.top);
  ctx.lineTo(padding.left, padding.top + chartH);
  ctx.lineTo(padding.left + chartW, padding.top + chartH);
  ctx.stroke();

  ctx.fillStyle = "#6b7670";
  ctx.font = "11px Segoe UI";
  for (let i = 0; i <= 4; i++) {
    const val = (maxVal / 4) * i;
    const y = padding.top + chartH - (chartH * i) / 4;
    ctx.fillText("₹" + Math.round(val).toLocaleString("en-IN"), 2, y + 4);
    ctx.strokeStyle = "#eee";
    ctx.beginPath(); ctx.moveTo(padding.left, y); ctx.lineTo(padding.left + chartW, y); ctx.stroke();
  }

  data.forEach((d, i) => {
    const groupX = padding.left + i * groupW + groupW / 2;
    const groupStartX = groupX - (series.length * barW) / 2 - (series.length - 1);

    series.forEach((s, si) => {
      const val = d.totals[s.key];
      const barH = (val / maxVal) * chartH;
      const barX = groupStartX + si * (barW + 1);
      ctx.fillStyle = s.color;
      ctx.fillRect(barX, padding.top + chartH - barH, barW, barH);

      ctx.textAlign = "center";
      ctx.font = "9px Segoe UI";
      ctx.fillText(fmtCompactINR(val), barX + barW / 2, padding.top + chartH - barH - 4);
    });

    ctx.fillStyle = "#1f2a24";
    ctx.font = "11px Segoe UI";
    ctx.textAlign = "center";
    ctx.fillText(d.label.split(" ")[0], groupX, padding.top + chartH + 16);
    ctx.textAlign = "left";
  });

  // legend
  ctx.fillStyle = "#c0504d"; ctx.fillRect(padding.left, 2, 10, 10);
  ctx.fillStyle = "#1f2a24"; ctx.fillText("Expense", padding.left + 14, 11);
  ctx.fillStyle = "#2563eb"; ctx.fillRect(padding.left + 90, 2, 10, 10);
  ctx.fillStyle = "#1f2a24"; ctx.fillText("Income", padding.left + 104, 11);
  ctx.fillStyle = "#2f6f4e"; ctx.fillRect(padding.left + 170, 2, 10, 10);
  ctx.fillStyle = "#1f2a24"; ctx.fillText("Invested", padding.left + 184, 11);
}

// ---------- Export / Import ----------
function exportCsv() {
  const rows = allEntriesFlat();
  const header = ["Month", "Date", "Type", "Category", "Description", "Amount (INR)"];
  const lines = [header.join(",")];
  rows.forEach((r) => {
    const cells = [store.months[r.monthKey].label, r.date, r.type, r.category, (r.description || "").replace(/"/g, '""'), r.amount];
    lines.push(cells.map((c) => `"${c}"`).join(","));
  });
  downloadFile("budget-tracker-export.csv", lines.join("\n"), "text/csv");
}

function exportJson() {
  downloadFile("budget-tracker-backup.json", JSON.stringify(store, null, 2), "application/json");
}

function downloadFile(filename, content, mime) {
  const blob = new Blob([content], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

function setupDataTab() {
  document.getElementById("exportCsvBtn").addEventListener("click", exportCsv);
  document.getElementById("exportJsonBtn").addEventListener("click", exportJson);
  document.getElementById("importJsonInput").addEventListener("change", (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const parsed = JSON.parse(reader.result);
        if (!parsed.months || !parsed.categories) throw new Error("Invalid backup file.");
        store = parsed;
        saveData();
        currentMonth = null;
        renderAll();
        showDataMsg("Backup imported successfully.");
      } catch (err) {
        alert("Could not import file: " + err.message);
      }
    };
    reader.readAsText(file);
  });
  document.getElementById("resetBtn").addEventListener("click", () => {
    if (!confirm("This will erase your cloud data and reload the original Excel data. Continue?")) return;
    seedFromExcelData();
    saveData();
    currentMonth = null;
    renderAll();
    showDataMsg("Data reset to original seed.");
  });
}

function showDataMsg(text) {
  const el = document.getElementById("dataMsg");
  el.textContent = text;
  setTimeout(() => (el.textContent = ""), 3000);
}

// ---------- Main render ----------
function renderAll() {
  renderMonthSelect();
  renderDashboard();
  populateFilters();
  renderTransactionsTable();
  renderBudgetTab();
  renderHistory();
  populateCategorySelect();
}

function init() {
  setupTabs();
  setupAddEntryForm();
  setupBudgetTab();
  setupDataTab();
  setupAuthUI();

  document.getElementById("monthSelect").addEventListener("change", (e) => {
    currentMonth = e.target.value;
    renderDashboard();
    renderBudgetTab();
  });
  document.getElementById("addMonthBtn").addEventListener("click", addNewMonth);
  ["filterMonth", "filterType", "filterCategory", "filterDateFrom", "filterDateTo"].forEach((id) =>
    document.getElementById(id).addEventListener("change", renderTransactionsTable)
  );
  document.getElementById("filterSearch").addEventListener("input", renderTransactionsTable);
  document.getElementById("applyFiltersBtn").addEventListener("click", renderTransactionsTable);
  document.getElementById("clearFiltersBtn").addEventListener("click", () => {
    document.getElementById("filterMonth").value = "__all__";
    document.getElementById("filterType").value = "__all__";
    document.getElementById("filterCategory").value = "__all__";
    document.getElementById("filterDateFrom").value = "";
    document.getElementById("filterDateTo").value = "";
    document.getElementById("filterSearch").value = "";
    renderTransactionsTable();
  });
  document.getElementById("exportHistoryPdfBtn").addEventListener("click", exportHistoryPdf);

  initFirebaseAuth();
}

function setupAuthUI() {
  document.getElementById("loginForm").addEventListener("submit", (e) => {
    e.preventDefault();
    const email = document.getElementById("loginEmail").value.trim();
    const password = document.getElementById("loginPassword").value;
    const errEl = document.getElementById("loginError");
    errEl.textContent = "";
    auth.signInWithEmailAndPassword(email, password).catch((err) => {
      errEl.textContent = err.message;
    });
  });
  document.getElementById("signOutBtn").addEventListener("click", () => auth.signOut());
  document.getElementById("userEmailLabel").addEventListener("click", renameUser);
}

document.addEventListener("DOMContentLoaded", init);
