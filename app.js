/**
 * Checks & Balances (v2)
 * - Planned budget (expenses)
 * - Actual spending ledger filtered by month
 * - Budget vs Actual by category + alerts
 * - Debt payoff planner (avalanche/snowball) with feasibility detection
 * - PDF + Excel exports
 * - LocalStorage persistence with versioning + throttled writes
 */

const APP_VERSION = 2;
const LS_KEY = "checks_balances_v2";

/* ------------------------------ Categories ------------------------------ */

const CATEGORIES = [
  "Housing",
  "Utilities",
  "Food",
  "Transportation",
  "Insurance",
  "Health",
  "Shopping",
  "Entertainment",
  "Subscriptions",
  "Childcare",
  "Savings",
  "Debt Minimums",
  "Other"
];

/* ------------------------------ Utilities ------------------------------ */

const moneyFmt = new Intl.NumberFormat(undefined, { style: "currency", currency: "USD" });

function money(n) {
  const v = Number(n);
  return Number.isFinite(v) ? moneyFmt.format(v) : moneyFmt.format(0);
}

function num(x) {
  const v = Number(x);
  return Number.isFinite(v) ? v : 0;
}

function clamp(n, min, max) {
  return Math.max(min, Math.min(max, n));
}

function uid(prefix = "id") {
  return `${prefix}_${Math.random().toString(16).slice(2)}_${Date.now().toString(16)}`;
}

function todayISO() {
  return new Date().toISOString().slice(0, 10);
}

function currentMonthISO() {
  return new Date().toISOString().slice(0, 7); // YYYY-MM
}

function paymentsPerMonth(freq) {
  switch (freq) {
    case "monthly": return 1;
    case "semimonthly": return 2;
    case "biweekly": return 26 / 12;
    case "weekly": return 52 / 12;
    default: return 1;
  }
}

function monthFromDateISO(dateISO) {
  // dateISO: YYYY-MM-DD -> YYYY-MM
  return String(dateISO || "").slice(0, 7);
}

/* ------------------------------ State ------------------------------ */

function defaultState() {
  return {
    version: APP_VERSION,
    ledgerMonth: currentMonthISO(),
    income: {
      frequency: "biweekly",
      grossPerPaycheck: 0,
      taxRatePct: 20,
      otherDeductionsPerPaycheck: 0,
      otherMonthlyIncome: 0
    },
    expenses: [], // planned monthly budget items
    spending: [], // ledger transactions
    debts: [],
    payoff: {
      strategy: "avalanche",
      extraPayment: 0
    }
  };
}

let state = defaultState();

/* ------------------------------ Persistence ------------------------------ */

function safeParseJSON(s) {
  try { return JSON.parse(s); } catch { return null; }
}

function migrateState(raw) {
  // Future-proofing: accept older shapes safely
  if (!raw || typeof raw !== "object") return defaultState();
  const s = defaultState();

  // Copy known keys with fallbacks
  s.ledgerMonth = typeof raw.ledgerMonth === "string" ? raw.ledgerMonth : s.ledgerMonth;

  if (raw.income && typeof raw.income === "object") {
    s.income.frequency = raw.income.frequency ?? s.income.frequency;
    s.income.grossPerPaycheck = num(raw.income.grossPerPaycheck);
    s.income.taxRatePct = clamp(num(raw.income.taxRatePct), 0, 60);
    s.income.otherDeductionsPerPaycheck = num(raw.income.otherDeductionsPerPaycheck);
    s.income.otherMonthlyIncome = num(raw.income.otherMonthlyIncome);
  }

  s.expenses = Array.isArray(raw.expenses) ? raw.expenses.map(e => ({
    id: e.id || uid("exp"),
    name: String(e.name || "").trim(),
    category: CATEGORIES.includes(e.category) ? e.category : "Other",
    amount: num(e.amount)
  })).filter(e => e.name && e.amount > 0) : [];

  s.spending = Array.isArray(raw.spending) ? raw.spending.map(t => ({
    id: t.id || uid("txn"),
    date: (String(t.date || "").slice(0, 10) || todayISO()),
    category: CATEGORIES.includes(t.category) ? t.category : "Other",
    description: String(t.description || "").trim(),
    amount: num(t.amount)
  })).filter(t => t.amount > 0) : [];

  s.debts = Array.isArray(raw.debts) ? raw.debts.map(d => ({
    id: d.id || uid("debt"),
    name: String(d.name || "").trim(),
    balance: num(d.balance),
    aprPct: Math.max(0, num(d.aprPct)),
    minPayment: Math.max(0, num(d.minPayment))
  })).filter(d => d.name && d.balance > 0) : [];

  if (raw.payoff && typeof raw.payoff === "object") {
    s.payoff.strategy = raw.payoff.strategy === "snowball" ? "snowball" : "avalanche";
    s.payoff.extraPayment = Math.max(0, num(raw.payoff.extraPayment));
  }

  return s;
}

function load() {
  const raw = safeParseJSON(localStorage.getItem(LS_KEY) || "");
  state = migrateState(raw);
}

const saveThrottled = (() => {
  let timeout = null;
  return () => {
    if (timeout) return;
    timeout = setTimeout(() => {
      timeout = null;
      localStorage.setItem(LS_KEY, JSON.stringify(state));
    }, 250);
  };
})();

/* ------------------------------ Core Calculations ------------------------------ */

function calcIncomeMonthly(inc) {
  // Step-by-step:
  // 1) monthlyGross = grossPerPaycheck * paychecksPerMonth
  // 2) taxes = monthlyGross * (taxRatePct / 100)
  // 3) deductions = otherDeductionsPerPaycheck * paychecksPerMonth
  // 4) monthlyNet = monthlyGross - taxes - deductions + otherMonthlyIncome
  const ppm = paymentsPerMonth(inc.frequency);
  const monthlyGross = num(inc.grossPerPaycheck) * ppm;
  const taxes = monthlyGross * (num(inc.taxRatePct) / 100);
  const deductions = num(inc.otherDeductionsPerPaycheck) * ppm;
  const monthlyNet = monthlyGross - taxes - deductions + num(inc.otherMonthlyIncome);
  return { ppm, monthlyGross, taxes, deductions, monthlyNet };
}

function sumPlannedExpenses(expenses) {
  return expenses.reduce((a, e) => a + num(e.amount), 0);
}

function spendingForMonth(spending, ledgerMonth) {
  return spending.filter(t => monthFromDateISO(t.date) === ledgerMonth);
}

function sumSpending(transactions) {
  return transactions.reduce((a, t) => a + num(t.amount), 0);
}

function groupByCategory(items, getAmount) {
  const map = new Map();
  for (const it of items) {
    const cat = it.category || "Other";
    map.set(cat, (map.get(cat) || 0) + getAmount(it));
  }
  return map;
}

function budgetVsActual(expenses, spendingMonthTxns) {
  const planned = groupByCategory(expenses, e => num(e.amount));
  const actual = groupByCategory(spendingMonthTxns, t => num(t.amount));

  const rows = CATEGORIES.map(cat => {
    const p = planned.get(cat) || 0;
    const a = actual.get(cat) || 0;
    const remaining = p - a;
    const pctUsed = p <= 0 ? (a > 0 ? 100 : 0) : (a / p) * 100;
    let status = "OK";
    if (p > 0 && a > p) status = "Over";
    else if (p > 0 && pctUsed >= 90) status = "Near";
    else if (p === 0 && a > 0) status = "Over"; // spending without a plan
    return { category: cat, planned: p, actual: a, remaining, pctUsed, status };
  }).filter(r => r.planned > 0 || r.actual > 0);

  // Alerts: include only “Over” or “Near”
  const alerts = rows
    .filter(r => r.status === "Over" || r.status === "Near")
    .sort((a, b) => (b.actual - b.planned) - (a.actual - a.planned));

  return { rows, alerts };
}

/* ------------------------------ Debt Payoff Engine ------------------------------ */

function buildPayoffPlan(debtsInput, strategy, extraPayment) {
  const debts = debtsInput
    .map(d => ({
      id: d.id,
      name: String(d.name || "").trim() || "Debt",
      balance: Math.max(0, num(d.balance)),
      aprPct: Math.max(0, num(d.aprPct)),
      minPayment: Math.max(0, num(d.minPayment))
    }))
    .filter(d => d.balance > 0);

  if (debts.length === 0) return { months: 0, payoffLabel: "—", schedule: [] };

  const baseExtra = Math.max(0, num(extraPayment));
  const maxMonths = 600;

  const sortActive = () => {
    const active = debts.filter(d => d.balance > 0);
    if (strategy === "snowball") active.sort((a, b) => a.balance - b.balance || b.aprPct - a.aprPct);
    else active.sort((a, b) => b.aprPct - a.aprPct || a.balance - b.balance);
    return active;
  };

  const schedule = [];
  let consecutiveNoProgress = 0;

  for (let m = 1; m <= maxMonths; m++) {
    const active = sortActive();
    if (active.length === 0) return { months: m - 1, payoffLabel: `${m - 1} month(s)`, schedule };

    // 1) apply interest
    let interestTotal = 0;
    for (const d of active) {
      const monthlyRate = (d.aprPct / 100) / 12;
      const interest = d.balance * monthlyRate;
      d.balance += interest;
      interestTotal += interest;
    }

    // 2) pay minimums
    let paidTotal = 0;
    for (const d of active) {
      const pay = Math.min(d.minPayment, d.balance);
      d.balance -= pay;
      paidTotal += pay;
    }

    // 3) apply extra to target(s) in strategy order
    let remainingExtra = baseExtra;

    const cascade = sortActive(); // re-sort after min payments
    for (const d of cascade) {
      if (remainingExtra <= 0) break;
      if (d.balance <= 0) continue;
      const pay = Math.min(remainingExtra, d.balance);
      d.balance -= pay;
      remainingExtra -= pay;
      paidTotal += pay;
    }

    const principalTotal = paidTotal - interestTotal;
    const totalBalanceRemaining = debts.reduce((a, d) => a + d.balance, 0);
    const targetName = sortActive()[0]?.name || active[0].name;

    schedule.push({
      month: m,
      target: targetName,
      paid: paidTotal,
      interest: interestTotal,
      principal: principalTotal,
      totalBalanceRemaining
    });

    // Feasibility detection: if principal is not positive for 3 consecutive months, it likely will not pay down.
    if (principalTotal <= 0.01) consecutiveNoProgress++;
    else consecutiveNoProgress = 0;

    if (consecutiveNoProgress >= 3) {
      return { months: null, payoffLabel: "Not feasible (payments not reducing principal)", schedule };
    }
  }

  return { months: null, payoffLabel: "Over limit / Not feasible", schedule };
}

/* ------------------------------ DOM Helpers ------------------------------ */

function $(id) {
  return document.getElementById(id);
}

function setText(id, text) {
  $(id).textContent = text;
}

function clear(el) {
  while (el.firstChild) el.removeChild(el.firstChild);
}

function td(text, className) {
  const cell = document.createElement("td");
  if (className) cell.className = className;
  cell.textContent = text;
  return cell;
}

function btn(label, attrs = {}) {
  const b = document.createElement("button");
  b.className = "btn secondary";
  b.type = "button";
  b.textContent = label;
  for (const [k, v] of Object.entries(attrs)) b.setAttribute(k, v);
  return b;
}

/* ------------------------------ Rendering ------------------------------ */

function renderCategorySelects() {
  const expSel = $("expenseCategory");
  const spendSel = $("spendCategory");

  clear(expSel);
  clear(spendSel);

  for (const cat of CATEGORIES) {
    const o1 = document.createElement("option");
    o1.value = cat; o1.textContent = cat;
    expSel.appendChild(o1);

    const o2 = document.createElement("option");
    o2.value = cat; o2.textContent = cat;
    spendSel.appendChild(o2);
  }
}

function renderInputs() {
  $("incomeFrequency").value = state.income.frequency;
  $("grossPerPaycheck").value = String(state.income.grossPerPaycheck ?? 0);
  $("taxRatePct").value = String(state.income.taxRatePct ?? 20);
  $("otherDeductionsPerPaycheck").value = String(state.income.otherDeductionsPerPaycheck ?? 0);
  $("otherMonthlyIncome").value = String(state.income.otherMonthlyIncome ?? 0);

  $("ledgerMonth").value = state.ledgerMonth;
  $("spendDate").value = todayISO();

  $("payoffStrategy").value = state.payoff.strategy;
  $("extraPayment").value = String(state.payoff.extraPayment ?? 0);
}

function renderKPIs() {
  const inc = calcIncomeMonthly(state.income);
  const plannedTotal = sumPlannedExpenses(state.expenses);

  const monthTxns = spendingForMonth(state.spending, state.ledgerMonth);
  const actualTotal = sumSpending(monthTxns);

  const cashLeft = inc.monthlyNet - plannedTotal;
  const variance = plannedTotal - actualTotal;

  const debtTotal = state.debts.reduce((a, d) => a + num(d.balance), 0);

  setText("kpiMonthlyGross", money(inc.monthlyGross));
  setText("kpiMonthlyNet", money(inc.monthlyNet));
  setText("kpiPlannedExpenses", money(plannedTotal));
  setText("kpiCashLeft", money(cashLeft));
  setText("kpiActualSpending", money(actualTotal));
  setText("kpiVariance", money(variance));
  setText("kpiDebtTotal", money(debtTotal));
}

function renderExpensesTable() {
  const tbody = $("expenseTbody");
  clear(tbody);

  for (const e of state.expenses) {
    const tr = document.createElement("tr");
    tr.appendChild(td(e.name));
    tr.appendChild(td(e.category));
    tr.appendChild(td(money(e.amount), "right"));

    const actions = document.createElement("td");
    actions.className = "right";
    actions.appendChild(btn("Remove", { "data-action": "delete-expense", "data-id": e.id }));
    tr.appendChild(actions);

    tbody.appendChild(tr);
  }
}

function renderSpendingTable() {
  const tbody = $("spendTbody");
  clear(tbody);

  const monthTxns = spendingForMonth(state.spending, state.ledgerMonth)
    .slice()
    .sort((a, b) => (a.date || "").localeCompare(b.date || ""));

  for (const t of monthTxns) {
    const tr = document.createElement("tr");
    tr.appendChild(td(t.date));
    tr.appendChild(td(t.category));
    tr.appendChild(td(t.description));
    tr.appendChild(td(money(t.amount), "right"));

    const actions = document.createElement("td");
    actions.className = "right";
    actions.appendChild(btn("Remove", { "data-action": "delete-spending", "data-id": t.id }));
    tr.appendChild(actions);

    tbody.appendChild(tr);
  }
}

function renderBudgetVsActual() {
  const tbody = $("bvaTbody");
  const alertsEl = $("alerts");
  clear(tbody);
  clear(alertsEl);

  const monthTxns = spendingForMonth(state.spending, state.ledgerMonth);
  const { rows, alerts } = budgetVsActual(state.expenses, monthTxns);

  // Alerts
  if (alerts.length === 0) {
    const a = document.createElement("div");
    a.className = "alert";
    a.textContent = "No alerts for this month based on your planned budget.";
    alertsEl.appendChild(a);
  } else {
    for (const r of alerts.slice(0, 6)) {
      const div = document.createElement("div");
      div.className = `alert ${r.status === "Over" ? "over" : "near"}`;

      const badge = document.createElement("span");
      badge.className = `badge ${r.status === "Over" ? "over" : "near"}`;
      badge.textContent = r.status.toUpperCase();

      const text = document.createElement("span");
      if (r.planned <= 0) {
        text.textContent = `${r.category}: ${money(r.actual)} spent with no planned budget set.`;
      } else if (r.status === "Over") {
        text.textContent = `${r.category}: Over budget by ${money(r.actual - r.planned)} (planned ${money(r.planned)}, actual ${money(r.actual)}).`;
      } else {
        text.textContent = `${r.category}: ${r.pctUsed.toFixed(0)}% used (planned ${money(r.planned)}, actual ${money(r.actual)}).`;
      }

      div.appendChild(badge);
      div.appendChild(text);
      alertsEl.appendChild(div);
    }
  }

  // Table
  for (const r of rows) {
    const tr = document.createElement("tr");
    tr.appendChild(td(r.category));
    tr.appendChild(td(money(r.planned), "right"));
    tr.appendChild(td(money(r.actual), "right"));
    tr.appendChild(td(money(r.remaining), "right"));
    tr.appendChild(td(`${r.pctUsed.toFixed(0)}%`, "right"));

    const statusCell = document.createElement("td");
    const b = document.createElement("span");
    const badgeClass = r.status === "Over" ? "over" : (r.status === "Near" ? "near" : "ok");
    b.className = `badge ${badgeClass}`;
    b.textContent = r.status;
    statusCell.appendChild(b);

    tr.appendChild(statusCell);
    tbody.appendChild(tr);
  }
}

function renderDebtsTableAndPlan() {
  const tbody = $("debtTbody");
  clear(tbody);

  for (const d of state.debts) {
    const tr = document.createElement("tr");
    tr.appendChild(td(d.name));
    tr.appendChild(td(money(d.balance), "right"));
    tr.appendChild(td(`${num(d.aprPct).toFixed(2)}%`, "right"));
    tr.appendChild(td(money(d.minPayment), "right"));

    const actions = document.createElement("td");
    actions.className = "right";
    actions.appendChild(btn("Remove", { "data-action": "delete-debt", "data-id": d.id }));
    tr.appendChild(actions);

    tbody.appendChild(tr);
  }

  const plan = buildPayoffPlan(state.debts, state.payoff.strategy, state.payoff.extraPayment);
  setText("kpiPayoffTime", plan.payoffLabel);

  const ptbody = $("payoffTbody");
  clear(ptbody);

  for (const r of plan.schedule.slice(0, 240)) {
    const tr = document.createElement("tr");
    tr.appendChild(td(String(r.month)));
    tr.appendChild(td(r.target));
    tr.appendChild(td(money(r.paid), "right"));
    tr.appendChild(td(money(r.interest), "right"));
    tr.appendChild(td(money(r.principal), "right"));
    tr.appendChild(td(money(r.totalBalanceRemaining), "right"));
    ptbody.appendChild(tr);
  }
}

let renderQueued = false;
function scheduleRender() {
  if (renderQueued) return;
  renderQueued = true;
  requestAnimationFrame(() => {
    renderQueued = false;
    renderAll();
  });
}

function renderAll() {
  renderKPIs();
  renderExpensesTable();
  renderSpendingTable();
  renderBudgetVsActual();
  renderDebtsTableAndPlan();
  saveThrottled();
}

/* ------------------------------ Mutations ------------------------------ */

function addExpense() {
  const name = $("expenseName").value.trim();
  const category = $("expenseCategory").value;
  const amount = num($("expenseAmount").value);

  if (!name || amount <= 0) return;

  state.expenses.push({ id: uid("exp"), name, category, amount });
  $("expenseName").value = "";
  $("expenseAmount").value = "";
  scheduleRender();
}

function addSpending() {
  const date = ($("spendDate").value || todayISO());
  const category = $("spendCategory").value;
  const description = $("spendDesc").value.trim();
  const amount = num($("spendAmount").value);

  if (amount <= 0) return;

  state.spending.push({ id: uid("txn"), date, category, description, amount });
  $("spendDesc").value = "";
  $("spendAmount").value = "";
  scheduleRender();
}

function addDebt() {
  const name = $("debtName").value.trim();
  const balance = num($("debtBalance").value);
  const aprPct = Math.max(0, num($("debtApr").value));
  const minPayment = Math.max(0, num($("debtMin").value));

  if (!name || balance <= 0) return;

  state.debts.push({ id: uid("debt"), name, balance, aprPct, minPayment });
  $("debtName").value = "";
  $("debtBalance").value = "";
  $("debtApr").value = "";
  $("debtMin").value = "";
  scheduleRender();
}

function deleteById(collection, id) {
  const idx = collection.findIndex(x => x.id === id);
  if (idx >= 0) collection.splice(idx, 1);
}

/* ------------------------------ Exports ------------------------------ */

async function exportPdf() {
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ unit: "pt", format: "letter" });

  const inc = calcIncomeMonthly(state.income);
  const plannedTotal = sumPlannedExpenses(state.expenses);
  const monthTxns = spendingForMonth(state.spending, state.ledgerMonth);
  const actualTotal = sumSpending(monthTxns);

  const cashLeft = inc.monthlyNet - plannedTotal;
  const variance = plannedTotal - actualTotal;

  const debtTotal = state.debts.reduce((a, d) => a + num(d.balance), 0);
  const plan = buildPayoffPlan(state.debts, state.payoff.strategy, state.payoff.extraPayment);

  const bva = budgetVsActual(state.expenses, monthTxns);

  doc.setFont("helvetica", "bold");
  doc.setFontSize(16);
  doc.text("Checks & Balances – Report", 40, 50);

  doc.setFont("helvetica", "normal");
  doc.setFontSize(10);
  doc.text(`Generated: ${new Date().toLocaleString()}`, 40, 68);
  doc.text(`Ledger Month: ${state.ledgerMonth}`, 40, 82);

  doc.setFontSize(12);
  doc.text("Summary", 40, 108);

  const summaryRows = [
    ["Monthly Gross (est.)", money(inc.monthlyGross)],
    ["Monthly Take-Home (est.)", money(inc.monthlyNet)],
    ["Planned Expenses", money(plannedTotal)],
    ["Actual Spending (ledger)", money(actualTotal)],
    ["Variance (Planned − Actual)", money(variance)],
    ["Cash Left (Net − Planned)", money(cashLeft)],
    ["Total Debt", money(debtTotal)],
    ["Debt Strategy", state.payoff.strategy],
    ["Extra Monthly Debt Payment", money(state.payoff.extraPayment)],
    ["Estimated Payoff Time", plan.payoffLabel]
  ];

  doc.autoTable({
    startY: 118,
    head: [["Metric", "Value"]],
    body: summaryRows,
    styles: { fontSize: 9 },
    headStyles: { fillColor: [20, 22, 29] }
  });

  let y = doc.lastAutoTable.finalY + 16;

  doc.setFontSize(12);
  doc.text("Planned Expenses", 40, y);
  y += 8;

  doc.autoTable({
    startY: y,
    head: [["Name", "Category", "Amount"]],
    body: state.expenses.map(e => [e.name, e.category, money(e.amount)]),
    styles: { fontSize: 9 },
    headStyles: { fillColor: [20, 22, 29] }
  });

  y = doc.lastAutoTable.finalY + 16;

  doc.setFontSize(12);
  doc.text(`Actual Spending (Ledger: ${state.ledgerMonth})`, 40, y);
  y += 8;

  doc.autoTable({
    startY: y,
    head: [["Date", "Category", "Description", "Amount"]],
    body: monthTxns
      .slice()
      .sort((a, b) => (a.date || "").localeCompare(b.date || ""))
      .map(t => [t.date, t.category, t.description, money(t.amount)]),
    styles: { fontSize: 9 },
    headStyles: { fillColor: [20, 22, 29] }
  });

  y = doc.lastAutoTable.finalY + 16;

  doc.setFontSize(12);
  doc.text("Budget vs Actual (by category)", 40, y);
  y += 8;

  doc.autoTable({
    startY: y,
    head: [["Category", "Planned", "Actual", "Remaining", "% Used", "Status"]],
    body: bva.rows.map(r => [
      r.category,
      money(r.planned),
      money(r.actual),
      money(r.remaining),
      `${r.pctUsed.toFixed(0)}%`,
      r.status
    ]),
    styles: { fontSize: 9 },
    headStyles: { fillColor: [20, 22, 29] }
  });

  y = doc.lastAutoTable.finalY + 16;

  doc.setFontSize(12);
  doc.text("Debts", 40, y);
  y += 8;

  doc.autoTable({
    startY: y,
    head: [["Name", "Balance", "APR", "Min Payment"]],
    body: state.debts.map(d => [d.name, money(d.balance), `${num(d.aprPct).toFixed(2)}%`, money(d.minPayment)]),
    styles: { fontSize: 9 },
    headStyles: { fillColor: [20, 22, 29] }
  });

  y = doc.lastAutoTable.finalY + 16;

  doc.setFontSize(12);
  doc.text("Debt Payoff Schedule (first 60 months)", 40, y);
  y += 8;

  doc.autoTable({
    startY: y,
    head: [["Month", "Target", "Paid", "Interest", "Principal", "Total Balance Remaining"]],
    body: plan.schedule.slice(0, 60).map(r => [
      String(r.month),
      r.target,
      money(r.paid),
      money(r.interest),
      money(r.principal),
      money(r.totalBalanceRemaining)
    ]),
    styles: { fontSize: 8 },
    headStyles: { fillColor: [20, 22, 29] }
  });

  doc.save(`checks-and-balances_${state.ledgerMonth}.pdf`);
}

function exportExcel() {
  const wb = XLSX.utils.book_new();

  const inc = calcIncomeMonthly(state.income);
  const plannedTotal = sumPlannedExpenses(state.expenses);
  const monthTxns = spendingForMonth(state.spending, state.ledgerMonth);
  const actualTotal = sumSpending(monthTxns);
  const variance = plannedTotal - actualTotal;
  const cashLeft = inc.monthlyNet - plannedTotal;

  const debtTotal = state.debts.reduce((a, d) => a + num(d.balance), 0);
  const plan = buildPayoffPlan(state.debts, state.payoff.strategy, state.payoff.extraPayment);
  const bva = budgetVsActual(state.expenses, monthTxns);

  // Summary
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
    ["Metric", "Value"],
    ["Ledger Month", state.ledgerMonth],
    ["Monthly Gross (est.)", inc.monthlyGross],
    ["Monthly Take-Home (est.)", inc.monthlyNet],
    ["Planned Expenses", plannedTotal],
    ["Actual Spending (ledger)", actualTotal],
    ["Variance (Planned − Actual)", variance],
    ["Cash Left (Net − Planned)", cashLeft],
    ["Total Debt", debtTotal],
    ["Debt Strategy", state.payoff.strategy],
    ["Extra Monthly Debt Payment", num(state.payoff.extraPayment)],
    ["Estimated Payoff Time (months)", plan.months === null ? "" : plan.months]
  ]), "Summary");

  // Income details
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
    ["Field", "Value"],
    ["Pay frequency", state.income.frequency],
    ["Gross per paycheck", num(state.income.grossPerPaycheck)],
    ["Tax rate (%)", num(state.income.taxRatePct)],
    ["Other deductions per paycheck", num(state.income.otherDeductionsPerPaycheck)],
    ["Other monthly income", num(state.income.otherMonthlyIncome)],
    ["Paychecks per month factor", inc.ppm],
    ["Monthly gross", inc.monthlyGross],
    ["Monthly taxes", inc.taxes],
    ["Monthly deductions", inc.deductions],
    ["Monthly take-home (net)", inc.monthlyNet]
  ]), "Income");

  // Planned expenses
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
    ["Name", "Category", "Amount"],
    ...state.expenses.map(e => [e.name, e.category, num(e.amount)])
  ]), "PlannedExpenses");

  // Spending (all)
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
    ["Date", "Month", "Category", "Description", "Amount"],
    ...state.spending.map(t => [t.date, monthFromDateISO(t.date), t.category, t.description, num(t.amount)])
  ]), "SpendingAll");

  // Spending (month)
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
    ["Date", "Category", "Description", "Amount"],
    ...monthTxns.map(t => [t.date, t.category, t.description, num(t.amount)])
  ]), `Spending_${state.ledgerMonth}`);

  // Budget vs actual
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
    ["Category", "Planned", "Actual", "Remaining", "% Used", "Status"],
    ...bva.rows.map(r => [r.category, r.planned, r.actual, r.remaining, r.pctUsed, r.status])
  ]), "BudgetVsActual");

  // Debts
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
    ["Name", "Balance", "APR (%)", "Min Payment"],
    ...state.debts.map(d => [d.name, num(d.balance), num(d.aprPct), num(d.minPayment)])
  ]), "Debts");

  // Payoff
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
    ["Month", "Target", "Paid", "Interest", "Principal", "Total Balance Remaining"],
    ...plan.schedule.map(r => [r.month, r.target, r.paid, r.interest, r.principal, r.totalBalanceRemaining])
  ]), "PayoffPlan");

  XLSX.writeFile(wb, `checks-and-balances_${state.ledgerMonth}.xlsx`);
}

/* ------------------------------ Event Wiring ------------------------------ */

function wire() {
  // category selects
  renderCategorySelects();

  // Inputs
  $("incomeFrequency").addEventListener("change", e => { state.income.frequency = e.target.value; scheduleRender(); });
  $("grossPerPaycheck").addEventListener("input", e => { state.income.grossPerPaycheck = num(e.target.value); scheduleRender(); });
  $("taxRatePct").addEventListener("input", e => { state.income.taxRatePct = clamp(num(e.target.value), 0, 60); scheduleRender(); });
  $("otherDeductionsPerPaycheck").addEventListener("input", e => { state.income.otherDeductionsPerPaycheck = num(e.target.value); scheduleRender(); });
  $("otherMonthlyIncome").addEventListener("input", e => { state.income.otherMonthlyIncome = num(e.target.value); scheduleRender(); });

  $("ledgerMonth").addEventListener("change", e => { state.ledgerMonth = e.target.value || currentMonthISO(); scheduleRender(); });

  $("payoffStrategy").addEventListener("change", e => { state.payoff.strategy = e.target.value; scheduleRender(); });
  $("extraPayment").addEventListener("input", e => { state.payoff.extraPayment = Math.max(0, num(e.target.value)); scheduleRender(); });

  // Adds
  $("btnAddExpense").addEventListener("click", addExpense);
  $("btnAddSpend").addEventListener("click", addSpending);
  $("btnAddDebt").addEventListener("click", addDebt);

  // Enter-to-add UX
  $("expenseAmount").addEventListener("keydown", (e) => { if (e.key === "Enter") addExpense(); });
  $("spendAmount").addEventListener("keydown", (e) => { if (e.key === "Enter") addSpending(); });
  $("debtMin").addEventListener("keydown", (e) => { if (e.key === "Enter") addDebt(); });

  // Event delegation for deletes
  $("expenseTbody").addEventListener("click", (e) => {
    const b = e.target.closest("button[data-action]");
    if (!b) return;
    if (b.getAttribute("data-action") === "delete-expense") {
      deleteById(state.expenses, b.getAttribute("data-id"));
      scheduleRender();
    }
  });

  $("spendTbody").addEventListener("click", (e) => {
    const b = e.target.closest("button[data-action]");
    if (!b) return;
    if (b.getAttribute("data-action") === "delete-spending") {
      deleteById(state.spending, b.getAttribute("data-id"));
      scheduleRender();
    }
  });

  $("debtTbody").addEventListener("click", (e) => {
    const b = e.target.closest("button[data-action]");
    if (!b) return;
    if (b.getAttribute("data-action") === "delete-debt") {
      deleteById(state.debts, b.getAttribute("data-id"));
      scheduleRender();
    }
  });

  // Exports
  $("btnExportPdf").addEventListener("click", exportPdf);
  $("btnExportExcel").addEventListener("click", exportExcel);

  // Reset
  $("btnReset").addEventListener("click", () => {
    const ok = confirm("Reset all data? This clears saved entries in this browser.");
    if (!ok) return;
    localStorage.removeItem(LS_KEY);
    state = defaultState();
    renderInputs();
    scheduleRender();
  });
}

/* ------------------------------ Init ------------------------------ */

(function init() {
  load();
  renderCategorySelects();
  renderInputs();

  // Ensure ledgerMonth is valid
  if (!state.ledgerMonth || state.ledgerMonth.length !== 7) state.ledgerMonth = currentMonthISO();
  $("ledgerMonth").value = state.ledgerMonth;

  wire();
  scheduleRender();
})();
