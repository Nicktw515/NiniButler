/**
 * 尼尼管家公｜Google Apps Script 後端
 * ------------------------------------------------
 * 使用方式：
 * 1. 打開你要用來記帳的 Google Sheet
 * 2. 「擴充功能」→「Apps Script」
 * 3. 把這個檔案的內容整份貼進去（取代預設的 Code.gs）
 * 4. 右上角「部署」→「新增部署作業」
 *      - 類型選「網頁應用程式」
 *      - 執行身分：我
 *      - 誰可以存取：知道連結的任何人
 * 5. 複製產生的網址，貼到「尼尼管家公」App 的設定頁「記帳 API 網址」
 *
 * 這個腳本會自動在你的 Sheet 裡建立兩個工作表：Members、Expenses
 *
 * 設計說明：
 * Expenses 工作表前半部是給「人」看的欄位（日期、付款人姓名、分帳對象姓名…），
 * 後半部是給 App 自己讀寫用的系統欄位（id、時間戳記等），不會影響你直接在
 * Google Sheet 裡讀懂帳目內容。系統欄位建議不要手動修改，以免 App 讀取錯誤。
 */

const MEMBERS_SHEET = "Members";
const EXPENSES_SHEET = "Expenses";
const SETTLEMENTS_SHEET = "Settlements";

// 內部使用的欄位鍵值（順序需與下方 LABELS 一一對應）
const MEMBERS_HEADERS = ["id", "name", "temporary"];
const MEMBERS_LABELS  = ["旅伴編號(系統用)", "姓名", "臨時旅伴"];

const EXPENSES_HEADERS = [
  // ---- 給人看的欄位 ----
  "date_display", "payer_name", "amount", "currency", "category_display", "note",
  "participants_display", "split_type_display", "split_display", "paid_display", "created_display",
  // ---- App 系統欄位（請勿手動編輯）----
  "id", "payer", "participants", "split_type", "split_data", "date", "created_at", "category", "paid"
];
const EXPENSES_LABELS = [
  "日期", "付款人", "金額", "幣別", "類別", "備註",
  "分帳對象", "分帳方式", "分帳明細", "還款狀態", "建立時間",
  "編號(系統用)", "付款人ID(系統用)", "分帳對象ID(系統用)", "分帳方式代碼(系統用)",
  "分帳資料(系統用)", "日期時間戳記(系統用)", "建立時間戳記(系統用)", "類別代碼(系統用)", "還款狀態(系統用)"
];

const SETTLEMENTS_HEADERS = [
  // ---- 給人看的欄位 ----
  "date_display", "from_name", "to_name", "amount", "currency", "note",
  // ---- App 系統欄位（請勿手動編輯）----
  "id", "from_id", "to_id", "created_at"
];
const SETTLEMENTS_LABELS = [
  "還款日期", "還款人", "收款人", "金額", "幣別", "備註",
  "編號(系統用)", "還款人ID(系統用)", "收款人ID(系統用)", "建立時間戳記(系統用)"
];

const READ_ACTIONS = ["getMembers", "getExpenses", "getSettlements", "getAll"];

function doPost(e) {
  let lock = null;
  let lockAcquired = false;
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const needsLock = READ_ACTIONS.indexOf(action) === -1;

    if (needsLock) {
      lock = LockService.getScriptLock();
      try {
        lock.waitLock(10000);
        lockAcquired = true;
      } catch (lockErr) {
        return jsonOut({ error: "尼尼現在有點忙（同時有其他人在寫入），請稍等一下再試一次" });
      }
    }

    let result;
    switch (action) {
      case "getAll":
        result = { members: getMembers(ss), expenses: getExpenses(ss), settlements: getSettlements(ss) };
        break;
      case "getMembers":
        result = { members: getMembers(ss) };
        break;
      case "addMember":
        addOrUpdateMember(ss, body.member);
        result = { ok: true };
        break;
      case "deleteMember":
        deleteMember(ss, body.id);
        result = { ok: true };
        break;
      case "getExpenses":
        result = { expenses: getExpenses(ss) };
        break;
      case "addExpense":
        addExpense(ss, body.expense);
        result = { ok: true };
        break;
      case "updateExpense":
        updateExpense(ss, body.expense);
        result = { ok: true };
        break;
      case "deleteExpense":
        deleteExpense(ss, body.id);
        result = { ok: true };
        break;
      case "getSettlements":
        result = { settlements: getSettlements(ss) };
        break;
      case "addSettlement":
        addSettlement(ss, body.settlement);
        result = { ok: true };
        break;
      case "deleteSettlement":
        deleteSettlement(ss, body.id);
        result = { ok: true };
        break;
      default:
        result = { error: "unknown action: " + action };
    }
    return jsonOut(result);
  } catch (err) {
    return jsonOut({ error: String(err) });
  } finally {
    if (lockAcquired) lock.releaseLock();
  }
}

function doGet(e) {
  return jsonOut({ ok: true, message: "尼尼管家公 API 運作中" });
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ---------------- Sheet helpers ---------------- */

function getOrCreateSheet(ss, name, headers, labels) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(labels);
    sheet.setFrozenRows(1);
  } else if (sheet.getLastRow() === 0) {
    sheet.appendRow(labels);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function sheetToObjects(sheet, headers) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  return values
    .filter(row => row[0] !== "")
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = row[i]; });
      return obj;
    });
}

function findRowIndexByKey(sheet, headers, key, value) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const colIdx = headers.indexOf(key) + 1; // 1-indexed column
  if (colIdx < 1) return -1;
  const ids = sheet.getRange(2, colIdx, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0] === value) return i + 2; // 1-indexed row, +1 for header
  }
  return -1;
}

function formatDateDisplay(ms) {
  if (!ms) return "";
  try {
    return Utilities.formatDate(new Date(Number(ms)), Session.getScriptTimeZone() || "Asia/Taipei", "yyyy/MM/dd HH:mm");
  } catch (e) {
    return String(ms);
  }
}

function splitTypeLabel(type) {
  if (type === "ratio") return "按比例分攤";
  if (type === "custom") return "自訂金額";
  return "平均分攤";
}

const CATEGORY_LABELS = {
  food: "🍜 餐飲",
  transport: "🚕 交通",
  stay: "🏨 住宿",
  ticket: "🎫 門票",
  shopping: "🛍️ 購物",
  other: "🧾 其他"
};
function categoryLabel(key) {
  return CATEGORY_LABELS[key] || CATEGORY_LABELS.other;
}

/* ---------------- Members ---------------- */

function getMembers(ss) {
  const sheet = getOrCreateSheet(ss, MEMBERS_SHEET, MEMBERS_HEADERS, MEMBERS_LABELS);
  const rows = sheetToObjects(sheet, MEMBERS_HEADERS);
  return rows.map(r => ({
    id: r.id,
    name: r.name,
    temporary: (r.temporary === "是" || r.temporary === true)
  }));
}

function getMemberNameMap(ss) {
  const map = {};
  getMembers(ss).forEach(m => { map[m.id] = m.name; });
  return map;
}

function addOrUpdateMember(ss, member) {
  const sheet = getOrCreateSheet(ss, MEMBERS_SHEET, MEMBERS_HEADERS, MEMBERS_LABELS);
  const rowIdx = findRowIndexByKey(sheet, MEMBERS_HEADERS, "id", member.id);
  const row = [member.id, member.name, member.temporary ? "是" : "否"];
  if (rowIdx > -1) {
    sheet.getRange(rowIdx, 1, 1, row.length).setValues([row]);
  } else {
    sheet.appendRow(row);
  }
}

function deleteMember(ss, id) {
  const sheet = getOrCreateSheet(ss, MEMBERS_SHEET, MEMBERS_HEADERS, MEMBERS_LABELS);
  const rowIdx = findRowIndexByKey(sheet, MEMBERS_HEADERS, "id", id);
  if (rowIdx > -1) sheet.deleteRow(rowIdx);
}

/* ---------------- Expenses ---------------- */

function getExpenses(ss) {
  const sheet = getOrCreateSheet(ss, EXPENSES_SHEET, EXPENSES_HEADERS, EXPENSES_LABELS);
  const rows = sheetToObjects(sheet, EXPENSES_HEADERS);
  return rows.map(r => ({
    id: r.id,
    date: r.date,
    payer: r.payer,
    amount: Number(r.amount),
    currency: r.currency,
    note: r.note,
    participants: safeParseArray(r.participants),
    split_type: r.split_type,
    split_data: safeParseObject(r.split_data),
    created_at: r.created_at,
    category: r.category || "other",
    paid: safeParseObject(r.paid)
  }));
}

function buildPaidDisplay(expense, nameMap) {
  const paid = expense.paid || {};
  const parts = (expense.participants || []).filter(id => id !== expense.payer);
  if (parts.length === 0) return "";
  return parts.map(id => (nameMap[id] || id) + (paid[id] ? " ✅已還" : " 未還")).join("、");
}

function buildSplitDisplay(expense, nameMap) {
  const type = expense.split_type || "equal";
  const parts = expense.participants || [];
  const cur = expense.currency || "";
  if (type === "equal") {
    const each = parts.length ? (Number(expense.amount) || 0) / parts.length : 0;
    return parts.map(id => (nameMap[id] || id)).join("、") + " 各付 " + round2(each) + " " + cur;
  }
  if (type === "ratio") {
    return parts.map(id => (nameMap[id] || id) + " " + (Number((expense.split_data || {})[id]) || 0) + "%").join("、");
  }
  // custom
  return parts.map(id => (nameMap[id] || id) + " " + round2(Number((expense.split_data || {})[id]) || 0) + " " + cur).join("、");
}

function round2(n) {
  return Math.round((Number(n) || 0) * 100) / 100;
}

function addExpense(ss, expense) {
  const sheet = getOrCreateSheet(ss, EXPENSES_SHEET, EXPENSES_HEADERS, EXPENSES_LABELS);
  sheet.appendRow(expenseToRow(ss, expense));
}

function updateExpense(ss, expense) {
  const sheet = getOrCreateSheet(ss, EXPENSES_SHEET, EXPENSES_HEADERS, EXPENSES_LABELS);
  const rowIdx = findRowIndexByKey(sheet, EXPENSES_HEADERS, "id", expense.id);
  const row = expenseToRow(ss, expense);
  if (rowIdx > -1) {
    sheet.getRange(rowIdx, 1, 1, row.length).setValues([row]);
  } else {
    sheet.appendRow(row);
  }
}

function deleteExpense(ss, id) {
  const sheet = getOrCreateSheet(ss, EXPENSES_SHEET, EXPENSES_HEADERS, EXPENSES_LABELS);
  const rowIdx = findRowIndexByKey(sheet, EXPENSES_HEADERS, "id", id);
  if (rowIdx > -1) sheet.deleteRow(rowIdx);
}

function expenseToRow(ss, e) {
  const nameMap = getMemberNameMap(ss);
  const participants = e.participants || [];
  return [
    // ---- 給人看 ----
    formatDateDisplay(e.date),
    nameMap[e.payer] || e.payer,
    e.amount,
    e.currency,
    categoryLabel(e.category),
    e.note || "",
    participants.map(id => nameMap[id] || id).join("、"),
    splitTypeLabel(e.split_type),
    buildSplitDisplay(e, nameMap),
    buildPaidDisplay(e, nameMap),
    formatDateDisplay(e.created_at),
    // ---- 系統欄位 ----
    e.id,
    e.payer,
    JSON.stringify(participants),
    e.split_type || "equal",
    JSON.stringify(e.split_data || {}),
    e.date,
    e.created_at,
    e.category || "other",
    JSON.stringify(e.paid || {})
  ];
}

function safeParseArray(str) {
  try { const v = JSON.parse(str); return Array.isArray(v) ? v : []; }
  catch (e) { return []; }
}
function safeParseObject(str) {
  try { const v = JSON.parse(str); return (v && typeof v === "object") ? v : {}; }
  catch (e) { return {}; }
}

/* ---------------- Settlements (已還款紀錄) ---------------- */

function getSettlements(ss) {
  const sheet = getOrCreateSheet(ss, SETTLEMENTS_SHEET, SETTLEMENTS_HEADERS, SETTLEMENTS_LABELS);
  const rows = sheetToObjects(sheet, SETTLEMENTS_HEADERS);
  return rows.map(r => ({
    id: r.id,
    from: r.from_id,
    to: r.to_id,
    amount: Number(r.amount),
    currency: r.currency,
    note: r.note,
    created_at: r.created_at
  }));
}

function addSettlement(ss, s) {
  const sheet = getOrCreateSheet(ss, SETTLEMENTS_SHEET, SETTLEMENTS_HEADERS, SETTLEMENTS_LABELS);
  const nameMap = getMemberNameMap(ss);
  sheet.appendRow([
    formatDateDisplay(s.created_at),
    nameMap[s.from] || s.from,
    nameMap[s.to] || s.to,
    s.amount,
    s.currency,
    s.note || "",
    s.id,
    s.from,
    s.to,
    s.created_at
  ]);
}

function deleteSettlement(ss, id) {
  const sheet = getOrCreateSheet(ss, SETTLEMENTS_SHEET, SETTLEMENTS_HEADERS, SETTLEMENTS_LABELS);
  const rowIdx = findRowIndexByKey(sheet, SETTLEMENTS_HEADERS, "id", id);
  if (rowIdx > -1) sheet.deleteRow(rowIdx);
}
