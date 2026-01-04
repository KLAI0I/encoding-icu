/*******************************************************
 * ICU ENCODING DASHBOARD - Code.gs (FINAL)
 *******************************************************/

const CONFIG = {
  SHEET_ID: "1CefMMqIKHicwo_ZFyBT6uLw_XWz-1nKkYgjbJq4InUM",

  // OPTIONAL: Force a specific tab name. Leave "" to auto-detect.
  SUPPLIES_TAB: "",

  // OPTIONAL: Force a services tab name. Leave "" to auto-detect (if possible).
  SERVICES_TAB: "",

  USERNAME: "icu",
  PASSWORD: "ghonim",

  // Optional default recipient for Zero Stock email (leave "" to type in UI)
  DEFAULT_ZERO_EMAIL: "",

  // Qty state saved after XLS upload and after copy(-1)
  PROP_QTY_STATE: "ICU_QTY_STATE_V2",
  PROP_LAST_UPLOAD_META: "ICU_LAST_UPLOAD_META_V2",

  // Manual override map for unmatched XLS rows: { rowKey: "Category Name" }
  PROP_UNMATCHED_OVERRIDES: "ICU_UNMATCHED_OVERRIDES_V1",

  // Manual override map for ANY item category:
  // { "SUPPLY|CODE:xxxxx": "New Category", "SERVICE|CODE:yyyy": "New Category" }
  PROP_ITEM_CATEGORY_OVERRIDES: "ICU_ITEM_CATEGORY_OVERRIDES_V1",
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile("index")
    .setTitle("ICU Encoding Dashboard")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function login(username, password) {
  const ok =
    String(username || "") === CONFIG.USERNAME &&
    String(password || "") === CONFIG.PASSWORD;
  return { ok };
}

function getDashboardData() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);

  const suppliesTabName =
    CONFIG.SUPPLIES_TAB && CONFIG.SUPPLIES_TAB.trim()
      ? CONFIG.SUPPLIES_TAB.trim()
      : detectSuppliesTab_(ss);

  const supplies = parseSuppliesTab_(ss, suppliesTabName);

  let servicesTabName = "";
  let services = [];
  try {
    servicesTabName =
      CONFIG.SERVICES_TAB && CONFIG.SERVICES_TAB.trim()
        ? CONFIG.SERVICES_TAB.trim()
        : detectServicesTab_(ss);

    if (servicesTabName) services = parseServicesTab_(ss, servicesTabName);
  } catch (e) {
    servicesTabName = "";
    services = [];
  }

  // Apply item category overrides
  const catOverrides = getItemCategoryOverrides_();
  applyCategoryOverrides_(supplies, "SUPPLY", catOverrides);
  applyCategoryOverrides_(services, "SERVICE", catOverrides);

  return {
    supplies,
    services,
    categoryOverrides: catOverrides,
    meta: {
      sheetId: CONFIG.SHEET_ID,
      suppliesTab: suppliesTabName,
      servicesTab: servicesTabName || "—",
      suppliesCount: supplies.length,
      servicesCount: services.length,
    },
  };
}

// ---------- tab detection ----------
function detectSuppliesTab_(ss) {
  const sheets = ss.getSheets();
  let best = { name: null, score: -1 };

  for (const sh of sheets) {
    const name = sh.getName();
    const rows = Math.min(250, sh.getLastRow());
    const cols = Math.min(6, sh.getLastColumn());
    if (rows < 2 || cols < 2) continue;

    const values = sh.getRange(1, 1, rows, cols).getDisplayValues();

    let foundItemCode = false;
    let foundItemName = false;
    let numericCodes = 0;

    for (let r = 0; r < values.length; r++) {
      const a = (values[r][0] || "").trim().toUpperCase();
      const b = (values[r][1] || "").trim().toUpperCase();

      if (a === "ITEM CODE") foundItemCode = true;
      if (b === "ITEM NAME" || b.includes("ITEM NAME")) foundItemName = true;

      const aRaw = (values[r][0] || "").trim();
      if (looksLikeItemCode_(aRaw)) numericCodes++;
    }

    let score = 0;
    if (foundItemCode) score += 5;
    if (foundItemName) score += 5;
    score += Math.min(25, numericCodes);
    if (sh.getLastRow() < 20) score -= 2;

    if (score > best.score) best = { name, score };
  }

  if (!best.name) {
    throw new Error(
      "Could not auto-detect supplies tab. Set CONFIG.SUPPLIES_TAB to the exact tab name."
    );
  }
  return best.name;
}

function detectServicesTab_(ss) {
  const sheets = ss.getSheets();
  for (const sh of sheets) {
    if (sh.getLastRow() < 2) continue;

    const header = sh
      .getRange(1, 1, 1, Math.min(20, sh.getLastColumn()))
      .getDisplayValues()[0]
      .map((x) => String(x || "").trim().toLowerCase());

    const hasCode = header.some((h) => h === "code" || h.includes("service code"));
    const hasDesc = header.some((h) => h.includes("description") || h.includes("name"));

    if (hasCode && hasDesc) return sh.getName();
  }
  return "";
}

// ---------- supplies parsing ----------
function parseSuppliesTab_(ss, tabName) {
  const sh = ss.getSheetByName(tabName);
  if (!sh) throw new Error("Supplies tab not found: " + tabName);

  const values = sh.getDataRange().getDisplayValues();
  if (!values || values.length < 2) return [];

  let currentCategory = "Uncategorized";
  let currentSubcategory = "";
  const items = [];

  for (let r = 0; r < values.length; r++) {
    const row = values[r];
    const a = (row[0] || "").trim();
    const b = (row[1] || "").trim();
    if (!a && !b) continue;

    const aUpper = a.toUpperCase();
    const bUpper = b.toUpperCase();

    if (aUpper === "ITEM CODE" && (bUpper === "ITEM NAME" || bUpper.includes("ITEM"))) continue;

    const looksLikeCategory =
      a &&
      !b &&
      !looksLikeItemCode_(a) &&
      aUpper === a &&
      a.length <= 80;

    if (looksLikeCategory) {
      currentCategory = a;
      currentSubcategory = "";
      continue;
    }

    if (looksLikeItemCode_(a) && b) {
      items.push({
        type: "SUPPLY",
        category: currentCategory,
        subcategory: currentSubcategory,
        itemCode: a,
        itemName: b,
      });
    }
  }

  return items;
}

function looksLikeItemCode_(s) {
  const x = String(s || "").trim();
  if (!x) return false;
  const digits = x.replace(/\D/g, "");
  return digits.length >= 6;
}

// ---------- services parsing ----------
function parseServicesTab_(ss, tabName) {
  const sh = ss.getSheetByName(tabName);
  if (!sh) return [];

  const values = sh.getDataRange().getValues();
  if (!values || values.length < 2) return [];

  const header = values[0].map((h) => String(h || "").trim().toLowerCase());
  const rows = values.slice(1);

  const col = {
    category: header.findIndex((x) => x.includes("category")),
    code: header.findIndex((x) => x === "code" || x.includes("service code")),
    desc: header.findIndex((x) => x.includes("description") || x.includes("name")),
  };

  if (col.code < 0 || col.desc < 0) return [];

  const out = [];
  for (const r of rows) {
    const code = String(r[col.code] || "").trim();
    const desc = String(r[col.desc] || "").trim();
    if (!code && !desc) continue;

    out.push({
      type: "SERVICE",
      category: col.category >= 0 ? String(r[col.category] || "").trim() : "Services",
      subcategory: "",
      itemCode: code,
      itemName: desc,
    });
  }
  return out;
}

// ---------- qty state ----------
function getQtyState() {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty(CONFIG.PROP_QTY_STATE);
  const metaRaw = props.getProperty(CONFIG.PROP_LAST_UPLOAD_META);
  const ovRaw = props.getProperty(CONFIG.PROP_UNMATCHED_OVERRIDES);
  const catRaw = props.getProperty(CONFIG.PROP_ITEM_CATEGORY_OVERRIDES);

  return {
    qtyState: raw ? JSON.parse(raw) : {},
    meta: metaRaw ? JSON.parse(metaRaw) : null,
    overrides: ovRaw ? JSON.parse(ovRaw) : {},
    categoryOverrides: catRaw ? JSON.parse(catRaw) : {},
  };
}

function saveQtyState(qtyState, meta) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty(CONFIG.PROP_QTY_STATE, JSON.stringify(qtyState || {}));
  if (meta) props.setProperty(CONFIG.PROP_LAST_UPLOAD_META, JSON.stringify(meta));
  return { ok: true };
}

// ---------- unmatched XLS override ----------
function setUnmatchedOverride(rowKey, category) {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty(CONFIG.PROP_UNMATCHED_OVERRIDES);
  const overrides = raw ? JSON.parse(raw) : {};

  const k = String(rowKey || "").trim();
  const c = String(category || "").trim();

  if (!k) return { ok: false, message: "Missing rowKey" };

  if (!c || c === "__AUTO__") {
    delete overrides[k];
  } else {
    overrides[k] = c;
  }

  props.setProperty(CONFIG.PROP_UNMATCHED_OVERRIDES, JSON.stringify(overrides));
  return { ok: true, overrides };
}

// ---------- item category overrides (FIXED normalization) ----------
function setItemCategoryOverride(itemType, itemKey, category) {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty(CONFIG.PROP_ITEM_CATEGORY_OVERRIDES);
  const map = raw ? JSON.parse(raw) : {};

  const t = String(itemType || "").trim().toUpperCase(); // SUPPLY|SERVICE
  let k = String(itemKey || "").trim();
  const c = String(category || "").trim();

  if (!t || !k) return { ok: false, message: "Missing itemType or itemKey" };

  // ✅ normalize key to match apply logic
  if (k.toUpperCase().startsWith("CODE:")) {
    k = "CODE:" + k.slice(5).trim().toLowerCase();
  } else if (k.toUpperCase().startsWith("NAME:")) {
    k = "NAME:" + k.slice(5).trim().toLowerCase();
  } else {
    k = k.trim().toLowerCase();
  }

  const storageKey = `${t}|${k}`;

  if (!c) {
    delete map[storageKey];
  } else {
    map[storageKey] = c;
  }

  props.setProperty(CONFIG.PROP_ITEM_CATEGORY_OVERRIDES, JSON.stringify(map));
  return { ok: true, overrides: map };
}

function getItemCategoryOverrides_() {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty(CONFIG.PROP_ITEM_CATEGORY_OVERRIDES);
  return raw ? JSON.parse(raw) : {};
}

function applyCategoryOverrides_(items, type, overrides) {
  for (const it of items) {
    const code = String(it.itemCode || "").trim().toLowerCase();
    const name = String(it.itemName || "").trim().toLowerCase();
    const k = code ? "CODE:" + code : "NAME:" + name;

    // ✅ exact normalized
    let ov = overrides[`${type}|${k}`];

    // ✅ backward compatibility for any old mixed-case keys
    if (!ov) {
      ov = overrides[`${type}|${k.toUpperCase()}`] || overrides[`${type}|${k.toLowerCase()}`];
    }

    if (ov) it.category = ov;
  }
}
// ---------- zero stock email (Excel) ----------
function sendZeroStockExcelEmail(email) {
  const to = String(email || "").trim() || String(CONFIG.DEFAULT_ZERO_EMAIL || "").trim();
  if (!to) return { ok: false, message: "Missing recipient email" };

  // Read supplies with overrides applied + current qty state
  const data = getDashboardData();
  const qtyObj = getQtyState();
  const qtyState = qtyObj && qtyObj.qtyState ? qtyObj.qtyState : {};

  const zeroRows = [];
  for (const it of (data.supplies || [])) {
    const k = itemKeyForState_(it);
    const q = Number(qtyState[k] ?? 0);
    if (q === 0) {
      zeroRows.push([it.category || "Uncategorized", it.itemCode || "", it.itemName || "", 0]);
    }
  }

  if (!zeroRows.length) {
    return { ok: false, message: "No zero-stock items to send" };
  }

  // Create a temporary spreadsheet and export as XLSX
  const tmp = SpreadsheetApp.create("ICU_Zero_Stock_Temp");
  const sh = tmp.getSheets()[0];
  sh.setName("ZeroStock");
  sh.getRange(1, 1, 1, 4).setValues([["Category", "Item Code", "Item Name", "Qty"]]);
  sh.getRange(2, 1, zeroRows.length, 4).setValues(zeroRows);
  sh.autoResizeColumns(1, 4);

  const tmpId = tmp.getId();
  const fileName = "ICU_Zero_Stock_Supplies_" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmm") + ".xlsx";

  const url = "https://docs.google.com/spreadsheets/d/" + tmpId + "/export?format=xlsx";
  const token = ScriptApp.getOAuthToken();
  const resp = UrlFetchApp.fetch(url, { headers: { Authorization: "Bearer " + token } });
  const blob = resp.getBlob().setName(fileName);

  // Email
  const subject = "ICU Zero Stock Supplies (Excel)";
  const body =
    "Hello,\n\n" +
    "Please find attached the ICU Zero Stock Supplies list (Qty = 0).\n\n" +
    "Generated: " + new Date().toLocaleString() + "\n\n" +
    "Designed by Mr.Mohamed Ali Ghonim – ICU Head Nurse";

  GmailApp.sendEmail(to, subject, body, { attachments: [blob] });

  // Cleanup temp file
  DriveApp.getFileById(tmpId).setTrashed(true);

  return { ok: true, sentTo: to, rows: zeroRows.length };
}

function itemKeyForState_(it) {
  const code = String(it.itemCode || "").trim().toLowerCase();
  if (code) return "CODE:" + code;
  const name = String(it.itemName || "").trim().toLowerCase();
  return "NAME:" + name;
}
