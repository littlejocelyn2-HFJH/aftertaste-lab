/**
 * 回味製研所：情感特調 — 後端（Google Apps Script）
 * 部署：部署 → 新增部署作業 → 類型「網頁應用程式」
 *   執行身分：我　／　誰可以存取：所有人
 * ★ 請把第 8 行的密碼改成你自己的（維持 change-me 時，一律視為密碼錯誤）
 */

const TEACHER_PASSWORD = "change-me";

const ACT = "diner";
const DATA_SHEET = "data";
const LOG_SHEET = "作品紀錄";
const DATA_HEADERS = ["key", "班級", "座號", "姓名", "活動", "最多字數", "作品數", "最後完成", "最新料理", "最新情感", "最新圖片", "attemptsJSON"];
const LOG_HEADERS = ["完成時間", "班級", "座號", "姓名", "料理", "共餐者", "步驟二 記憶濾鏡", "步驟三 人物與刻痕", "情感比例", "一句話總結", "總字數", "用時(秒)", "提醒"];
const FLAVOR_NAMES = { sour: "微酸", sweet: "溫甜", bitter: "苦澀", spicy: "嗆辣", salty: "鹹香" };

function doGet() {
  return json_({ ok: true, service: "aftertaste-lab", time: new Date().toISOString() });
}

function doPost(e) {
  let req = {};
  try { req = JSON.parse((e && e.postData && e.postData.contents) || "{}"); }
  catch (err) { return json_({ ok: false, error: "bad json" }); }
  try {
    if (req.action === "check") return json_({ ok: true, teacher: auth_(req.password) });
    if (req.action === "list") return json_(list_(req));
    if (req.action === "save") return json_(save_(req));
    if (req.action === "clearClass") return json_(clear_(req));
    return json_({ ok: false, error: "unknown action" });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function auth_(p) {
  return TEACHER_PASSWORD !== "change-me" && typeof p === "string" && p === TEACHER_PASSWORD;
}

function sheet_(name, headers) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1000, headers.length).setNumberFormat("@"); // 全部當文字，避免座號、時間被自動轉格式
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
    sh.setFrozenRows(1);
  }
  return sh;
}

function safe_(s, d) {
  try { return s ? JSON.parse(s) : d; } catch (e) { return d; }
}

function txt_(s, max) {
  s = String(s == null ? "" : s).slice(0, max || 5000);
  return /^[=+\-@]/.test(s) ? "'" + s : s; // 避免被當成公式
}

function list_(req) {
  const teacher = auth_(req.password);
  const v = sheet_(DATA_SHEET, DATA_HEADERS).getDataRange().getValues();
  const out = [];
  for (let i = 1; i < v.length; i++) {
    const r = v[i];
    if (!r[0]) continue;
    const item = {
      cls: String(r[1]), seat: String(r[2]), name: String(r[3]),
      best: Number(r[5]) || 0, count: Number(r[6]) || 0,
      food: String(r[8] || ""), flavors: safe_(r[9], {})
    };
    if (teacher) {
      item.last = String(r[7] || "");
      item.image = String(r[10] || "");
      item.attempts = safe_(r[11], []);
    }
    out.push(item);
  }
  return { ok: true, teacher: teacher, records: out };
}

function save_(req) {
  const rec = req.record || {};
  const cls = String(rec.cls || "").trim().slice(0, 10);
  const seat = String(rec.seat || "").trim().slice(0, 5);
  const name = String(rec.name || "").trim().slice(0, 20);
  if (!cls || !seat) return { ok: false, error: "缺少班級或座號" };
  const incoming = (Array.isArray(rec.attempts) ? rec.attempts : []).filter(function (a) { return a && a.at; });

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_(DATA_SHEET, DATA_HEADERS);
    const v = sh.getDataRange().getValues();
    const key = cls + "|" + seat + "|" + ACT;
    let idx = -1;
    for (let i = 1; i < v.length; i++) if (String(v[i][0]) === key) { idx = i; break; }

    const old = idx > 0 ? safe_(v[idx][11], []) : [];
    const seen = {};
    old.forEach(function (a) { seen[a.at] = 1; });
    const fresh = incoming.filter(function (a) { return !seen[a.at]; });

    let all = old.concat(fresh).sort(function (a, b) { return String(a.at).localeCompare(String(b.at)); }).slice(-20);
    let js = JSON.stringify(all);
    while (js.length > 45000 && all.length > 1) { all.shift(); js = JSON.stringify(all); }

    const latest = all[all.length - 1] || {};
    const best = all.reduce(function (m, a) { return Math.max(m, Number(a.chars) || 0); }, idx > 0 ? Number(v[idx][5]) || 0 : 0);
    const count = (idx > 0 ? Number(v[idx][6]) || 0 : 0) + fresh.length;
    let img = (rec.image && fresh.length) ? String(rec.image) : (idx > 0 ? String(v[idx][10] || "") : "");
    if (img.length > 48000) img = idx > 0 ? String(v[idx][10] || "") : "";

    const row = [key, cls, seat, txt_(name, 20), ACT, best, count, latest.at || "", txt_(latest.food, 40),
      JSON.stringify(latest.flavors || {}), img, js];
    if (idx > 0) sh.getRange(idx + 1, 1, 1, row.length).setValues([row]);
    else sh.appendRow(row);

    if (fresh.length) {
      const log = sheet_(LOG_SHEET, LOG_HEADERS);
      fresh.forEach(function (a) {
        const fl = a.flavors || {};
        const flTxt = Object.keys(FLAVOR_NAMES).map(function (k) { return FLAVOR_NAMES[k] + (Number(fl[k]) || 0) + "%"; }).join(" ");
        const notes = [];
        if ((Number(a.secs) || 0) < 180) notes.push("用時過短");
        if ((Number(a.chars) || 0) < 150) notes.push("字數偏少");
        log.appendRow([
          Utilities.formatDate(new Date(a.at), "Asia/Taipei", "yyyy/MM/dd HH:mm"),
          cls, seat, txt_(name, 20), txt_(a.food, 40), txt_((a.people || []).join("、"), 80),
          txt_(a.sense), txt_(a.core), flTxt, txt_(a.summary, 200),
          Number(a.chars) || 0, Number(a.secs) || 0, notes.join("、")
        ]);
      });
    }
    return { ok: true, added: fresh.length };
  } finally {
    lock.releaseLock();
  }
}

function clear_(req) {
  if (!auth_(req.password)) return { ok: false, error: "密碼錯誤" };
  const cls = String(req.cls || "").trim();
  if (!cls) return { ok: false, error: "缺少班級" };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    let removed = 0;
    [[DATA_SHEET, DATA_HEADERS], [LOG_SHEET, LOG_HEADERS]].forEach(function (p) {
      const sh = sheet_(p[0], p[1]);
      const v = sh.getDataRange().getValues();
      for (let i = v.length - 1; i >= 1; i--) {
        if (String(v[i][1]) === cls) { sh.deleteRow(i + 1); if (p[0] === DATA_SHEET) removed++; }
      }
    });
    return { ok: true, removed: removed };
  } finally {
    lock.releaseLock();
  }
}
