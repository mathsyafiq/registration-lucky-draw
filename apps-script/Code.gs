// ============================================================
// EVENT REGISTRATION & LUCKY DRAW - Code.gs (API backend)
// Deploy as Web App: Execute as "Me", Access "Anyone"
// Frontend is hosted on GitHub Pages (see /docs)
// ============================================================

const SHEETS = {
  settings:     "Settings",
  participants: "Participants",
  winners:      "Winners",
  departments:  "Departments"
};

const HEADERS = {
  Settings:     ["Key", "Value", "Notes"],
  Participants: ["Staff ID", "Name", "Department", "Status", "Registered At", "Method", "Family (Survey)", "Family (Present)", "Spouse Of"],
  Winners:      ["Draw No", "Staff ID", "Name", "Department", "Drawn At", "Status"],
  Departments:  ["Department"]
};

const REGISTERED = "Registered";
const WON = "WON";
const VOID = "VOID";
const WALK_IN = "Walk-in";
const P_COLS = 9;     // Participants columns A–I
const SPOUSE = "Spouse";
const FAMILY_MAX = 50;

// ── JSON helper ──────────────────────────────────────────────
function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── GET  ?action=info  → public event info ───────────────────
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || "info";
  try {
    if (action === "info") return json({ success: true, event: getPublicEvent() });
    if (action === "headcount") {
      const ev = getPublicEvent();
      return json({ success: true, event: { name: ev.name, date: ev.date }, headcount: getHeadcount() });
    }
    return json({ success: false, message: "Unknown action." });
  } catch (err) {
    return json({ success: false, message: String(err.message || err) });
  }
}

// ── POST (text/plain JSON body) → all other actions ──────────
function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return json({ success: false, message: "Invalid request." });
  }
  try {
    return json(route(body));
  } catch (err) {
    return json({ success: false, message: String(err.message || err) });
  }
}

// Actions anyone can call
const PUBLIC_ACTIONS = {
  register:    function (b) { return registerParticipant(b.staffId, "QR"); },
  addSpouse:      function (b) { return linkSpouse(b.staffId, b.spouseId, b.name, b.department, true); },
  registerSpouse: function (b) { return linkSpouse(b.headId, b.staffId, b.name, b.department, false); },
  registerNew: function (b) { return registerNew(b); },
  setFamily:   function (b) { return setFamily(b.staffId, b.count); }
};

// Actions that need the admin PIN
const ADMIN_ACTIONS = {
  login:              function ()  { return { success: true, event: getEvent() }; },
  dashboard:          function ()  { return getDashboard(); },
  participants:       function ()  { return { success: true, participants: readParticipants().map(publicParticipant) }; },
  importParticipants: function (b) { return importParticipants(b.rows, b.mode); },
  manualRegister:     function (b) { return registerParticipant(b.staffId, "Manual"); },
  unregister:         function (b) { return unregister(b.staffId); },
  saveSettings:       function (b) { return saveSettings(b); },
  saveDepartments:    function (b) { return saveDepartments(b.departments); },
  newEvent:           function (b) { return newEvent(b); },
  drawState:          function ()  { return getDrawState(); },
  setDrawCount:       function (b) { return setDrawCount(b.count); },
  draw:               function ()  { return drawWinner(); },
  voidWinner:         function (b) { return voidWinner(b.staffId); },
  resetDraw:          function ()  { return resetDraw(); }
};

function route(body) {
  const action = body && body.action;
  if (PUBLIC_ACTIONS[action]) return PUBLIC_ACTIONS[action](body);
  if (ADMIN_ACTIONS[action]) {
    const auth = checkPin(body.pin);
    if (!auth.success) return auth;
    return ADMIN_ACTIONS[action](body);
  }
  return { success: false, message: "Unknown action." };
}

// ── PIN check (locks out after too many wrong attempts) ──────
function checkPin(pin) {
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get("pinFails") || 0);
  if (fails >= 20) {
    return { success: false, auth: false, message: "Too many wrong PIN attempts. Try again in 10 minutes." };
  }
  if (String(pin || "").trim() !== "" && String(pin).trim() === getSetting("AdminPin")) {
    return { success: true };
  }
  cache.put("pinFails", String(fails + 1), 600);
  return { success: false, auth: false, message: "Wrong PIN." };
}

// ============================================================
// SETTINGS
// ============================================================
function defaultSettings() {
  return [
    ["EventName",        "My Event",            "Shown on every page"],
    ["EventDate",        "",                    "Free text, e.g. 12 Oct 2026"],
    ["AdminPin",         randomPin(),           "PIN for Admin and Lucky Draw pages. Change it!"],
    ["RegistrationOpen", "TRUE",                "TRUE or FALSE"],
    ["DrawCount",        "10",                  "Number of lucky draw prizes"],
    ["AnnounceLang",     "ms",                  "ms = Bahasa Melayu, en = English"],
    ["AllowWalkIn",      "TRUE",                "Staff not in the list may register themselves (TRUE/FALSE)"]
  ];
}

function randomPin() {
  let pin = "";
  for (let i = 0; i < 6; i++) pin += Math.floor(Math.random() * 10);
  return pin;
}

function getSettingsMap() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get("settings");
  if (cached) return JSON.parse(cached);

  const data = getSheet(SHEETS.settings).getDataRange().getValues();
  const map = {};
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) map[String(data[i][0]).trim()] = String(data[i][1]).trim();
  }
  cache.put("settings", JSON.stringify(map), 30);
  return map;
}

function getSetting(key) {
  const v = getSettingsMap()[key];
  return v === undefined ? "" : v;
}

function setSetting(key, value) {
  const sheet = getSheet(SHEETS.settings);
  const data = sheet.getDataRange().getValues();
  let row = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === key) { row = i + 1; break; }
  }
  if (row === -1) {
    sheet.appendRow([key, String(value), ""]);
  } else {
    sheet.getRange(row, 2).setValue(String(value));
  }
  CacheService.getScriptCache().remove("settings");
}

function isTrue(v) {
  return String(v).trim().toUpperCase() === "TRUE";
}

// On unless explicitly FALSE (older sheets don't have this setting yet)
function walkInAllowed(s) {
  return String(s.AllowWalkIn || "").trim().toUpperCase() !== "FALSE";
}

function getPublicEvent() {
  const s = getSettingsMap();
  return {
    name: s.EventName || "",
    date: s.EventDate || "",
    registrationOpen: isTrue(s.RegistrationOpen),
    allowWalkIn: walkInAllowed(s),
    departments: getDepartmentList().list
  };
}

function getEvent() {
  const s = getSettingsMap();
  return {
    name: s.EventName || "",
    date: s.EventDate || "",
    registrationOpen: isTrue(s.RegistrationOpen),
    drawCount: Number(s.DrawCount) || 0,
    announceLang: s.AnnounceLang || "ms",
    allowWalkIn: walkInAllowed(s)
  };
}

function saveSettings(b) {
  if (b.eventName !== undefined) setSetting("EventName", String(b.eventName).trim().slice(0, 150));
  if (b.eventDate !== undefined) setSetting("EventDate", String(b.eventDate).trim().slice(0, 60));
  if (b.registrationOpen !== undefined) setSetting("RegistrationOpen", b.registrationOpen ? "TRUE" : "FALSE");
  if (b.announceLang !== undefined) setSetting("AnnounceLang", b.announceLang === "en" ? "en" : "ms");
  if (b.allowWalkIn !== undefined) setSetting("AllowWalkIn", b.allowWalkIn ? "TRUE" : "FALSE");
  if (b.newPin !== undefined && b.newPin !== "") {
    const pin = String(b.newPin).trim();
    if (pin.length < 4) return { success: false, message: "PIN must be at least 4 characters." };
    setSetting("AdminPin", pin);
  }
  return { success: true, event: getEvent() };
}

// ============================================================
// PARTICIPANTS & REGISTRATION
// ============================================================

// Normalise a staff ID for matching: trim, uppercase, no spaces,
// and ignore leading zeros on purely numeric IDs.
function idKey(v) {
  let s = String(v === null || v === undefined ? "" : v).trim().toUpperCase().replace(/\s+/g, "");
  if (/^\d+$/.test(s)) s = s.replace(/^0+(?=\d)/, "");
  return s;
}

function readParticipants() {
  const sheet = getSheet(SHEETS.participants);
  ensureExtraHeaders(sheet);
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const data = sheet.getRange(2, 1, last - 1, P_COLS).getValues();
  const list = [];
  for (let i = 0; i < data.length; i++) {
    const r = data[i];
    if (String(r[0]).trim() === "") continue;
    list.push({
      row:          i + 2,
      staffId:      String(r[0]).trim(),
      name:         String(r[1]).trim(),
      department:   String(r[2]).trim(),
      registered:   String(r[3]).trim() === REGISTERED,
      registeredAt: r[4] instanceof Date ? r[4] : null,
      method:       String(r[5]).trim(),
      familyExpected: parseFamily(r[6]),
      familyPresent:  parseFamily(r[7]),
      linkedTo:       String(r[8]).trim()    // head of family's Staff ID, for a staff spouse
    });
  }
  return list;
}

// "3", 3, "3 orang" → 3; blank → null (unknown)
function parseFamily(v) {
  if (v === null || v === undefined || String(v).trim() === "") return null;
  const m = String(v).match(/\d+/);
  return m ? Math.min(FAMILY_MAX, Number(m[0])) : null;
}

// Older sheets were created before the family / spouse columns existed
function ensureExtraHeaders(sheet) {
  const cache = CacheService.getScriptCache();
  if (cache.get("extraHeaders")) return;
  const range = sheet.getRange(1, 7, 1, 3);
  const current = range.getValues()[0];
  const wanted = HEADERS.Participants.slice(6, 9);
  if (current.some(function (v) { return String(v).trim() === ""; })) {
    range.setValues([wanted.map(function (h, i) { return String(current[i]).trim() || h; })]);
    range.setBackground("#1a1a2e").setFontColor("#ffffff").setFontWeight("bold");
  }
  cache.put("extraHeaders", "1", 21600);
}

function publicParticipant(p) {
  return {
    staffId:      p.staffId,
    name:         p.name,
    department:   p.department,
    registered:   p.registered,
    registeredAt: p.registeredAt ? p.registeredAt.getTime() : null,
    method:       p.method,
    familyExpected: p.familyExpected,
    familyPresent:  p.familyPresent,
    linkedTo:       p.linkedTo
  };
}

// No lock here on purpose: each registration only writes its own row,
// so arrivals at peak time never queue behind each other.
function registerParticipant(staffId, method) {
  const key = idKey(staffId);
  if (!key) return { success: false, message: "Please enter your Staff ID." };
  if (key.length > 40) return { success: false, message: "Invalid Staff ID." };

  if (method === "QR" && !isTrue(getSetting("RegistrationOpen"))) {
    return { success: false, closed: true, message: "Registration is closed." };
  }

  const sheet = getSheet(SHEETS.participants);
  const last = sheet.getLastRow();
  if (last >= 2) {
    const data = sheet.getRange(2, 1, last - 1, P_COLS).getValues();
    for (let i = 0; i < data.length; i++) {
      if (idKey(data[i][0]) !== key) continue;

      const row = i + 2;
      const rec = data[i];
      const participant = {
        staffId:        String(rec[0]).trim(),
        name:           String(rec[1]).trim(),
        department:     String(rec[2]).trim(),
        familyExpected: parseFamily(rec[6]),
        familyPresent:  parseFamily(rec[7]),
        linkedTo:       String(rec[8]).trim(),
        linkedName:     "",
        spouses:        []
      };
      data.forEach(function (r) {
        if (participant.linkedTo && idKey(r[0]) === idKey(participant.linkedTo)) participant.linkedName = String(r[1]).trim();
        if (idKey(r[8]) === key && String(r[3]).trim() === REGISTERED) {
          participant.spouses.push({ staffId: String(r[0]).trim(), name: String(r[1]).trim() });
        }
      });
      if (String(rec[3]).trim() === REGISTERED) {
        return { success: true, already: true, participant: participant, registeredAt: formatTime(rec[4]) };
      }
      const now = new Date();
      sheet.getRange(row, 4, 1, 3).setValues([[REGISTERED, now, participant.linkedTo ? SPOUSE : method]]);
      return { success: true, already: false, participant: participant, registeredAt: formatTime(now) };
    }
  }
  return {
    success: false,
    notFound: true,
    message: "Staff ID not found. Please check your ID or see the registration counter."
  };
}

function unregister(staffId) {
  const key = idKey(staffId);
  const p = readParticipants().find(function (x) { return idKey(x.staffId) === key; });
  if (!p) return { success: false, message: "Staff ID not found." };
  const sheet = getSheet(SHEETS.participants);
  sheet.getRange(p.row, 4, 1, 3).clearContent();
  sheet.getRange(p.row, 8).clearContent();
  return { success: true };
}

// Family members who came with a registered participant (not counting them).
function setFamily(staffId, count) {
  const key = idKey(staffId);
  const n = Number(count);
  if (!key) return { success: false, message: "Please enter your Staff ID." };
  if (!(Math.floor(n) === n && n >= 0 && n <= FAMILY_MAX)) {
    return { success: false, message: "Please enter a number between 0 and " + FAMILY_MAX + "." };
  }
  const sheet = getSheet(SHEETS.participants);
  const last = sheet.getLastRow();
  if (last >= 2) {
    const ids = sheet.getRange(2, 1, last - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (idKey(ids[i][0]) !== key) continue;
      const row = i + 2;
      if (String(sheet.getRange(row, 4).getValue()).trim() !== REGISTERED) {
        return { success: false, message: "Please register first." };
      }
      ensureExtraHeaders(sheet);
      sheet.getRange(row, 8).setValue(n);
      return { success: true, familyPresent: n };
    }
  }
  return { success: false, message: "Staff ID not found." };
}

// Walk-in: staff whose ID is not in the imported list register themselves.
function registerNew(b) {
  const id = String(b.staffId || "").trim().slice(0, 40);
  const key = idKey(id);
  const name = String(b.name || "").trim().replace(/\s+/g, " ").slice(0, 150);
  const dept = String(b.department || "").trim().slice(0, 150);
  const family = Math.max(0, Math.min(FAMILY_MAX, Math.floor(Number(b.family) || 0)));
  if (!key) return { success: false, message: "Please enter your Staff ID." };
  if (!name) return { success: false, message: "Please enter your full name." };
  if (!dept) return { success: false, message: "Please choose your department." };

  const s = getSettingsMap();
  if (!isTrue(s.RegistrationOpen)) return { success: false, closed: true, message: "Registration is closed." };
  if (!walkInAllowed(s)) {
    return { success: false, message: "Only staff on the participant list can register. Please see the registration counter." };
  }
  const depts = getDepartmentList();
  if (depts.list.length && depts.list.indexOf(dept) === -1) {
    return { success: false, message: "Please choose a department from the list." };
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return { success: false, message: "System is busy. Please try again." };
  try {
    // Already on the list (e.g. typo the first time) → normal registration
    if (readParticipants().some(function (p) { return idKey(p.staffId) === key; })) {
      return registerParticipant(id, "QR");
    }
    const now = new Date();
    const sheet = getSheet(SHEETS.participants);
    const row = sheet.getLastRow() + 1;
    sheet.getRange(row, 1).setNumberFormat("@");
    ensureExtraHeaders(sheet);
    sheet.getRange(row, 1, 1, P_COLS).setValues([[id, name, dept, REGISTERED, now, WALK_IN, "", family, ""]]);
    return {
      success: true,
      already: false,
      walkIn: true,
      participant: { staffId: id, name: name, department: dept, familyExpected: null, familyPresent: family, linkedTo: "", spouses: [] },
      registeredAt: formatTime(now)
    };
  } finally {
    lock.releaseLock();
  }
}

// ── Staff spouse ─────────────────────────────────────────────
// The head of family's survey total already includes the spouse, so the
// spouse is linked to the head: in the lucky draw, but not counted again
// for food. byHead = true when the head adds the spouse at check-in;
// false when the spouse registers themselves.
function linkSpouse(headId, spouseId, name, department, byHead) {
  const hKey = idKey(headId), sKey = idKey(spouseId);
  name = String(name || "").trim().replace(/\s+/g, " ").slice(0, 150);
  department = String(department || "").trim().slice(0, 150);
  if (!hKey || !sKey) return { success: false, message: "Please enter both Staff IDs." };
  if (hKey === sKey) return { success: false, message: "The spouse's Staff ID must be different." };
  if (!isTrue(getSetting("RegistrationOpen"))) return { success: false, closed: true, message: "Registration is closed." };

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return { success: false, message: "System is busy. Please try again." };
  try {
    const ps = readParticipants();
    const find = function (k) { return ps.find(function (p) { return idKey(p.staffId) === k; }); };
    const head = find(hKey);
    if (!head) return { success: false, message: byHead ? "Staff ID not found." : "Your spouse's Staff ID is not in the participant list. Please check it, or register as a new participant." };
    if (head.linkedTo) return { success: false, message: head.staffId + " is registered as a spouse of " + head.linkedTo + ". Please use that Staff ID." };
    if (byHead && !head.registered) return { success: false, message: "Please register first." };

    const sheet = getSheet(SHEETS.participants);
    const now = new Date();
    let spouse = find(sKey);
    if (spouse) {
      if (spouse.linkedTo && idKey(spouse.linkedTo) !== hKey) {
        return { success: false, message: spouse.staffId + " is already registered with another family. Please see the registration counter." };
      }
      if (ps.some(function (p) { return idKey(p.linkedTo) === sKey; })) {
        return { success: false, message: spouse.staffId + " has registered their own family. Please see the registration counter." };
      }
      sheet.getRange(spouse.row, 4, 1, 3).setValues([[REGISTERED, spouse.registeredAt || now, spouse.registered ? spouse.method : SPOUSE]]);
      sheet.getRange(spouse.row, 9).setValue(head.staffId);
      spouse = { staffId: spouse.staffId, name: spouse.name, department: spouse.department };
    } else {
      if (!name) return { success: false, message: byHead ? "Please enter your spouse's full name." : "Please enter your full name." };
      if (!department) return { success: false, message: "Please choose a department." };
      const depts = getDepartmentList();
      if (depts.list.length && depts.list.indexOf(department) === -1) {
        return { success: false, message: "Please choose a department from the list." };
      }
      const id = String(spouseId).trim().slice(0, 40);
      const row = sheet.getLastRow() + 1;
      sheet.getRange(row, 1).setNumberFormat("@");
      sheet.getRange(row, 9).setNumberFormat("@");
      sheet.getRange(row, 1, 1, P_COLS).setValues([[id, name, department, REGISTERED, now, SPOUSE, "", "", head.staffId]]);
      spouse = { staffId: id, name: name, department: department };
    }

    // The head's confirmed group must include every linked spouse
    const linked = ps.filter(function (p) {
      return idKey(p.linkedTo) === hKey && p.registered && idKey(p.staffId) !== sKey;
    }).length + 1;
    let headFamily = head.familyPresent;
    if (head.registered && headFamily !== null && headFamily < linked) {
      headFamily = linked;
      sheet.getRange(head.row, 8).setValue(headFamily);
    }
    return {
      success: true,
      spouse: spouse,
      head: { staffId: head.staffId, name: head.name, registered: head.registered, familyPresent: headFamily },
      registeredAt: formatTime(now)
    };
  } finally {
    lock.releaseLock();
  }
}

// ── Departments for the walk-in form ─────────────────────────
// Uses the Departments tab; if that is empty, the departments
// found in the participant list.
function getDepartmentList() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get("departments");
  if (cached) return JSON.parse(cached);

  const seen = {};
  let list = [], source = "list";
  const sheet = getSheet(SHEETS.departments);
  const last = sheet.getLastRow();
  if (last >= 2) {
    sheet.getRange(2, 1, last - 1, 1).getValues().forEach(function (r) {
      const d = String(r[0]).trim();
      if (d && !seen[d]) { seen[d] = true; list.push(d); }
    });
  }
  if (!list.length) {
    source = "participants";
    readParticipants().forEach(function (p) {
      if (p.department && !seen[p.department]) { seen[p.department] = true; list.push(p.department); }
    });
    list.sort(function (a, b) { return a.localeCompare(b); });
  }
  const out = { list: list, source: source };
  cache.put("departments", JSON.stringify(out), 60);
  return out;
}

function saveDepartments(departments) {
  if (!Array.isArray(departments)) return { success: false, message: "Invalid list." };
  const seen = {};
  const rows = [];
  departments.forEach(function (d) {
    d = String(d || "").trim().slice(0, 150);
    if (d && !seen[d]) { seen[d] = true; rows.push([d]); }
  });
  if (rows.length > 2000) return { success: false, message: "Too many departments (max 2,000)." };

  const sheet = getSheet(SHEETS.departments);
  const last = sheet.getLastRow();
  if (last >= 2) sheet.getRange(2, 1, last - 1, 1).clearContent();
  if (rows.length) sheet.getRange(2, 1, rows.length, 1).setValues(rows);
  CacheService.getScriptCache().remove("departments");
  return { success: true, departmentList: getDepartmentList() };
}

// rows: [[staffId, name, department, familyFromSurvey], ...]
// mode: "replace" | "append" | "family" (only update the survey family numbers)
function importParticipants(rows, mode) {
  if (!Array.isArray(rows) || rows.length === 0) return { success: false, message: "No rows to import." };
  if (rows.length > 20000) return { success: false, message: "Too many rows (max 20,000)." };
  if (mode === "family") return importFamily(rows);

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return { success: false, message: "System is busy. Please try again." };
  try {
    const sheet = getSheet(SHEETS.participants);
    const existing = mode === "append" ? readParticipants() : [];
    const seen = {};
    existing.forEach(function (p) { seen[idKey(p.staffId)] = true; });

    const out = [];
    let duplicates = 0, skipped = 0;
    rows.forEach(function (r) {
      const id = String(r[0] === undefined ? "" : r[0]).trim().slice(0, 40);
      const key = idKey(id);
      if (!key) { skipped++; return; }
      if (seen[key]) { duplicates++; return; }
      seen[key] = true;
      out.push([
        id,
        String(r[1] === undefined ? "" : r[1]).trim().slice(0, 150),
        String(r[2] === undefined ? "" : r[2]).trim().slice(0, 150),
        "", "", "",
        parseFamily(r[3]) === null ? "" : parseFamily(r[3]),
        "", ""
      ]);
    });

    if (mode !== "append") {
      const last = sheet.getLastRow();
      if (last >= 2) sheet.getRange(2, 1, last - 1, P_COLS).clearContent();
    }
    ensureExtraHeaders(sheet);
    if (out.length) {
      const start = sheet.getLastRow() + 1;
      sheet.getRange(start, 1, out.length, 1).setNumberFormat("@");
      sheet.getRange(start, 1, out.length, P_COLS).setValues(out);
    }
    CacheService.getScriptCache().remove("departments");
    return {
      success: true,
      imported: out.length,
      duplicates: duplicates,
      skipped: skipped,
      total: readParticipants().length
    };
  } finally {
    lock.releaseLock();
  }
}

// Survey results: match by Staff ID and fill in "Family (Survey)".
function importFamily(rows) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return { success: false, message: "System is busy. Please try again." };
  try {
    const sheet = getSheet(SHEETS.participants);
    const last = sheet.getLastRow();
    if (last < 2) return { success: false, message: "Upload the participant list first, then the survey numbers." };
    ensureExtraHeaders(sheet);

    const ids = sheet.getRange(2, 1, last - 1, 1).getValues();
    const col = sheet.getRange(2, 7, last - 1, 1).getValues();
    const rowOf = {};
    ids.forEach(function (r, i) { if (String(r[0]).trim()) rowOf[idKey(r[0])] = i; });

    let updated = 0;
    const notFound = [];
    rows.forEach(function (r) {
      const key = idKey(r[0]);
      if (!key) return;
      const n = parseFamily(r[3]);
      if (n === null) return;
      if (rowOf[key] === undefined) { notFound.push(String(r[0]).trim()); return; }
      col[rowOf[key]][0] = n;
      updated++;
    });
    sheet.getRange(2, 7, last - 1, 1).setValues(col);
    return { success: true, updated: updated, notFound: notFound.length, notFoundIds: notFound.slice(0, 20) };
  } finally {
    lock.releaseLock();
  }
}

// ── Headcount (staff + family) for the food committee ───────
// A head of family's group total includes a staff spouse, so a linked
// spouse adds a staff member but no extra pax. Registered heads who
// haven't confirmed their group are counted with their survey number.
function getHeadcount(ps) {
  ps = ps || readParticipants();
  const byId = {}, spousesOf = {};
  ps.forEach(function (p) {
    byId[idKey(p.staffId)] = p;
    if (p.linkedTo && p.registered) spousesOf[idKey(p.linkedTo)] = (spousesOf[idKey(p.linkedTo)] || 0) + 1;
  });

  let staffPresent = 0, spousesPresent = 0, pax = 0, heads = 0;
  let familyConfirmed = 0, familyEstimated = 0, unconfirmed = 0, familyExpected = 0, surveyed = 0;
  ps.forEach(function (p) {
    if (p.linkedTo) {
      if (!p.registered) return;
      staffPresent++;
      spousesPresent++;
      const head = byId[idKey(p.linkedTo)];
      if (!head || !head.registered) pax++;   // head not here yet: count the spouse alone for now
      return;
    }
    heads++;
    if (p.familyExpected !== null) { familyExpected += p.familyExpected; surveyed++; }
    if (!p.registered) return;
    staffPresent++;
    const linked = spousesOf[idKey(p.staffId)] || 0;
    let fam;
    if (p.familyPresent !== null) { fam = Math.max(p.familyPresent, linked); familyConfirmed += fam; }
    else { fam = Math.max(p.familyExpected || 0, linked); familyEstimated += fam; unconfirmed++; }
    pax += 1 + fam;
  });
  return {
    staffPresent:    staffPresent,
    spousesPresent:  spousesPresent,
    familyPresent:   pax - staffPresent,   // family members who are not staff
    familyConfirmed: familyConfirmed,
    familyEstimated: familyEstimated,
    unconfirmed:     unconfirmed,
    paxPresent:      pax,
    staffExpected:   heads,
    familyExpected:  familyExpected,
    surveyed:        surveyed,
    paxExpected:     heads + familyExpected,
    updated:         formatTime(new Date())
  };
}

// Archive the current event into new tabs, then start fresh.
// keepParticipants = true keeps the staff list and only clears attendance.
function newEvent(b) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return { success: false, message: "System is busy. Please try again." };
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HHmm");
    const oldName = (getSetting("EventName") || "Event").replace(/[\[\]\*\?\/\\:']/g, "").slice(0, 40);

    const pSheet = getSheet(SHEETS.participants);
    const wSheet = getSheet(SHEETS.winners);

    if (pSheet.getLastRow() >= 2 || wSheet.getLastRow() >= 2) {
      pSheet.copyTo(ss).setName(uniqueSheetName(ss, stamp + " " + oldName + " - Attendance"));
      wSheet.copyTo(ss).setName(uniqueSheetName(ss, stamp + " " + oldName + " - Winners"));
    }

    const pLast = pSheet.getLastRow();
    if (pLast >= 2) {
      const block = pSheet.getRange(2, 1, pLast - 1, P_COLS);
      if (b.keepParticipants) {
        // keep the staff list; drop spouses added at check-in; clear attendance, family and links
        const kept = block.getValues()
          .filter(function (r) { return String(r[0]).trim() && String(r[5]).trim() !== SPOUSE && !String(r[8]).trim(); })
          .map(function (r) { return [r[0], r[1], r[2], "", "", "", "", "", ""]; });
        block.clearContent();
        if (kept.length) pSheet.getRange(2, 1, kept.length, P_COLS).setValues(kept);
      } else {
        block.clearContent();
      }
    }
    const wLast = wSheet.getLastRow();
    if (wLast >= 2) wSheet.getRange(2, 1, wLast - 1, 6).clearContent();

    if (b.eventName) setSetting("EventName", String(b.eventName).trim().slice(0, 150));
    setSetting("EventDate", String(b.eventDate || "").trim().slice(0, 60));
    setSetting("RegistrationOpen", "TRUE");
    CacheService.getScriptCache().remove("departments");
    return { success: true, event: getEvent() };
  } finally {
    lock.releaseLock();
  }
}

function uniqueSheetName(ss, base) {
  let name = base.slice(0, 95), n = 2;
  while (ss.getSheetByName(name)) name = base.slice(0, 90) + " (" + (n++) + ")";
  return name;
}

// ── Admin dashboard ──────────────────────────────────────────
function getDashboard() {
  const ps = readParticipants();
  const depts = {};
  let registered = 0, walkIns = 0;
  ps.forEach(function (p) {
    if (p.method === WALK_IN) walkIns++;
    const d = p.department || "(No department)";
    if (!depts[d]) depts[d] = { name: d, total: 0, registered: 0 };
    depts[d].total++;
    if (p.registered) { depts[d].registered++; registered++; }
  });

  const recent = ps
    .filter(function (p) { return p.registered && p.registeredAt; })
    .sort(function (a, b) { return b.registeredAt - a.registeredAt; })
    .slice(0, 15)
    .map(function (p) {
      return { staffId: p.staffId, name: p.name, department: p.department, time: formatTime(p.registeredAt), method: p.method };
    });

  return {
    success: true,
    event: getEvent(),
    stats: { total: ps.length, registered: registered, walkIns: walkIns },
    headcount: getHeadcount(ps),
    departmentList: getDepartmentList(),
    departments: Object.keys(depts).map(function (k) { return depts[k]; })
      .sort(function (a, b) { return b.total - a.total; }),
    recent: recent,
    winners: readWinners(),
    serverTime: formatTime(new Date())
  };
}

// ============================================================
// LUCKY DRAW
// ============================================================
function readWinners() {
  const sheet = getSheet(SHEETS.winners);
  const last = sheet.getLastRow();
  if (last < 2) return [];
  return sheet.getRange(2, 1, last - 1, 6).getValues()
    .filter(function (r) { return String(r[1]).trim() !== ""; })
    .map(function (r, i) {
      return {
        row:        i + 2,
        drawNo:     Number(r[0]) || 0,
        staffId:    String(r[1]).trim(),
        name:       String(r[2]).trim(),
        department: String(r[3]).trim(),
        drawnAt:    formatTime(r[4]),
        status:     String(r[5]).trim() || WON
      };
    });
}

function eligiblePool(participants, winners) {
  const taken = {};
  winners.forEach(function (w) { taken[idKey(w.staffId)] = true; });
  return participants.filter(function (p) { return p.registered && !taken[idKey(p.staffId)]; });
}

function getDrawState() {
  const ps = readParticipants();
  const winners = readWinners();
  const pool = eligiblePool(ps, winners);
  return {
    success: true,
    event: getEvent(),
    winners: winners,
    validCount: winners.filter(function (w) { return w.status !== VOID; }).length,
    pool: pool.map(function (p) { return p.staffId; })
  };
}

function setDrawCount(count) {
  const n = Math.floor(Number(count));
  if (!(n >= 1 && n <= 1000)) return { success: false, message: "Number of lucky draws must be between 1 and 1000." };
  setSetting("DrawCount", String(n));
  return { success: true, event: getEvent() };
}

// The winner is picked here on the server (not in the browser),
// recorded immediately, and then the wheel animates to it.
function drawWinner() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return { success: false, message: "System is busy. Please try again." };
  try {
    const drawCount = Number(getSetting("DrawCount")) || 0;
    const winners = readWinners();
    const validCount = winners.filter(function (w) { return w.status !== VOID; }).length;
    if (validCount >= drawCount) {
      return { success: false, done: true, message: "All " + drawCount + " lucky draws are complete." };
    }

    const pool = eligiblePool(readParticipants(), winners);
    if (pool.length === 0) return { success: false, message: "No eligible participants left in the draw." };

    const pick = pool[Math.floor(Math.random() * pool.length)];
    const drawNo = validCount + 1;
    const now = new Date();
    const sheet = getSheet(SHEETS.winners);
    const row = sheet.getLastRow() + 1;
    sheet.getRange(row, 2).setNumberFormat("@");
    sheet.getRange(row, 1, 1, 6).setValues([[drawNo, pick.staffId, pick.name, pick.department, now, WON]]);

    return {
      success: true,
      winner: { drawNo: drawNo, staffId: pick.staffId, name: pick.name, department: pick.department, drawnAt: formatTime(now) },
      validCount: drawNo,
      drawCount: drawCount,
      remaining: pool.length - 1
    };
  } finally {
    lock.releaseLock();
  }
}

// Winner not present: mark the draw VOID. The prize goes back into play,
// but the person stays out of the pool.
function voidWinner(staffId) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return { success: false, message: "System is busy. Please try again." };
  try {
    const key = idKey(staffId);
    const winners = readWinners();
    for (let i = winners.length - 1; i >= 0; i--) {
      if (idKey(winners[i].staffId) === key && winners[i].status !== VOID) {
        getSheet(SHEETS.winners).getRange(winners[i].row, 6).setValue(VOID);
        return { success: true };
      }
    }
    return { success: false, message: "Winner not found." };
  } finally {
    lock.releaseLock();
  }
}

function resetDraw() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return { success: false, message: "System is busy. Please try again." };
  try {
    const sheet = getSheet(SHEETS.winners);
    const last = sheet.getLastRow();
    if (last >= 2) sheet.getRange(2, 1, last - 1, 6).clearContent();
    return { success: true };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// HELPERS
// ============================================================
function formatTime(value) {
  if (!(value instanceof Date)) return value ? String(value) : "";
  return Utilities.formatDate(value, Session.getScriptTimeZone(), "d MMM, h:mm:ss a");
}

function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(name) || setupSheet(ss, name);
}

function setupSheet(ss, name) {
  const sheet = ss.insertSheet(name);
  const headers = HEADERS[name];
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground("#1a1a2e");
  headerRange.setFontColor("#ffffff");
  headerRange.setFontWeight("bold");
  sheet.setFrozenRows(1);

  if (name === SHEETS.settings) {
    const rows = defaultSettings();
    sheet.getRange(2, 2, rows.length, 1).setNumberFormat("@");
    sheet.getRange(2, 1, rows.length, 3).setValues(rows);
  }
  if (name === SHEETS.participants) sheet.getRange("A:A").setNumberFormat("@");
  if (name === SHEETS.winners) sheet.getRange("B:B").setNumberFormat("@");

  sheet.autoResizeColumns(1, headers.length);
  return sheet;
}

// Run this once manually to create the sheets. Your admin PIN is printed in the log.
function initialize() {
  getSheet(SHEETS.settings);
  getSheet(SHEETS.participants);
  getSheet(SHEETS.winners);
  getSheet(SHEETS.departments);
  CacheService.getScriptCache().remove("extraHeaders");
  ensureExtraHeaders(getSheet(SHEETS.participants));
  CacheService.getScriptCache().remove("settings");
  Logger.log("Sheets ready. Admin PIN: " + getSetting("AdminPin") + "  (change it in the Settings tab or the Admin page)");
}
