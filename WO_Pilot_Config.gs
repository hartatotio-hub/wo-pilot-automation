/***************************************
 * WO Pilot Automation (Approval + Offer + Done Album) - v5 (Split Files)
 * File: WO_Pilot_Config.gs
 ***************************************/

/** =========================
 *  1) DEFAULT CONFIG
 *  ========================= */
const DEFAULT_CONFIG = {
  // Sheets
  SHEET_WO_DB: "WO_DB",
  SHEET_TEAM: "TEAM_STATUS",
  SHEET_MATERIAL: "WO_MATERIAL",
  SHEET_LOG: "DISPATCH_LOG",

  // Calendar (pilot)
  CALENDAR_ID: "",

  // Telegram
  TELEGRAM_BOT_TOKEN: "",
  TELEGRAM_CHAT_ID: "",            // legacy: team group
  TELEGRAM_CHAT_TEAM_ID: "",       // preferred: team group
  TELEGRAM_CHAT_COORD_ID: "",      // coordinator private chat id
  TELEGRAM_CHAT_REALISASI_ID: "",  // group laporan realisasi

  // WO format
  WO_PREFIX: "BL/",
  WO_PAD: 3,

  // Support thresholds (konstruksi)
  SUPPORT_MIN_2TEAM: 5,  // 5-7 -> 2 teams (1 support)
  SUPPORT_MIN_3TEAM: 8,  // 8+  -> 3 teams (2 support)

  // Gangguan support rule
  GANGGUAN_SUPPORT_MIN_SET: 5, // >=5 -> 1 support (Total Set > 4)

  // Telegram polling
  POLL_EVERY_MINUTES: 1,

  // Misc
  DEFAULT_EVENT_DURATION_MIN: 60
};

// Field wajib (alias header yang diizinkan)
const REQUIRED_FIELDS = [
  { label: "Penyulang", keys: ["Penyulang", "PENYULANG"] }
  // (opsional kalau mau diwajibkan nanti)
  // { label: "UP3", keys: ["UP3"] },
  // { label: "ULP", keys: ["ULP"] }
];

/** =========================
 *  2) UI MENU
 *  ========================= */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("WO Pilot")
    .addItem("✅ Cek kesehatan sistem", "uiCheckHealth")
    .addSeparator()
    .addItem("⚙️ Setup/Save config (Script Properties)", "setupPilotConfig")
    .addItem("🧾 Lihat config (Script Properties)", "uiShowConfig")
    .addSeparator()
    .addItem("⚙️ Install trigger", "uiInstallTriggers")
    .addItem("🧹 Reset beban team", "uiResetTeamLoads")
    .addSeparator()
    .addItem("🧱 Ensure WO_DB headers (sekali)", "uiEnsureWoDbHeaders")
    .addToUi();
}

function uiInstallTriggers() {
  installTriggers();
  SpreadsheetApp.getUi().alert("OK", "Triggers installed: onFormSubmit + pollTelegramUpdates.", SpreadsheetApp.getUi().ButtonSet.OK);
}

function uiResetTeamLoads() {
  resetTeamLoads();
  SpreadsheetApp.getUi().alert("OK", "Team loads reset to 0 + gangguan rotation reset.", SpreadsheetApp.getUi().ButtonSet.OK);
}

function uiEnsureWoDbHeaders() {
  const cfg = getConfig_();
  const ss = getSpreadsheet_();
  const shWo = ss.getSheetByName(cfg.SHEET_WO_DB);
  if (!shWo) throw new Error(`Sheet not found: ${cfg.SHEET_WO_DB}`);
  ensureWoDbColumns_(shWo);
  SpreadsheetApp.getUi().alert("OK", "WO_DB headers ensured (kolom otomatis ditambah jika belum ada).", SpreadsheetApp.getUi().ButtonSet.OK);
}

function uiShowConfig() {
  const props = PropertiesService.getScriptProperties();
  const cfg = getConfig_();
  const lines = [];
  lines.push("CONFIG (resolved):");
  lines.push(JSON.stringify(cfg, null, 2));
  lines.push("");
  lines.push("SCRIPT PROPERTIES (raw keys):");
  const all = props.getProperties();
  Object.keys(all).sort().forEach(k => lines.push(`${k} = ${all[k]}`));
  SpreadsheetApp.getUi().alert(lines.join("\n"));
}

function uiCheckHealth() {
  const msg = checkHealth_();
  SpreadsheetApp.getUi().alert(msg);
}

/** =========================
 *  3) SETUP / TRIGGERS / RESET
 *  ========================= */
function setupPilotConfig() {
  const props = PropertiesService.getScriptProperties();
  const cfg = getConfig_();

  // store consolidated config
  props.setProperty("PILOT_CONFIG", JSON.stringify(cfg));

  // store single keys for backward compatibility and easy edit
  props.setProperty("CALENDAR_ID", cfg.CALENDAR_ID || "");
  props.setProperty("TELEGRAM_BOT_TOKEN", cfg.TELEGRAM_BOT_TOKEN || "");

  props.setProperty("TELEGRAM_CHAT_ID", String(cfg.TELEGRAM_CHAT_ID || ""));
  props.setProperty("TELEGRAM_CHAT_TEAM_ID", String(cfg.TELEGRAM_CHAT_TEAM_ID || ""));
  props.setProperty("TELEGRAM_CHAT_COORD_ID", String(cfg.TELEGRAM_CHAT_COORD_ID || ""));
  props.setProperty("TELEGRAM_CHAT_REALISASI_ID", String(cfg.TELEGRAM_CHAT_REALISASI_ID || ""));

  props.setProperty("WO_PREFIX", cfg.WO_PREFIX || "BL/");
  props.setProperty("WO_PAD", String(cfg.WO_PAD || 3));

  logInfo_("OK. Pilot config saved to Script Properties.");
}

function installTriggers() {
  const cfg = getConfig_();
  const ss = getSpreadsheet_();

  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === "pollTelegramUpdates" ||
        t.getHandlerFunction() === "onFormSubmit") {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger("onFormSubmit")
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();

  ScriptApp.newTrigger("pollTelegramUpdates")
    .timeBased()
    .everyMinutes(cfg.POLL_EVERY_MINUTES || 1)
    .create();

  logInfo_("OK. Triggers installed: onFormSubmit + pollTelegramUpdates.");
}

function resetTeamLoads() {
  const cfg = getConfig_();
  const ss = getSpreadsheet_();
  const shTeam = ss.getSheetByName(cfg.SHEET_TEAM);
  if (!shTeam) throw new Error(`Sheet not found: ${cfg.SHEET_TEAM}`);

  const { headerMap, values } = readTable_(shTeam);
  const colWo = mustCol_(headerMap, "ACTIVE_WO_COUNT");
  const colGg = mustCol_(headerMap, "ACTIVE_GANGGUAN_COUNT");
  const colLastAssigned = headerMap["LAST_ASSIGNED_AT"] ?? null;

  for (let i = 0; i < values.length; i++) {
    values[i][colWo] = 0;
    values[i][colGg] = 0;
    if (colLastAssigned !== null) values[i][colLastAssigned] = "";
  }
  writeTable_(shTeam, values);

  const props = PropertiesService.getScriptProperties();
  props.deleteProperty("GANGGUAN_LAST_TEAM_CODE");

  logInfo_("OK. Team loads reset to 0, gangguan rotation reset.");
}

/** =========================
 *  4) CONFIG RESOLUTION
 *  ========================= */
function getConfig_() {
  const props = PropertiesService.getScriptProperties();
  let cfg = { ...DEFAULT_CONFIG };

  const raw = props.getProperty("PILOT_CONFIG");
  if (raw) {
    try {
      const parsed = JSON.parse(raw);
      cfg = { ...cfg, ...parsed };
    } catch (e) {}
  }

  // override from individual keys (most important)
  const cal = props.getProperty("CALENDAR_ID");
  const token = props.getProperty("TELEGRAM_BOT_TOKEN");

  const teamLegacy = props.getProperty("TELEGRAM_CHAT_ID");
  const teamId = props.getProperty("TELEGRAM_CHAT_TEAM_ID");
  const coordId = props.getProperty("TELEGRAM_CHAT_COORD_ID");
  const realId = props.getProperty("TELEGRAM_CHAT_REALISASI_ID");

  const wp = props.getProperty("WO_PREFIX");
  const wpad = props.getProperty("WO_PAD");

  if (cal) cfg.CALENDAR_ID = cal;
  if (token) cfg.TELEGRAM_BOT_TOKEN = token;

  if (teamLegacy) cfg.TELEGRAM_CHAT_ID = teamLegacy;
  if (teamId) cfg.TELEGRAM_CHAT_TEAM_ID = teamId;
  if (coordId) cfg.TELEGRAM_CHAT_COORD_ID = coordId;
  if (realId) cfg.TELEGRAM_CHAT_REALISASI_ID = realId;

  if (wp) cfg.WO_PREFIX = wp;
  if (wpad) cfg.WO_PAD = Number(wpad);

  return cfg;
}

/** =========================
 *  5) HEALTH CHECK
 *  ========================= */
function checkHealth_() {
  const cfg = getConfig_();
  const ss = getSpreadsheet_();

  const lines = [];
  lines.push("WO Pilot - Health Check");

  // Sheets
  const requiredSheets = [cfg.SHEET_WO_DB, cfg.SHEET_TEAM, cfg.SHEET_LOG];
  requiredSheets.forEach(name => {
    const ok = !!ss.getSheetByName(name);
    lines.push(`${ok ? "✅" : "❌"} Sheet: ${name}`);
  });

  // Calendar
  if (!cfg.CALENDAR_ID) {
    lines.push("⚠️ CALENDAR_ID kosong (Calendar akan error saat dispatch).");
  } else {
    try {
      const cal = CalendarApp.getCalendarById(cfg.CALENDAR_ID);
      lines.push(`${cal ? "✅" : "❌"} Calendar ID valid`);
    } catch (e) {
      lines.push("❌ Calendar check error: " + (e && e.message ? e.message : e));
    }
  }

  // Telegram token
  if (!cfg.TELEGRAM_BOT_TOKEN) {
    lines.push("⚠️ TELEGRAM_BOT_TOKEN kosong (Telegram akan error).");
  } else {
    try {
      const me = telegramGetMe_(cfg.TELEGRAM_BOT_TOKEN);
      lines.push(`${me && me.ok ? "✅" : "❌"} Telegram token valid`);
    } catch (e) {
      lines.push("❌ Telegram token check error: " + (e && e.message ? e.message : e));
    }
  }

  // Chat IDs
  const teamChat = String(cfg.TELEGRAM_CHAT_TEAM_ID || cfg.TELEGRAM_CHAT_ID || "");
  if (!teamChat) lines.push("⚠️ TELEGRAM_CHAT_TEAM_ID/TELEGRAM_CHAT_ID kosong (dispatch ke group team akan error).");
  else lines.push("✅ Group team chat id terisi");

  if (!String(cfg.TELEGRAM_CHAT_COORD_ID || "")) lines.push("⚠️ TELEGRAM_CHAT_COORD_ID kosong (approval koordinator akan error).");
  else lines.push("✅ Koordinator chat id terisi");

  if (!String(cfg.TELEGRAM_CHAT_REALISASI_ID || "")) lines.push("⚠️ TELEGRAM_CHAT_REALISASI_ID kosong (laporan realisasi akan error).");
  else lines.push("✅ Group realisasi chat id terisi");

  lines.push("");
  lines.push("Catatan:");
  lines.push("- Jika missing field wajib (Penyulang), WO diblok (tidak diproses).");
  lines.push("- Approval: WO dikirim dulu ke koordinator (klik tombol).");
  lines.push("- DONE wajib album foto + SN1:, SN2:, dst.");

  return lines.join("\n");
}

/** =========================
 *  6) LOGGING
 *  ========================= */
function logInfo_(msg) { console.log(msg); }
function logError_(msg) { console.error(msg); }
