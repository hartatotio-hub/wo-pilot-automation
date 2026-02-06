/***************************************
 * WO Pilot Automation - v5 (Split Files)
 * File: WO_Pilot_Core.gs
 *
 * New features:
 * - Approval koordinator via tombol (inline keyboard)
 * - Offer ke group team via tombol (pilih team, slot sesuai Total Set)
 * - Manual pilih team (koordinator) menampilkan semua team
 * - DONE wajib album foto + SN1:, SN2:, dst. (multi foto dikirim ke grup realisasi sebagai album)
 ***************************************/

/** =========================
 *  1) TRIGGERS
 *  ========================= */
function onFormSubmit(e) {
  const cfg = getConfig_();
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const ss = getSpreadsheet_();
    const srcSheet = e && e.range ? e.range.getSheet() : null;
    const srcRow = e && e.range ? e.range.getRow() : null;

    if (!srcSheet || !srcRow) {
      logError_("onFormSubmit called without event object; abort.");
      return;
    }

    const shWo = ss.getSheetByName(cfg.SHEET_WO_DB);
    if (!shWo) throw new Error(`Sheet not found: ${cfg.SHEET_WO_DB}`);

    let woRowIndex = null;
    if (srcSheet.getName() === cfg.SHEET_WO_DB) {
      woRowIndex = srcRow;
    } else {
      woRowIndex = appendRowToWoDbFromSource_(srcSheet, srcRow, shWo);
    }

    processWoRow_(shWo, woRowIndex, { source: "FORM_SUBMIT" });

  } catch (err) {
    logError_("onFormSubmit error: " + (err && err.stack ? err.stack : err));
  } finally {
    lock.releaseLock();
  }
}

function pollTelegramUpdates() {
  const cfg = getConfig_();
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) return;

  try {
    const props = PropertiesService.getScriptProperties();
    const token = cfg.TELEGRAM_BOT_TOKEN;
    if (!token) throw new Error("TELEGRAM_BOT_TOKEN is empty.");

    const offset = Number(props.getProperty("TG_OFFSET") || 0);
    const updates = telegramGetUpdates_(token, offset);

    if (!updates || !updates.ok) {
      logError_("Telegram getUpdates failed: " + JSON.stringify(updates));
      return;
    }

    const results = updates.result || [];
    if (results.length === 0) return;

    const ss = getSpreadsheet_();
    const shWo = ss.getSheetByName(cfg.SHEET_WO_DB);
    if (!shWo) throw new Error(`Sheet not found: ${cfg.SHEET_WO_DB}`);

    // Pre-map media groups within this batch (for album done)
    const mediaGroups = {};
    for (const upd of results) {
      const msg = upd.message || upd.edited_message || null;
      if (!msg) continue;
      if (msg.media_group_id) {
        const gid = String(msg.media_group_id);
        if (!mediaGroups[gid]) mediaGroups[gid] = [];
        mediaGroups[gid].push(msg);
      }
    }

    let maxUpdateId = offset - 1;

    for (const upd of results) {
      if (typeof upd.update_id === "number") {
        maxUpdateId = Math.max(maxUpdateId, upd.update_id);
      }

      // 1) CALLBACK QUERIES (inline buttons)
      if (upd.callback_query) {
        handleCallbackQuery_(shWo, upd.callback_query);
        continue;
      }

      // 2) MESSAGES (reply commands in team group)
      const msg = upd.message || upd.edited_message || null;
      if (!msg) continue;

      const teamChatId = String(cfg.TELEGRAM_CHAT_TEAM_ID || cfg.TELEGRAM_CHAT_ID || "");
      if (!teamChatId) continue;

      // only process team group replies
      if (!msg.chat || String(msg.chat.id) !== String(teamChatId)) continue;

      // must be reply to WO message
      if (!msg.reply_to_message) continue;

      const repliedText = (msg.reply_to_message.text || msg.reply_to_message.caption || "").trim();
      const woId = extractWoIdFromText_(repliedText);
      if (!woId) continue;

      const by = msg.from && (msg.from.username || msg.from.first_name || "UNKNOWN")
        ? (msg.from.username || msg.from.first_name)
        : "UNKNOWN";

      // DONE ALBUM handling (caption)
      const captionOrText = String(msg.caption || msg.text || "").trim();
      if (isDoneCaption_(captionOrText)) {
        const gid = msg.media_group_id ? String(msg.media_group_id) : "";
        const albumMsgs = gid && mediaGroups[gid] ? mediaGroups[gid] : [msg];
        handleTeamDoneAlbum_(shWo, woId, albumMsgs, captionOrText, by, msg);
        continue;
      }

      // other commands (progress/cancel/reschedule/done without album)
      handleTelegramCommand_(shWo, woId, captionOrText, by, msg);
    }

    if (maxUpdateId >= 0) {
      props.setProperty("TG_OFFSET", String(maxUpdateId + 1));
    }

  } catch (err) {
    logError_("pollTelegramUpdates error: " + (err && err.stack ? err.stack : err));
  } finally {
    lock.releaseLock();
  }
}

/** =========================
 *  2) CORE PROCESSING (FORM -> COORD APPROVAL)
 *  ========================= */
function processWoRow_(shWo, rowIndex, meta) {
  const cfg = getConfig_();
  const ss = getSpreadsheet_();

  ensureWoDbColumns_(shWo);

  const { headerMap } = readHeader_(shWo);
  const row = shWo.getRange(rowIndex, 1, 1, shWo.getLastColumn()).getValues()[0];

  const woIdCol = mustCol_(headerMap, "WO_ID");
  const statusCol = mustCol_(headerMap, "WO_STATUS");
  const assignedCol = mustCol_(headerMap, "ASSIGNED_TO");
  const supportCol = mustCol_(headerMap, "SUPPORT_TEAM");

  const approvalCol = mustCol_(headerMap, "APPROVAL_STATUS");
  const coordMsgCol = mustCol_(headerMap, "COORD_TG_MESSAGE_ID");
  const coordSentCol = mustCol_(headerMap, "COORD_TG_SENT_AT");
  const coordResCol = mustCol_(headerMap, "COORD_TG_RESULT");

  const syncCol = mustCol_(headerMap, "SyncStatus");
  const lastUpCol = mustCol_(headerMap, "LAST_UPDATE_AT");
  const lastByCol = mustCol_(headerMap, "LAST_UPDATE_BY");

  // Skip if already sent to coordinator
  if (row[coordMsgCol]) {
    logInfo_(`Skip: row ${rowIndex} already has COORD_TG_MESSAGE_ID (${row[coordMsgCol]}).`);
    return;
  }

  // WO_ID
  let woId = String(row[woIdCol] || "").trim();
  if (!woId) {
    woId = nextWoId_();
    row[woIdCol] = woId;
  }

  // WO_STATUS
  let woStatus = String(row[statusCol] || "").trim();
  if (!woStatus) {
    woStatus = "Planned";
    row[statusCol] = woStatus;
  }

  // ✅ Validasi field wajib (mis. Penyulang)
  const missing = validateRequiredFields_(headerMap, row);
  if (missing.length) {
    row[syncCol] = "Missing required field: " + missing.join(", ");
    row[lastUpCol] = new Date();
    row[lastByCol] = "SYSTEM";
    shWo.getRange(rowIndex, 1, 1, shWo.getLastColumn()).setValues([row]);

    appendLog_(ss, cfg.SHEET_LOG, [
      new Date(),
      woId,
      "FORM_SUBMIT_BLOCKED",
      `Missing: ${missing.join(", ")}`,
      "SYSTEM"
    ]);
    return;
  }

  const jenis = getJenis_(headerMap, row);
  const totalSet = computeTotalSet_(headerMap, row);
  const requiredTeams = requiredTeamCountByTotalSet_(totalSet);

  // Preview assignment (auto)
  const assignment = assignTeams_(jenis, row, headerMap);
  row[assignedCol] = assignment.assignedTo || "-";
  row[supportCol] = assignment.supportTeams && assignment.supportTeams.length ? assignment.supportTeams.join(", ") : "-";

  // Approval state
  row[approvalCol] = "WAITING";
  setCellIfExists_(headerMap, row, "PICK_REQUIRED_TEAMS", requiredTeams);
  setCellIfExists_(headerMap, row, "PICK_SELECTED_TEAMS", "");
  setCellIfExists_(headerMap, row, "OFFER_STAGE", "");
  setCellIfExists_(headerMap, row, "MANUAL_STAGE", "");
  setCellIfExists_(headerMap, row, "MANUAL_SELECTED_TEAMS", "");

  row[syncCol] = "Waiting coordinator approval";
  row[lastUpCol] = new Date();
  row[lastByCol] = "SYSTEM";

  // write first (so coordinator actions can find row)
  shWo.getRange(rowIndex, 1, 1, shWo.getLastColumn()).setValues([row]);

  // Send to coordinator (with buttons)
  const coordId = String(cfg.TELEGRAM_CHAT_COORD_ID || "").trim();
  if (!coordId) throw new Error("TELEGRAM_CHAT_COORD_ID is empty (set Script Properties).");

  const msgText = buildCoordinatorMessage_(woId, row, headerMap, assignment, requiredTeams);

  const kb = tgInlineKeyboard_([
    [
      { text: "✅ Approve", callback_data: `C|A|${woId}` },
      { text: "❌ Reject", callback_data: `C|R|${woId}` }
    ],
    [
      { text: "🎯 Offer ke Group", callback_data: `C|O|${woId}` },
      { text: "🔧 Manual Pilih Team", callback_data: `C|M|${woId}` }
    ],
    [
      { text: "🚫 Cancel", callback_data: `C|X|${woId}` }
    ]
  ]);

  const tg = telegramSendMessage_(cfg.TELEGRAM_BOT_TOKEN, coordId, msgText, { replyMarkup: kb, disablePreview: true });
  if (tg && tg.ok && tg.result && tg.result.message_id) {
    row[coordMsgCol] = tg.result.message_id;
    row[coordSentCol] = new Date();
    row[coordResCol] = "OK";
  } else {
    row[coordResCol] = "ERROR";
    row[syncCol] = "Telegram Error (Coordinator): " + safeJson_(tg);
  }

  row[lastUpCol] = new Date();
  row[lastByCol] = "SYSTEM";
  shWo.getRange(rowIndex, 1, 1, shWo.getLastColumn()).setValues([row]);

  // optional: normalize materials to WO_MATERIAL
  try {
    syncMaterialsToSheet_(woId, row, headerMap);
  } catch (e) {
    logError_("syncMaterialsToSheet_ failed: " + (e && e.message ? e.message : e));
  }

  appendLog_(ss, cfg.SHEET_LOG, [
    new Date(),
    woId,
    "FORM_SUBMIT_TO_COORD",
    `Preview Assigned=${row[assignedCol]}, Support=${row[supportCol]}`,
    "SYSTEM"
  ]);
}

/** =========================
 *  3) CALLBACK HANDLER (COORD + GROUP OFFER)
 *  ========================= */
function handleCallbackQuery_(shWo, cb) {
  const cfg = getConfig_();
  const token = cfg.TELEGRAM_BOT_TOKEN;

  const data = String(cb.data || "");
  const parts = data.split("|");
  const kind = parts[0] || "";
  const action = parts[1] || "";
  const woId = parts[2] || "";

  const cbId = cb.id;
  const chatId = cb.message && cb.message.chat ? String(cb.message.chat.id) : "";
  const messageId = cb.message ? cb.message.message_id : null;

  const by = cb.from && (cb.from.username || cb.from.first_name || "UNKNOWN")
    ? (cb.from.username || cb.from.first_name)
    : "UNKNOWN";

  if (!kind || !action || !woId) {
    telegramAnswerCallbackQuery_(token, cbId, "Invalid button data.", false);
    return;
  }

  // Coordinator actions
  if (kind === "C") {
    const coordChat = String(cfg.TELEGRAM_CHAT_COORD_ID || "").trim();
    if (coordChat && chatId && String(chatId) !== String(coordChat)) {
      telegramAnswerCallbackQuery_(token, cbId, "Aksi ini hanya untuk koordinator.", true);
      return;
    }

    if (action === "A") {
      const ok = approveWo_(shWo, woId, by);
      telegramAnswerCallbackQuery_(token, cbId, ok ? "Approved ✅" : "Approve gagal.", false);
      // update coordinator message view
      if (messageId) {
        const text = buildCoordinatorStatusText_(shWo, woId, "APPROVED", by);
        telegramEditMessageText_(token, chatId, messageId, text, { replyMarkup: null, disablePreview: true });
      }
      return;
    }

    if (action === "R") {
      telegramAnswerCallbackQuery_(token, cbId, "Rejected. Pilih tindakan berikutnya.", false);
      if (messageId) {
        const text = buildCoordinatorStatusText_(shWo, woId, "REJECTED", by);
        const kb = tgInlineKeyboard_([
          [
            { text: "🎯 Offer ke Group", callback_data: `C|O|${woId}` },
            { text: "🔧 Manual Pilih Team", callback_data: `C|M|${woId}` }
          ],
          [
            { text: "🚫 Cancel", callback_data: `C|X|${woId}` },
            { text: "⬅️ Back", callback_data: `C|B|${woId}` }
          ]
        ]);
        telegramEditMessageText_(token, chatId, messageId, text, { replyMarkup: kb, disablePreview: true });
      }
      setWoApprovalStatus_(shWo, woId, "REJECTED", by);
      return;
    }

    if (action === "B") {
      telegramAnswerCallbackQuery_(token, cbId, "Back.", false);
      if (messageId) {
        // rebuild main coordinator message (with latest preview assignment)
        const ctx = getWoContext_(shWo, woId);
        if (!ctx) return;
        const { row, headerMap } = ctx;
        const jenis = getJenis_(headerMap, row);
        const totalSet = computeTotalSet_(headerMap, row);
        const requiredTeams = requiredTeamCountByTotalSet_(totalSet);
        const assignment = {
          assignedTo: String(row[mustCol_(headerMap, "ASSIGNED_TO")] || "-"),
          supportTeams: parseSupportTeams_(String(row[mustCol_(headerMap, "SUPPORT_TEAM")] || "-"))
        };
        const msgText = buildCoordinatorMessage_(woId, row, headerMap, assignment, requiredTeams);
        const kb = tgInlineKeyboard_([
          [
            { text: "✅ Approve", callback_data: `C|A|${woId}` },
            { text: "❌ Reject", callback_data: `C|R|${woId}` }
          ],
          [
            { text: "🎯 Offer ke Group", callback_data: `C|O|${woId}` },
            { text: "🔧 Manual Pilih Team", callback_data: `C|M|${woId}` }
          ],
          [
            { text: "🚫 Cancel", callback_data: `C|X|${woId}` }
          ]
        ]);
        telegramEditMessageText_(token, chatId, messageId, msgText, { replyMarkup: kb, disablePreview: true });
      }
      return;
    }

    if (action === "O") {
      telegramAnswerCallbackQuery_(token, cbId, "Mengirim penawaran ke group…", false);
      const ok = startOfferToGroup_(shWo, woId, by);
      if (messageId) {
        const text = buildCoordinatorStatusText_(shWo, woId, ok ? "OFFERING" : "OFFER_FAILED", by);
        const kb = tgInlineKeyboard_([
          [
            { text: "🔧 Manual Pilih Team", callback_data: `C|M|${woId}` },
            { text: "⬅️ Back", callback_data: `C|B|${woId}` }
          ],
          [
            { text: "🚫 Cancel", callback_data: `C|X|${woId}` }
          ]
        ]);
        telegramEditMessageText_(token, chatId, messageId, text, { replyMarkup: kb, disablePreview: true });
      }
      return;
    }

    if (action === "M") {
      telegramAnswerCallbackQuery_(token, cbId, "Mode manual: pilih team.", false);
      const ok = startManualPick_(shWo, woId, by, chatId, messageId);
      if (!ok) telegramAnswerCallbackQuery_(token, cbId, "Manual gagal (WO tidak ditemukan).", true);
      return;
    }

    if (action === "PM") {
      const teamCode = parts[3] || "";
      const ok = handleManualPickClick_(shWo, woId, teamCode, by, chatId, messageId, cbId);
      if (!ok) telegramAnswerCallbackQuery_(token, cbId, "Gagal memilih team.", true);
      return;
    }

    if (action === "X") {
      telegramAnswerCallbackQuery_(token, cbId, "Canceled.", false);
      cancelWo_(shWo, woId, by);
      if (messageId) {
        const text = buildCoordinatorStatusText_(shWo, woId, "CANCELED", by);
        telegramEditMessageText_(token, chatId, messageId, text, { replyMarkup: null, disablePreview: true });
      }
      return;
    }
  }

  // Group offer picks
  if (kind === "G" && action === "OF") {
    const teamCode = parts[3] || "";
    const ok = handleOfferPick_(shWo, woId, teamCode, by, cb, cbId);
    if (!ok) telegramAnswerCallbackQuery_(token, cbId, "Tidak bisa dipilih (sudah terisi/WO tidak aktif).", false);
    else telegramAnswerCallbackQuery_(token, cbId, `✅ Dipilih: ${teamCode}`, false);
    return;
  }

  telegramAnswerCallbackQuery_(token, cbId, "Unknown action.", false);
}

/** =========================
 *  4) APPROVE / OFFER / MANUAL / CANCEL
 *  ========================= */
function approveWo_(shWo, woId, by) {
  const cfg = getConfig_();
  const ss = getSpreadsheet_();
  ensureWoDbColumns_(shWo);

  const ctx = getWoContext_(shWo, woId);
  if (!ctx) return false;

  const { row, headerMap, idx, values } = ctx;

  const approvalCol = mustCol_(headerMap, "APPROVAL_STATUS");
  const statusCol = mustCol_(headerMap, "WO_STATUS");
  const assignedCol = mustCol_(headerMap, "ASSIGNED_TO");
  const supportCol = mustCol_(headerMap, "SUPPORT_TEAM");
  const teamMsgCol = mustCol_(headerMap, "TEAM_TG_MESSAGE_ID");
  const tgMsgLegacyCol = headerMap["TG_MESSAGE_ID"] ?? null;
  const syncCol = mustCol_(headerMap, "SyncStatus");
  const lastUpCol = mustCol_(headerMap, "LAST_UPDATE_AT");
  const lastByCol = mustCol_(headerMap, "LAST_UPDATE_BY");

  const curApproval = String(row[approvalCol] || "").trim().toUpperCase();
  if (curApproval === "APPROVED") return true;
  if (curApproval === "CANCELED") return false;

  const assignment = {
    assignedTo: String(row[assignedCol] || "-").trim(),
    supportTeams: parseSupportTeams_(String(row[supportCol] || "-"))
  };

  if (!assignment.assignedTo || assignment.assignedTo === "-") {
    values[idx][syncCol] = "Approve blocked: assigned team is empty. Use Manual/Offer.";
    writeTable_(shWo, values);
    return false;
  }

  // Dispatch to team group
  const teamChatId = String(cfg.TELEGRAM_CHAT_TEAM_ID || cfg.TELEGRAM_CHAT_ID || "");
  if (!teamChatId) throw new Error("TELEGRAM_CHAT_TEAM_ID / TELEGRAM_CHAT_ID is empty.");

  const msgText = buildTelegramMessage_(woId, row, headerMap, assignment);
  const tg = telegramSendMessage_(cfg.TELEGRAM_BOT_TOKEN, teamChatId, msgText, { disablePreview: false });

  if (tg && tg.ok && tg.result && tg.result.message_id) {
    values[idx][teamMsgCol] = tg.result.message_id;
    if (tgMsgLegacyCol !== null) values[idx][tgMsgLegacyCol] = tg.result.message_id;
    values[idx][syncCol] = "Dispatched to team group";
  } else {
    values[idx][syncCol] = "Telegram Error (Team): " + safeJson_(tg);
    writeTable_(shWo, values);
    return false;
  }

  // Apply load now (dispatch moment)
  const jenis = getJenis_(headerMap, row);
  applyLoadDelta_(assignment, jenis, +1);

  // Calendar create now (optional)
  try {
    const evId = ensureCalendarEvent_(cfg.CALENDAR_ID, woId, row, headerMap, assignment);
    setCellIfExists_(headerMap, values[idx], "CalendarEventId", evId || "");
  } catch (e) {
    // keep going
  }

  values[idx][approvalCol] = "APPROVED";
  setCellIfExists_(headerMap, values[idx], "APPROVED_AT", new Date());
  setCellIfExists_(headerMap, values[idx], "APPROVED_BY", by);

  values[idx][lastUpCol] = new Date();
  values[idx][lastByCol] = by;

  // keep status as planned/progress etc
  if (!values[idx][statusCol]) values[idx][statusCol] = "Planned";

  writeTable_(shWo, values);

  appendLog_(ss, cfg.SHEET_LOG, [
    new Date(),
    woId,
    "APPROVED_DISPATCHED",
    `Assigned=${assignment.assignedTo}, Support=${assignment.supportTeams.join(",") || "-"}`,
    by
  ]);

  return true;
}

function setWoApprovalStatus_(shWo, woId, status, by) {
  const ctx = getWoContext_(shWo, woId);
  if (!ctx) return;
  const { headerMap, values, idx } = ctx;
  const approvalCol = mustCol_(headerMap, "APPROVAL_STATUS");
  const lastUpCol = mustCol_(headerMap, "LAST_UPDATE_AT");
  const lastByCol = mustCol_(headerMap, "LAST_UPDATE_BY");
  values[idx][approvalCol] = String(status || "");
  values[idx][lastUpCol] = new Date();
  values[idx][lastByCol] = by;
  writeTable_(shWo, values);
}

function cancelWo_(shWo, woId, by) {
  const cfg = getConfig_();
  const ss = getSpreadsheet_();
  const ctx = getWoContext_(shWo, woId);
  if (!ctx) return;

  const { headerMap, values, idx, row } = ctx;
  const approvalCol = mustCol_(headerMap, "APPROVAL_STATUS");
  const statusCol = mustCol_(headerMap, "WO_STATUS");
  const assignedCol = mustCol_(headerMap, "ASSIGNED_TO");
  const supportCol = mustCol_(headerMap, "SUPPORT_TEAM");
  const syncCol = mustCol_(headerMap, "SyncStatus");
  const lastUpCol = mustCol_(headerMap, "LAST_UPDATE_AT");
  const lastByCol = mustCol_(headerMap, "LAST_UPDATE_BY");

  const oldStatus = String(values[idx][statusCol] || "").trim() || "Planned";
  const wasTerminal = isTerminalStatus_(oldStatus);

  values[idx][approvalCol] = "CANCELED";
  values[idx][statusCol] = "Cancel";
  values[idx][syncCol] = "Canceled by coordinator";
  values[idx][lastUpCol] = new Date();
  values[idx][lastByCol] = by;

  // If already dispatched and not terminal, release load
  const teamMsg = values[idx][mustCol_(headerMap, "TEAM_TG_MESSAGE_ID")] || "";
  if (teamMsg && !wasTerminal) {
    const assignment = {
      assignedTo: String(values[idx][assignedCol] || "-").trim(),
      supportTeams: parseSupportTeams_(String(values[idx][supportCol] || "-"))
    };
    const jenis = getJenis_(headerMap, row);
    applyLoadDelta_(assignment, jenis, -1);
  }

  writeTable_(shWo, values);

  appendLog_(ss, cfg.SHEET_LOG, [
    new Date(),
    woId,
    "CANCELED_BY_COORD",
    "-",
    by
  ]);
}

function startOfferToGroup_(shWo, woId, by) {
  const cfg = getConfig_();
  const ss = getSpreadsheet_();
  const ctx = getWoContext_(shWo, woId);
  if (!ctx) return false;

  const { row, headerMap, values, idx } = ctx;

  const approvalCol = mustCol_(headerMap, "APPROVAL_STATUS");
  const syncCol = mustCol_(headerMap, "SyncStatus");

  const totalSet = computeTotalSet_(headerMap, row);
  const requiredTeams = requiredTeamCountByTotalSet_(totalSet);

  setCellIfExists_(headerMap, values[idx], "PICK_REQUIRED_TEAMS", requiredTeams);
  setCellIfExists_(headerMap, values[idx], "PICK_SELECTED_TEAMS", "");
  setCellIfExists_(headerMap, values[idx], "OFFER_STAGE", "ASSIGNED");
  values[idx][approvalCol] = "OFFERING";
  values[idx][syncCol] = "Offering to group";

  writeTable_(shWo, values);

  // Send offer message to team group
  return sendOfferStageMessage_(shWo, woId, "ASSIGNED", by);
}

function handleOfferPick_(shWo, woId, teamCode, by, cb, cbId) {
  const cfg = getConfig_();
  const ss = getSpreadsheet_();
  const ctx = getWoContext_(shWo, woId);
  if (!ctx) return false;

  const { row, headerMap, values, idx } = ctx;
  const approvalCol = mustCol_(headerMap, "APPROVAL_STATUS");
  const approval = String(values[idx][approvalCol] || "").trim().toUpperCase();
  if (approval !== "OFFERING") return false;

  const requiredTeams = Number(getCell_(headerMap, values[idx], "PICK_REQUIRED_TEAMS") || 1);
  const selectedStr = String(getCell_(headerMap, values[idx], "PICK_SELECTED_TEAMS") || "").trim();
  const selected = selectedStr ? selectedStr.split(",").map(s => s.trim()).filter(Boolean) : [];

  if (!teamCode) return false;
  if (selected.includes(teamCode)) return false;
  if (selected.length >= requiredTeams) return false;

  selected.push(teamCode);
  setCellIfExists_(headerMap, values[idx], "PICK_SELECTED_TEAMS", selected.join(", "));

  // Close this offer message keyboard
  try {
    const offerMsg = cb.message;
    const txt = (offerMsg && offerMsg.text) ? offerMsg.text : "";
    const updatedTxt = txt + `\n\n✅ Dipilih: ${teamCode}`;
    telegramEditMessageText_(cfg.TELEGRAM_BOT_TOKEN, offerMsg.chat.id, offerMsg.message_id, updatedTxt, { replyMarkup: null, disablePreview: true });
  } catch (e) {}

  // If not complete, send next stage
  if (selected.length < requiredTeams) {
    const nextStage = selected.length === 1 ? "SUPPORT1" : "SUPPORT2";
    setCellIfExists_(headerMap, values[idx], "OFFER_STAGE", nextStage);
    writeTable_(shWo, values);
    sendOfferStageMessage_(shWo, woId, nextStage, by);
    return true;
  }

  // Complete -> auto dispatch
  setCellIfExists_(headerMap, values[idx], "OFFER_STAGE", "DONE");
  writeTable_(shWo, values);

  const assignment = {
    assignedTo: selected[0],
    supportTeams: selected.slice(1)
  };

  values[idx][mustCol_(headerMap, "ASSIGNED_TO")] = assignment.assignedTo;
  values[idx][mustCol_(headerMap, "SUPPORT_TEAM")] = assignment.supportTeams.length ? assignment.supportTeams.join(", ") : "-";
  writeTable_(shWo, values);

  const ok = approveWo_(shWo, woId, "AUTO_OFFER");
  appendLog_(ss, cfg.SHEET_LOG, [
    new Date(),
    woId,
    "OFFER_PICK_COMPLETE",
    `Picked=${selected.join(", ")}`,
    by
  ]);

  // notify coordinator (optional)
  if (ok) {
    const coordId = String(cfg.TELEGRAM_CHAT_COORD_ID || "").trim();
    if (coordId) telegramSendMessage_(cfg.TELEGRAM_BOT_TOKEN, coordId, `✅ WO ${woId} sudah ter-assign via OFFER: ${selected.join(", ")}`, { disablePreview: true });
  }

  return ok;
}

function sendOfferStageMessage_(shWo, woId, stage, by) {
  const cfg = getConfig_();
  const ctx = getWoContext_(shWo, woId);
  if (!ctx) return false;

  const { row, headerMap, values, idx } = ctx;
  const teamChatId = String(cfg.TELEGRAM_CHAT_TEAM_ID || cfg.TELEGRAM_CHAT_ID || "");
  if (!teamChatId) throw new Error("TELEGRAM_CHAT_TEAM_ID / TELEGRAM_CHAT_ID is empty.");

  const selectedStr = String(getCell_(headerMap, values[idx], "PICK_SELECTED_TEAMS") || "").trim();
  const selected = selectedStr ? selectedStr.split(",").map(s => s.trim()).filter(Boolean) : [];

  const jenis = getJenis_(headerMap, row);
  const totalSet = computeTotalSet_(headerMap, row);
  const requiredTeams = requiredTeamCountByTotalSet_(totalSet);

  const slotName = (stage === "ASSIGNED") ? "ASSIGNED (PIC)" : (stage === "SUPPORT1" ? "SUPPORT 1" : "SUPPORT 2");

  const candidates = getOfferCandidates_(jenis, row, headerMap, selected);
  if (candidates.length === 0) {
    // inform coordinator, fallback to manual
    const coordId = String(cfg.TELEGRAM_CHAT_COORD_ID || "").trim();
    if (coordId) telegramSendMessage_(cfg.TELEGRAM_BOT_TOKEN, coordId,
      `⚠️ OFFER gagal (tidak ada team TERSEDIA untuk slot ${slotName}) pada WO ${woId}. Silakan Manual pilih team.`,
      { disablePreview: true });
    return false;
  }

  const buttons = buildTeamButtons_(candidates, woId, "G|OF");
  const kb = tgInlineKeyboard_(buttons);

  const lines = [];
  lines.push("📣 OFFER WORK ORDER");
  lines.push(`WO_ID: ${woId}`);
  lines.push(`Slot: ${slotName}`);
  lines.push(`Total Set: ${Number(totalSet || 0)} | Butuh: ${requiredTeams} team`);
  if (selected.length) lines.push(`Sudah dipilih: ${selected.join(", ")}`);
  lines.push("");
  lines.push("Silakan klik nama team yang bersedia.");

  const tg = telegramSendMessage_(cfg.TELEGRAM_BOT_TOKEN, teamChatId, lines.join("\n"), { replyMarkup: kb, disablePreview: true });
  if (tg && tg.ok && tg.result && tg.result.message_id) {
    // store offer message id per stage
    const col = stage === "ASSIGNED" ? "OFFER_ASSIGNED_MSG_ID" : (stage === "SUPPORT1" ? "OFFER_SUP1_MSG_ID" : "OFFER_SUP2_MSG_ID");
    setCellIfExists_(headerMap, values[idx], col, tg.result.message_id);
    writeTable_(shWo, values);
    return true;
  }
  return false;
}

function startManualPick_(shWo, woId, by, coordChatId, coordMsgId) {
  const cfg = getConfig_();
  const ctx = getWoContext_(shWo, woId);
  if (!ctx) return false;

  const { row, headerMap, values, idx } = ctx;

  const totalSet = computeTotalSet_(headerMap, row);
  const requiredTeams = requiredTeamCountByTotalSet_(totalSet);

  setCellIfExists_(headerMap, values[idx], "PICK_REQUIRED_TEAMS", requiredTeams);
  setCellIfExists_(headerMap, values[idx], "MANUAL_STAGE", "ASSIGNED");
  setCellIfExists_(headerMap, values[idx], "MANUAL_SELECTED_TEAMS", "");

  values[idx][mustCol_(headerMap, "APPROVAL_STATUS")] = "MANUAL";

  writeTable_(shWo, values);

  // Show keyboard of ALL teams (manual)
  const teams = getAllTeams_(true); // include unavailable
  const kb = tgInlineKeyboard_(buildManualButtons_(teams, woId));

  const text = [
    `🔧 MANUAL PILIH TEAM`,
    `WO_ID: ${woId}`,
    `Slot: ASSIGNED (PIC)`,
    `Butuh total: ${requiredTeams} team`,
    ``,
    `Klik nama team untuk mengisi slot. (manual = tampil semua team)`,
  ].join("\n");

  telegramEditMessageText_(cfg.TELEGRAM_BOT_TOKEN, coordChatId, coordMsgId, text, { replyMarkup: kb, disablePreview: true });
  return true;
}

function handleManualPickClick_(shWo, woId, teamCode, by, coordChatId, coordMsgId, cbId) {
  const cfg = getConfig_();
  const ss = getSpreadsheet_();
  const ctx = getWoContext_(shWo, woId);
  if (!ctx) return false;

  const { row, headerMap, values, idx } = ctx;

  const requiredTeams = Number(getCell_(headerMap, values[idx], "PICK_REQUIRED_TEAMS") || 1);
  const stage = String(getCell_(headerMap, values[idx], "MANUAL_STAGE") || "ASSIGNED").trim().toUpperCase();
  const selStr = String(getCell_(headerMap, values[idx], "MANUAL_SELECTED_TEAMS") || "").trim();
  const selected = selStr ? selStr.split(",").map(s => s.trim()).filter(Boolean) : [];

  if (!teamCode) return false;
  if (selected.includes(teamCode)) {
    telegramAnswerCallbackQuery_(cfg.TELEGRAM_BOT_TOKEN, cbId, "Team sudah dipilih.", false);
    return true;
  }

  selected.push(teamCode);
  setCellIfExists_(headerMap, values[idx], "MANUAL_SELECTED_TEAMS", selected.join(", "));
  writeTable_(shWo, values);

  // Determine next step
  if (selected.length >= requiredTeams) {
    // finalize manual assignment -> dispatch
    const assignment = { assignedTo: selected[0], supportTeams: selected.slice(1) };
    values[idx][mustCol_(headerMap, "ASSIGNED_TO")] = assignment.assignedTo;
    values[idx][mustCol_(headerMap, "SUPPORT_TEAM")] = assignment.supportTeams.length ? assignment.supportTeams.join(", ") : "-";
    writeTable_(shWo, values);

    const ok = approveWo_(shWo, woId, "MANUAL_BY_COORD");
    const text = buildCoordinatorStatusText_(shWo, woId, ok ? "APPROVED" : "APPROVE_FAILED", by);
    telegramEditMessageText_(cfg.TELEGRAM_BOT_TOKEN, coordChatId, coordMsgId, text, { replyMarkup: null, disablePreview: true });

    appendLog_(ss, cfg.SHEET_LOG, [
      new Date(),
      woId,
      "MANUAL_PICK_COMPLETE",
      `Picked=${selected.join(", ")}`,
      by
    ]);
    return ok;
  }

  // need more teams (support slots)
  const nextSlot = selected.length === 1 ? "SUPPORT 1" : "SUPPORT 2";
  const teams = getAllTeams_(true); // include all
  const kb = tgInlineKeyboard_(buildManualButtons_(teams, woId));

  const text = [
    `🔧 MANUAL PILIH TEAM`,
    `WO_ID: ${woId}`,
    `Slot: ${nextSlot}`,
    `Butuh total: ${requiredTeams} team`,
    ``,
    `Terpilih sementara: ${selected.join(", ")}`,
    `Klik nama team berikutnya untuk mengisi slot.`,
  ].join("\n");

  telegramEditMessageText_(cfg.TELEGRAM_BOT_TOKEN, coordChatId, coordMsgId, text, { replyMarkup: kb, disablePreview: true });
  return true;
}

/** =========================
 *  5) TEAM REPLY: PROGRESS / CANCEL / RESCHEDULE / DONE(ALBUM)
 *  ========================= */
function handleTelegramCommand_(shWo, woId, cmdText, by, msgObj) {
  const cfg = getConfig_();
  const ss = getSpreadsheet_();

  ensureWoDbColumns_(shWo);

  const { headerMap, values } = readTable_(shWo);
  const woIdCol = mustCol_(headerMap, "WO_ID");
  const statusCol = mustCol_(headerMap, "WO_STATUS");
  const assignedCol = mustCol_(headerMap, "ASSIGNED_TO");
  const supportCol = mustCol_(headerMap, "SUPPORT_TEAM");
  const calIdCol = mustCol_(headerMap, "CalendarEventId");
  const syncCol = mustCol_(headerMap, "SyncStatus");
  const lastUpCol = mustCol_(headerMap, "LAST_UPDATE_AT");
  const lastByCol = mustCol_(headerMap, "LAST_UPDATE_BY");

  const idx = values.findIndex(r => String(r[woIdCol]).trim() === String(woId).trim());
  if (idx < 0) {
    // chat back to team group
    const teamChat = String(cfg.TELEGRAM_CHAT_TEAM_ID || cfg.TELEGRAM_CHAT_ID || "");
    telegramSendMessage_(cfg.TELEGRAM_BOT_TOKEN, teamChat, `❌ WO ${woId} tidak ditemukan di sheet.`, { disablePreview: true });
    return;
  }

  const row = values[idx];
  const oldStatus = String(row[statusCol] || "").trim() || "Planned";
  const parsed = parseCommand_(cmdText);
  if (!parsed) return;

  const jenis = getJenis_(headerMap, row);
  const jenisNorm = String(jenis || "").toLowerCase();
  const isGangguan = jenisNorm.includes("gangguan");

  if (parsed.type === "STATUS") {
    const newStatus = parsed.status;
    if (oldStatus.toLowerCase() === newStatus.toLowerCase()) return;

    // Done via text (tanpa foto) -> tolak
    if (newStatus.toLowerCase() === "done") {
      const teamChat = String(cfg.TELEGRAM_CHAT_TEAM_ID || cfg.TELEGRAM_CHAT_ID || "");
      telegramSendMessage_(cfg.TELEGRAM_BOT_TOKEN, teamChat,
        `❌ WO ${woId}: DONE wajib dikirim dengan ALBUM foto + caption SN1:, SN2:, dst.\nContoh:\nDone\nSN1: ...\nSN2: ...`,
        { replyToMessageId: msgObj && msgObj.message_id ? msgObj.message_id : null });
      return;
    }

    row[statusCol] = newStatus;
    row[lastUpCol] = new Date();
    row[lastByCol] = by;

    const assignment = {
      assignedTo: String(row[assignedCol] || "-").trim(),
      supportTeams: parseSupportTeams_(String(row[supportCol] || "-"))
    };

    const wasTerminal = isTerminalStatus_(oldStatus);
    const isTerminal = isTerminalStatus_(newStatus);

    if (!wasTerminal && isTerminal) {
      applyLoadDelta_(assignment, jenis, -1);
    }

    // ✅ Pindah MODE=PIKET saat gangguan masuk "On Progress" (opsi 2)
    if (isGangguan && newStatus.toLowerCase() === "on progress" && oldStatus.toLowerCase() !== "on progress") {
      try {
        const lastUsed = getLastUsedTeamForGangguanRotation_(assignment);
        if (lastUsed) {
          const nextPiket = computeNextGangguanTeam_(lastUsed);
          if (nextPiket) setOnlyOnePiket_(nextPiket);
        }
      } catch (e) {
        logError_("piket move error: " + (e && e.message ? e.message : e));
      }
    }

    try {
      const evId = String(row[calIdCol] || "");
      if (evId) {
        updateCalendarEvent_(cfg.CALENDAR_ID, evId, woId, row, headerMap, assignment);
        row[syncCol] = "Calendar Updated";
      }
    } catch (e) {
      row[syncCol] = "Calendar Update Error: " + (e && e.message ? e.message : String(e));
    }

    values[idx] = row;
    writeTable_(shWo, values);

    const teamChat = String(cfg.TELEGRAM_CHAT_TEAM_ID || cfg.TELEGRAM_CHAT_ID || "");
    telegramSendMessage_(cfg.TELEGRAM_BOT_TOKEN, teamChat,
      `✅ WO ${woId} status updated: ${oldStatus} → ${newStatus}`,
      { replyToMessageId: msgObj && msgObj.message_id ? msgObj.message_id : null });

    appendLog_(ss, cfg.SHEET_LOG, [
      new Date(),
      woId,
      "STATUS_CHANGED_FROM_TELEGRAM",
      `${oldStatus} -> ${newStatus}`,
      by
    ]);
  }

  if (parsed.type === "CANCEL") {
    const newStatus = "Cancel";
    const reason = parsed.reason || "-";

    const wasTerminal = isTerminalStatus_(oldStatus);
    row[statusCol] = newStatus;
    row[lastUpCol] = new Date();
    row[lastByCol] = by;

    const assignment = {
      assignedTo: String(row[assignedCol] || "-").trim(),
      supportTeams: parseSupportTeams_(String(row[supportCol] || "-"))
    };

    if (!wasTerminal) {
      applyLoadDelta_(assignment, jenis, -1);
    }

    try {
      const evId = String(row[calIdCol] || "");
      if (evId) {
        updateCalendarEvent_(cfg.CALENDAR_ID, evId, woId, row, headerMap, assignment, { cancelReason: reason });
        row[syncCol] = "Calendar Updated";
      }
    } catch (e) {
      row[syncCol] = "Calendar Update Error: " + (e && e.message ? e.message : String(e));
    }

    values[idx] = row;
    writeTable_(shWo, values);

    const teamChat = String(cfg.TELEGRAM_CHAT_TEAM_ID || cfg.TELEGRAM_CHAT_ID || "");
    telegramSendMessage_(cfg.TELEGRAM_BOT_TOKEN, teamChat,
      `✅ WO ${woId} status updated: ${oldStatus} → Cancel\nAlasan: ${reason}`,
      { replyToMessageId: msgObj && msgObj.message_id ? msgObj.message_id : null });

    appendLog_(ss, cfg.SHEET_LOG, [
      new Date(),
      woId,
      "CANCEL_FROM_TELEGRAM",
      `Reason: ${reason}`,
      by
    ]);
  }

  if (parsed.type === "RESCHEDULE") {
    const dt = parsed.datetime;
    setScheduleFields_(headerMap, row, dt);

    row[lastUpCol] = new Date();
    row[lastByCol] = by;

    try {
      const evId = String(row[calIdCol] || "");
      const assignment = {
        assignedTo: String(row[assignedCol] || "-").trim(),
        supportTeams: parseSupportTeams_(String(row[supportCol] || "-"))
      };
      if (evId) {
        updateCalendarEvent_(cfg.CALENDAR_ID, evId, woId, row, headerMap, assignment);
        row[syncCol] = "Calendar Updated";
      }
    } catch (e) {
      row[syncCol] = "Calendar Update Error: " + (e && e.message ? e.message : String(e));
    }

    values[idx] = row;
    writeTable_(shWo, values);

    const teamChat = String(cfg.TELEGRAM_CHAT_TEAM_ID || cfg.TELEGRAM_CHAT_ID || "");
    telegramSendMessage_(cfg.TELEGRAM_BOT_TOKEN, teamChat,
      `✅ WO ${woId} rescheduled to ${formatDateTime_(dt)}.`,
      { replyToMessageId: msgObj && msgObj.message_id ? msgObj.message_id : null });

    appendLog_(ss, cfg.SHEET_LOG, [
      new Date(),
      woId,
      "RESCHEDULE_FROM_TELEGRAM",
      formatDateTime_(dt),
      by
    ]);
  }
}

function handleTeamDoneAlbum_(shWo, woId, albumMsgs, captionOrText, by, msgObj) {
  const cfg = getConfig_();
  const ss = getSpreadsheet_();

  // Validate album has photo(s)
  const fileIds = extractPhotoFileIdsFromMessages_(albumMsgs);
  if (!fileIds.length) {
    const teamChat = String(cfg.TELEGRAM_CHAT_TEAM_ID || cfg.TELEGRAM_CHAT_ID || "");
    telegramSendMessage_(cfg.TELEGRAM_BOT_TOKEN, teamChat,
      `❌ WO ${woId}: DONE wajib melampirkan ALBUM FOTO.\nContoh caption:\nDone\nSN1: ...\nSN2: ...`,
      { replyToMessageId: msgObj && msgObj.message_id ? msgObj.message_id : null });
    return;
  }

  // Parse SN lines
  const snLines = parseSerialNumbersFromDone_(captionOrText);
  if (snLines.length === 0) {
    const teamChat = String(cfg.TELEGRAM_CHAT_TEAM_ID || cfg.TELEGRAM_CHAT_ID || "");
    telegramSendMessage_(cfg.TELEGRAM_BOT_TOKEN, teamChat,
      `❌ WO ${woId}: format SN wajib.\nContoh:\nDone\nSN1: 123\nSN2: 456`,
      { replyToMessageId: msgObj && msgObj.message_id ? msgObj.message_id : null });
    return;
  }

  // Load WO row
  ensureWoDbColumns_(shWo);
  const { headerMap, values } = readTable_(shWo);
  const woIdCol = mustCol_(headerMap, "WO_ID");
  const statusCol = mustCol_(headerMap, "WO_STATUS");
  const assignedCol = mustCol_(headerMap, "ASSIGNED_TO");
  const supportCol = mustCol_(headerMap, "SUPPORT_TEAM");
  const calIdCol = mustCol_(headerMap, "CalendarEventId");
  const syncCol = mustCol_(headerMap, "SyncStatus");
  const lastUpCol = mustCol_(headerMap, "LAST_UPDATE_AT");
  const lastByCol = mustCol_(headerMap, "LAST_UPDATE_BY");

  const idx = values.findIndex(r => String(r[woIdCol]).trim() === String(woId).trim());
  if (idx < 0) return;

  const row = values[idx];
  const oldStatus = String(row[statusCol] || "").trim() || "Planned";
  const wasTerminal = isTerminalStatus_(oldStatus);

  // Update status to Done (only affects load once)
  row[statusCol] = "Done";
  row[lastUpCol] = new Date();
  row[lastByCol] = by;

  const jenis = getJenis_(headerMap, row);
  const assignment = {
    assignedTo: String(row[assignedCol] || "-").trim(),
    supportTeams: parseSupportTeams_(String(row[supportCol] || "-"))
  };

  if (!wasTerminal) {
    applyLoadDelta_(assignment, jenis, -1);
  }

  // Store SN + photo evidence in WO_DB
  const snText = snLines.join("\n");
  const prevSn = String(getCell_(headerMap, row, "DONE_SERIALS") || "").trim();
  const mergedSn = prevSn ? (prevSn + "\n---\nTambahan oleh " + by + "\n" + snText) : snText;
  setCellIfExists_(headerMap, row, "DONE_SERIALS", mergedSn);
  setCellIfExists_(headerMap, row, "DONE_AT", new Date());
  setCellIfExists_(headerMap, row, "DONE_BY", by);

  const prevIds = String(getCell_(headerMap, row, "DONE_PHOTO_FILE_IDS") || "").trim();
  const mergedIds = prevIds ? (prevIds + "," + fileIds.join(",")) : fileIds.join(",");
  setCellIfExists_(headerMap, row, "DONE_PHOTO_FILE_IDS", mergedIds);
  setCellIfExists_(headerMap, row, "DONE_PHOTO_COUNT", (mergedIds ? mergedIds.split(",").filter(Boolean).length : fileIds.length));

  // Calendar update
  try {
    const evId = String(row[calIdCol] || "");
    if (evId) {
      updateCalendarEvent_(cfg.CALENDAR_ID, evId, woId, row, headerMap, assignment);
      row[syncCol] = "Calendar Updated";
    }
  } catch (e) {
    row[syncCol] = "Calendar Update Error: " + (e && e.message ? e.message : String(e));
  }

  values[idx] = row;
  writeTable_(shWo, values);

  // Bot reply thanks
  const teamChat = String(cfg.TELEGRAM_CHAT_TEAM_ID || cfg.TELEGRAM_CHAT_ID || "");
  telegramSendMessage_(cfg.TELEGRAM_BOT_TOKEN, teamChat,
    `✅ Terima kasih update.\nWO ${woId} ditandai DONE.\nSN:\n${snText}`,
    { replyToMessageId: msgObj && msgObj.message_id ? msgObj.message_id : null });

  // Send report to realisasi group (text + album)
  const realChat = String(cfg.TELEGRAM_CHAT_REALISASI_ID || "").trim();
  if (realChat) {
    const reportText = buildRealisasiReport_(woId, row, headerMap, assignment, snText, by, wasTerminal);
    telegramSendMessage_(cfg.TELEGRAM_BOT_TOKEN, realChat, reportText, { disablePreview: true });
    telegramSendMediaGroupChunked_(cfg.TELEGRAM_BOT_TOKEN, realChat, fileIds, {});
  }

  appendLog_(ss, cfg.SHEET_LOG, [
    new Date(),
    woId,
    wasTerminal ? "DONE_EVIDENCE_ADDED" : "DONE_FROM_TELEGRAM",
    `Photos=${fileIds.length}; SN=${snLines.length}`,
    by
  ]);
}

/** =========================
 *  6) ASSIGNMENT RULES (AUTO)
 *  ========================= */
function assignTeams_(jenis, woRow, woHeaderMap) {
  const cfg = getConfig_();
  const ss = getSpreadsheet_();
  const shTeam = ss.getSheetByName(cfg.SHEET_TEAM);
  if (!shTeam) throw new Error(`Sheet not found: ${cfg.SHEET_TEAM}`);

  const { headerMap, values } = readTable_(shTeam);

  const colTeam = mustCol_(headerMap, "TEAM_CODE");
  const colStatus = mustCol_(headerMap, "STATUS");
  const colMode = mustCol_(headerMap, "MODE");
  const colWo = mustCol_(headerMap, "ACTIVE_WO_COUNT");
  const colGg = mustCol_(headerMap, "ACTIVE_GANGGUAN_COUNT");

  const teams = values.map((r, i) => ({
    rowIndex: i,
    team: String(r[colTeam] || "").trim(),
    status: String(r[colStatus] || "").trim().toUpperCase(),
    mode: String(r[colMode] || "").trim().toUpperCase(),
    woCount: Number(r[colWo] || 0),
    ggCount: Number(r[colGg] || 0)
  })).filter(t => t.team);

  const jenisNorm = String(jenis || "").trim().toLowerCase();
  const isGangguan = jenisNorm.includes("gangguan");

  // Gangguan: same logic as v4 (rotation + optional support)
  if (isGangguan) {
    const piket = teams.find(t => t.mode === "PIKET" && t.status === "TERSEDIA") || null;
    if (!piket) {
      const anyAvail = teams.find(t => t.status === "TERSEDIA") || null;
      return { assignedTo: anyAvail ? anyAvail.team : "-", supportTeams: [] };
    }

    const props = PropertiesService.getScriptProperties();
    const lastTeamCode = (props.getProperty("GANGGUAN_LAST_TEAM_CODE") || "").trim();

    let assignedObj = null;
    if (!lastTeamCode) {
      assignedObj = piket;
    } else {
      const last = teams.find(t => t.team === lastTeamCode) || piket;
      assignedObj = nextAvailableTeamAfter_(teams, last.rowIndex) || piket;
    }

    const totalSet = computeTotalSet_(woHeaderMap, woRow);
    const supportTeams = [];

    if (Number(totalSet || 0) >= Number(cfg.GANGGUAN_SUPPORT_MIN_SET || 5)) {
      const supportObj = nextAvailableTeamAfter_(teams, assignedObj.rowIndex);
      if (supportObj && supportObj.team && supportObj.team !== assignedObj.team) {
        supportTeams.push(supportObj.team);
        props.setProperty("GANGGUAN_LAST_TEAM_CODE", supportObj.team); // rotasi maju sampai support
        return { assignedTo: assignedObj.team, supportTeams };
      }
    }

    props.setProperty("GANGGUAN_LAST_TEAM_CODE", assignedObj.team);
    return { assignedTo: assignedObj.team, supportTeams: [] };
  }

  // Konstruksi: NORMAL dulu, PIKET hanya untuk wilayah tertentu, dan hanya STATUS=TERSEDIA
  const totalSet = computeTotalSet_(woHeaderMap, woRow);
  const supportNeeded = supportCountByTotalSet_(totalSet);

  const allowedPiket = isPiketAllowedForKonstruksi_(woHeaderMap, woRow);

  const normalCandidates = teams.filter(t => t.mode === "NORMAL" && t.status === "TERSEDIA");
  const piketCandidates = teams.filter(t => t.mode === "PIKET" && t.status === "TERSEDIA" && allowedPiket);

  if (normalCandidates.length === 0 && piketCandidates.length === 0) {
    return { assignedTo: "-", supportTeams: [] };
  }

  normalCandidates.sort((a, b) => (a.woCount - b.woCount) || (a.rowIndex - b.rowIndex));
  piketCandidates.sort((a, b) => (a.woCount - b.woCount) || (a.rowIndex - b.rowIndex));

  const bestNormal = normalCandidates.length ? normalCandidates[0] : null;
  const bestPiket = piketCandidates.length ? piketCandidates[0] : null;

  const minWoNormal = bestNormal ? bestNormal.woCount : Infinity;
  const minWoPiket = bestPiket ? bestPiket.woCount : Infinity;

  const assigned = (bestPiket && (minWoPiket < minWoNormal)) ? bestPiket : (bestNormal || bestPiket);

  const supportTeams = [];
  if (supportNeeded > 0) {
    // 1) NORMAL dulu
    const normalSupportCandidates = teams
      .filter(t => t.mode === "NORMAL" && t.status === "TERSEDIA" && t.team !== assigned.team)
      .sort((a, b) => (a.woCount - b.woCount) || (a.rowIndex - b.rowIndex));

    for (const t of normalSupportCandidates) {
      if (supportTeams.length >= supportNeeded) break;
      supportTeams.push(t.team);
    }

    // 2) Kalau masih kurang, baru ambil dari PIKET (tapi hanya jika allowedPiket true)
    if (supportTeams.length < supportNeeded && allowedPiket) {
      const piketSupportCandidates = teams
        .filter(t => t.mode === "PIKET" && t.status === "TERSEDIA" && t.team !== assigned.team)
        .sort((a, b) => (a.woCount - b.woCount) || (a.rowIndex - b.rowIndex));

      for (const t of piketSupportCandidates) {
        if (supportTeams.length >= supportNeeded) break;
        if (!supportTeams.includes(t.team)) supportTeams.push(t.team);
      }
    }
  }

  return { assignedTo: assigned.team, supportTeams };
}

function isPiketAllowedForKonstruksi_(headerMap, row) {
  // rules from user:
  // - PIKET boleh untuk konstruksi:
  //   UP3 Bali Selatan + (ULP Sanur, Denpasar)
  //   UP3 Bali Timur + (ULP Gianyar)
  // - Bali Utara: jangan pernah konstruksi ke PIKET
  const up3 = String(getAny_(headerMap, row, ["UP3"]) || "").trim().toLowerCase();
  const ulp = String(getAny_(headerMap, row, ["ULP"]) || "").trim().toLowerCase();

  if (!up3 && !ulp) return false; // kalau tidak ada info, main aman: piket tidak dipakai

  if (up3.includes("utara")) return false;

  const isSelatan = up3.includes("selatan") && (ulp.includes("sanur") || ulp.includes("denpasar"));
  const isTimur = up3.includes("timur") && ulp.includes("gianyar");

  return isSelatan || isTimur;
}

function supportCountByTotalSet_(totalSet) {
  // user updated:
  // 1-4 set => 1 team
  // >4 set (5-7) => 2 team (1 support)
  // >7 set (8+) => 3 team (2 support)
  const cfg = getConfig_();
  const t = Number(totalSet || 0);
  if (t >= (cfg.SUPPORT_MIN_3TEAM || 8)) return 2;
  if (t >= (cfg.SUPPORT_MIN_2TEAM || 5)) return 1;
  return 0;
}

function requiredTeamCountByTotalSet_(totalSet) {
  return 1 + supportCountByTotalSet_(totalSet);
}

function nextAvailableTeamAfter_(teams, startIndex) {
  const n = teams.length;
  for (let step = 1; step <= n; step++) {
    const idx = (startIndex + step) % n;
    if (teams[idx].status === "TERSEDIA") return teams[idx];
  }
  return null;
}

/** Offer candidates: ONLY TERSEDIA, and apply piket restriction only for konstruksi */
function getOfferCandidates_(jenis, woRow, woHeaderMap, excludeTeams) {
  const cfg = getConfig_();
  const ss = getSpreadsheet_();
  const shTeam = ss.getSheetByName(cfg.SHEET_TEAM);
  if (!shTeam) throw new Error(`Sheet not found: ${cfg.SHEET_TEAM}`);

  const { headerMap, values } = readTable_(shTeam);
  const colTeam = mustCol_(headerMap, "TEAM_CODE");
  const colStatus = mustCol_(headerMap, "STATUS");
  const colMode = mustCol_(headerMap, "MODE");

  const jenisNorm = String(jenis || "").trim().toLowerCase();
  const isGangguan = jenisNorm.includes("gangguan");

  const allowedPiket = isGangguan ? true : isPiketAllowedForKonstruksi_(woHeaderMap, woRow);

  const ex = new Set((excludeTeams || []).map(x => String(x).trim()).filter(Boolean));

  return values.map(r => ({
    team: String(r[colTeam] || "").trim(),
    status: String(r[colStatus] || "").trim().toUpperCase(),
    mode: String(r[colMode] || "").trim().toUpperCase()
  })).filter(t => t.team && t.status === "TERSEDIA" && !ex.has(t.team))
    .filter(t => {
      if (isGangguan) return true;
      if (t.mode === "PIKET" && !allowedPiket) return false;
      return true;
    })
    .map(t => t.team);
}

/** Manual: show all teams (includeUnavailable=true) */
function getAllTeams_(includeUnavailable) {
  const cfg = getConfig_();
  const ss = getSpreadsheet_();
  const shTeam = ss.getSheetByName(cfg.SHEET_TEAM);
  if (!shTeam) throw new Error(`Sheet not found: ${cfg.SHEET_TEAM}`);

  const { headerMap, values } = readTable_(shTeam);
  const colTeam = mustCol_(headerMap, "TEAM_CODE");
  const colStatus = mustCol_(headerMap, "STATUS");
  const colMode = mustCol_(headerMap, "MODE");

  const list = values.map(r => ({
    team: String(r[colTeam] || "").trim(),
    status: String(r[colStatus] || "").trim(),
    mode: String(r[colMode] || "").trim()
  })).filter(t => t.team);

  return includeUnavailable ? list : list.filter(t => String(t.status).toUpperCase() === "TERSEDIA");
}

/** Build buttons: 2 columns */
function buildTeamButtons_(teamCodes, woId, prefix) {
  const rows = [];
  let cur = [];
  for (const t of teamCodes) {
    cur.push({ text: t, callback_data: `${prefix}|${woId}|${t}` });
    if (cur.length >= 2) { rows.push(cur); cur = []; }
  }
  if (cur.length) rows.push(cur);
  return rows;
}

/** Manual buttons: show label with status/mode, but callback carries teamCode */
function buildManualButtons_(teams, woId) {
  const rows = [];
  let cur = [];
  for (const t of teams) {
    const label = `${t.team} (${String(t.status || "-")}/${String(t.mode || "-")})`;
    cur.push({ text: label, callback_data: `C|PM|${woId}|${t.team}` });
    if (cur.length >= 2) { rows.push(cur); cur = []; }
  }
  if (cur.length) rows.push(cur);
  // add back button
  rows.push([{ text: "⬅️ Back", callback_data: `C|B|${woId}` }]);
  return rows;
}

/** =========================
 *  7) TELEGRAM MESSAGE BUILDERS
 *  ========================= */
function buildCoordinatorMessage_(woId, row, headerMap, assignment, requiredTeams) {
  const status = String(row[headerMap["WO_STATUS"]] || "Planned").trim();
  const jenis = getJenis_(headerMap, row) || "-";
  const pekerjaan = getAny_(headerMap, row, ["NAMA PEKERJAAN", "PEKERJAAN", "Nama Pekerjaan", "Nama PEKERJAAN"]) || "-";
  const lokasi = getAny_(headerMap, row, ["Lokasi", "LOKASI"]) || "-";
  const pengawas = getAny_(headerMap, row, ["Nama Pengawas", "PENGAWAS", "Pengawas"]) || "-";
  const penyulang = getAny_(headerMap, row, ["Penyulang", "PENYULANG"]) || "-";
  const up3 = getAny_(headerMap, row, ["UP3"]) || "-";
  const ulp = getAny_(headerMap, row, ["ULP"]) || "-";

  const mapsUrl = getMapsUrl_(headerMap, row);
  const schedule = getScheduleText_(headerMap, row);

  const assignedTo = assignment.assignedTo || "-";
  const supports = assignment.supportTeams && assignment.supportTeams.length ? assignment.supportTeams.join(", ") : "-";
  const totalSet = computeTotalSet_(headerMap, row);

  const materialsText = buildMaterialsText_(headerMap, row);

  const warn = [];
  if (assignedTo === "-" || assignedTo === "") warn.push("⚠️ Auto-assign tidak menemukan team. Gunakan Manual/Offer.");
  if (requiredTeams > 1 && (supports === "-" || supports === "")) warn.push("⚠️ Support belum lengkap. Bisa Offer/Manual.");

  const lines = [];
  lines.push("🧑‍💼 APPROVAL WO (Koordinator)");
  lines.push(`WO_ID: ${woId}`);
  lines.push(`Status: ${status}`);
  lines.push(`Jenis: ${jenis}`);
  lines.push(`Pekerjaan: ${pekerjaan}`);
  lines.push(`Penyulang: ${penyulang}`);
  lines.push(`UP3: ${up3}`);
  lines.push(`ULP: ${ulp}`);
  lines.push(`Lokasi: ${lokasi}`);
  lines.push(`Pengawas: ${pengawas}`);
  lines.push(`Maps: ${mapsUrl || "-"}`);
  lines.push("");
  lines.push("📅 Jadwal");
  lines.push(schedule);
  lines.push("");
  lines.push("👥 Preview Dispatch");
  lines.push(`Assigned: ${assignedTo}`);
  lines.push(`Support: ${supports}`);
  lines.push(`Total Set: ${Number(totalSet || 0)} | Butuh: ${requiredTeams} team`);
  if (warn.length) {
    lines.push("");
    warn.forEach(w => lines.push(w));
  }
  lines.push("");
  lines.push("📦 Material:");
  lines.push(materialsText);

  return lines.join("\n");
}

function buildCoordinatorStatusText_(shWo, woId, status, by) {
  const ctx = getWoContext_(shWo, woId);
  if (!ctx) return `WO ${woId} tidak ditemukan.`;
  const { row, headerMap } = ctx;

  const jenis = getJenis_(headerMap, row) || "-";
  const totalSet = computeTotalSet_(headerMap, row);
  const req = requiredTeamCountByTotalSet_(totalSet);
  const assigned = String(row[mustCol_(headerMap, "ASSIGNED_TO")] || "-");
  const support = String(row[mustCol_(headerMap, "SUPPORT_TEAM")] || "-");

  const lines = [];
  lines.push("🧑‍💼 WO STATUS (Koordinator)");
  lines.push(`WO_ID: ${woId}`);
  lines.push(`State: ${status}`);
  lines.push(`By: ${by}`);
  lines.push("");
  lines.push(`Jenis: ${jenis}`);
  lines.push(`Total Set: ${Number(totalSet || 0)} | Butuh: ${req} team`);
  lines.push(`Assigned: ${assigned}`);
  lines.push(`Support: ${support}`);
  lines.push("");
  lines.push("Catatan: jika Offer gagal/kurang, gunakan Manual untuk pilih team.");
  return lines.join("\n");
}

function buildTelegramMessage_(woId, row, headerMap, assignment) {
  const status = String(row[headerMap["WO_STATUS"]] || "Planned").trim();
  const jenis = getJenis_(headerMap, row) || "-";
  const pekerjaan = getAny_(headerMap, row, ["NAMA PEKERJAAN", "PEKERJAAN", "Nama Pekerjaan", "Nama PEKERJAAN"]) || "-";
  const lokasi = getAny_(headerMap, row, ["Lokasi", "LOKASI"]) || "-";
  const pengawas = getAny_(headerMap, row, ["Nama Pengawas", "PENGAWAS", "Pengawas"]) || "-";
  const penyulang = getAny_(headerMap, row, ["Penyulang", "PENYULANG"]) || "-";
  const up3 = getAny_(headerMap, row, ["UP3"]) || "-";
  const ulp = getAny_(headerMap, row, ["ULP"]) || "-";

  const mapsUrl = getMapsUrl_(headerMap, row);
  const schedule = getScheduleText_(headerMap, row);

  const assignedTo = assignment.assignedTo || "-";
  const supports = assignment.supportTeams && assignment.supportTeams.length ? assignment.supportTeams.join(", ") : "-";
  const totalSet = computeTotalSet_(headerMap, row);

  const materialsText = buildMaterialsText_(headerMap, row);

  const lines = [];
  lines.push("📝 WORK ORDER");
  lines.push(`WO_ID: ${woId}`);
  lines.push(`Status: ${status}`);
  lines.push(`Jenis: ${jenis}`);
  lines.push(`Pekerjaan: ${pekerjaan}`);
  lines.push(`Penyulang: ${penyulang}`);
  lines.push(`UP3: ${up3}`);
  lines.push(`ULP: ${ulp}`);
  lines.push(`Lokasi: ${lokasi}`);
  lines.push(`Pengawas: ${pengawas}`);
  lines.push(`Maps: ${mapsUrl || "-"}`);
  lines.push("");
  lines.push("📅 Jadwal");
  lines.push(schedule);
  lines.push("");
  lines.push("👥 Dispatch");
  lines.push(`Assigned: ${assignedTo}`);
  lines.push(`Support: ${supports}`);
  lines.push(`Total Set: ${Number(totalSet || 0)}`);
  lines.push("");
  lines.push("📦 Material:");
  lines.push(materialsText);
  lines.push("");
  lines.push("✅ Update status (reply pesan ini): progress | cancel: alasan | reschedule: dd/mm/yyyy HH:mm");
  lines.push("✅ DONE WAJIB: reply dengan ALBUM FOTO + caption:");
  lines.push("Done");
  lines.push("SN1: ....");
  lines.push("SN2: ....");
  lines.push("(dst sesuai jumlah material yang dipasang)");
  return lines.join("\n");
}

function buildRealisasiReport_(woId, row, headerMap, assignment, snText, by, wasTerminal) {
  const jenis = getJenis_(headerMap, row) || "-";
  const pekerjaan = getAny_(headerMap, row, ["NAMA PEKERJAAN", "PEKERJAAN", "Nama Pekerjaan", "Nama PEKERJAAN"]) || "-";
  const penyulang = getAny_(headerMap, row, ["Penyulang", "PENYULANG"]) || "-";
  const up3 = getAny_(headerMap, row, ["UP3"]) || "-";
  const ulp = getAny_(headerMap, row, ["ULP"]) || "-";
  const lokasi = getAny_(headerMap, row, ["Lokasi", "LOKASI"]) || "-";
  const totalSet = computeTotalSet_(headerMap, row);

  const label = wasTerminal ? "📎 Tambahan Bukti" : "✅ Laporan Realisasi";
  const lines = [];
  lines.push(label);
  lines.push(`WO_ID: ${woId}`);
  lines.push(`Jenis: ${jenis}`);
  lines.push(`Pekerjaan: ${pekerjaan}`);
  lines.push(`Penyulang: ${penyulang}`);
  lines.push(`UP3: ${up3}`);
  lines.push(`ULP: ${ulp}`);
  lines.push(`Lokasi: ${lokasi}`);
  lines.push(`Total Set: ${Number(totalSet || 0)}`);
  lines.push(`Assigned: ${assignment.assignedTo || "-"}`);
  lines.push(`Support: ${(assignment.supportTeams && assignment.supportTeams.length) ? assignment.supportTeams.join(", ") : "-"}`);
  lines.push(`Update oleh: ${by}`);
  lines.push("");
  lines.push("SN:");
  lines.push(snText || "-");
  return lines.join("\n");
}

/** =========================
 *  8) DONE PARSING & PHOTO EXTRACT
 *  ========================= */
function isDoneCaption_(text) {
  const t = String(text || "").trim();
  if (!t) return false;
  const firstLine = t.split("\n")[0].trim().toLowerCase();
  return firstLine === "done" || firstLine === "selesai";
}

function parseSerialNumbersFromDone_(text) {
  const lines = String(text || "").split("\n").map(s => s.trim()).filter(Boolean);
  const out = [];
  for (const ln of lines) {
    const m = ln.match(/^SN\s*(\d+)\s*:\s*(.+)$/i);
    if (m && m[1] && m[2]) {
      out.push(`SN${m[1]}: ${m[2].trim()}`);
    }
  }
  return out;
}

function extractPhotoFileIdsFromMessages_(msgs) {
  const ids = [];
  for (const m of (msgs || [])) {
    // photo
    if (m.photo && m.photo.length) {
      const best = m.photo[m.photo.length - 1];
      if (best && best.file_id) ids.push(String(best.file_id));
      continue;
    }
    // image document
    if (m.document && m.document.file_id) {
      const mime = String(m.document.mime_type || "");
      if (mime.startsWith("image/")) ids.push(String(m.document.file_id));
    }
  }
  // unique
  return Array.from(new Set(ids));
}

/** =========================
 *  9) LOAD UPDATE
 *  ========================= */
function applyLoadDelta_(assignment, jenis, delta) {
  const cfg = getConfig_();
  const ss = getSpreadsheet_();
  const shTeam = ss.getSheetByName(cfg.SHEET_TEAM);
  if (!shTeam) throw new Error(`Sheet not found: ${cfg.SHEET_TEAM}`);

  const { headerMap, values } = readTable_(shTeam);
  const colTeam = mustCol_(headerMap, "TEAM_CODE");
  const colWo = mustCol_(headerMap, "ACTIVE_WO_COUNT");
  const colGg = mustCol_(headerMap, "ACTIVE_GANGGUAN_COUNT");
  const colLastAssigned = headerMap["LAST_ASSIGNED_AT"] ?? null;

  const jenisNorm = String(jenis || "").toLowerCase();
  const isGangguan = jenisNorm.includes("gangguan");

  const assigned = (assignment.assignedTo || "-").trim();
  const supports = assignment.supportTeams || [];

  function updateTeam(teamName, isGg) {
    if (!teamName || teamName === "-") return;
    const i = values.findIndex(r => String(r[colTeam] || "").trim() === teamName);
    if (i < 0) return;

    const curWo = Number(values[i][colWo] || 0);
    const curGg = Number(values[i][colGg] || 0);

    values[i][colWo] = Math.max(0, curWo + delta);
    if (isGg) values[i][colGg] = Math.max(0, curGg + delta);

    if (delta > 0 && colLastAssigned !== null) values[i][colLastAssigned] = new Date();
  }

  updateTeam(assigned, isGangguan);
  supports.forEach(t => updateTeam(t, false));

  writeTable_(shTeam, values);
}

/** =========================
 *  10) SHEET + HEADERS
 *  ========================= */
function getSpreadsheet_() {
  const cfg = getConfig_();
  if (cfg.SPREADSHEET_ID) return SpreadsheetApp.openById(cfg.SPREADSHEET_ID);
  return SpreadsheetApp.getActiveSpreadsheet();
}

function readHeader_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h || "").trim());
  const headerMap = {};
  headers.forEach((h, idx) => { if (h) headerMap[h] = idx; });
  return { headers, headerMap };
}

function readTable_(sheet) {
  const { headers, headerMap } = readHeader_(sheet);
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const values = lastRow >= 2 ? sheet.getRange(2, 1, lastRow - 1, lastCol).getValues() : [];
  return { headers, headerMap, values };
}

function writeTable_(sheet, values) {
  const lastCol = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
  if (values.length > 0) sheet.getRange(2, 1, values.length, lastCol).setValues(values);
}

function ensureHeaders_(sheet, headers) {
  const cur = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h || "").trim());
  const set = new Set(cur.filter(Boolean));
  let changed = false;

  headers.forEach(h => {
    if (!set.has(h)) {
      cur.push(h);
      changed = true;
    }
  });

  if (changed) sheet.getRange(1, 1, 1, cur.length).setValues([cur]);
}

function ensureWoDbColumns_(shWo) {
  const required = [
    "WO_ID",
    "WO_STATUS",
    "ASSIGNED_TO",
    "SUPPORT_TEAM",

    "APPROVAL_STATUS",
    "APPROVED_AT",
    "APPROVED_BY",
    "PICK_REQUIRED_TEAMS",
    "PICK_SELECTED_TEAMS",
    "OFFER_STAGE",
    "OFFER_ASSIGNED_MSG_ID",
    "OFFER_SUP1_MSG_ID",
    "OFFER_SUP2_MSG_ID",
    "MANUAL_STAGE",
    "MANUAL_SELECTED_TEAMS",

    "COORD_TG_MESSAGE_ID",
    "COORD_TG_SENT_AT",
    "COORD_TG_RESULT",
    "TEAM_TG_MESSAGE_ID",

    "CalendarEventId",
    "SyncStatus",

    "DONE_SERIALS",
    "DONE_AT",
    "DONE_BY",
    "DONE_PHOTO_COUNT",
    "DONE_PHOTO_FILE_IDS",

    "LAST_UPDATE_AT",
    "LAST_UPDATE_BY"
  ];
  ensureHeaders_(shWo, required);
}

function mustCol_(headerMap, name) {
  if (headerMap[name] === undefined) throw new Error(`Required column missing: ${name}`);
  return headerMap[name];
}

function setCellIfExists_(headerMap, row, name, value) {
  if (headerMap[name] !== undefined) row[headerMap[name]] = value;
}

function getCell_(headerMap, row, name) {
  if (headerMap[name] === undefined) return "";
  return row[headerMap[name]];
}

function getWoContext_(shWo, woId) {
  const { headerMap, values } = readTable_(shWo);
  const woIdCol = headerMap["WO_ID"];
  if (woIdCol === undefined) return null;
  const idx = values.findIndex(r => String(r[woIdCol]).trim() === String(woId).trim());
  if (idx < 0) return null;
  return { headerMap, values, idx, row: values[idx] };
}

/** =========================
 *  11) FORM ROW COPY (if needed)
 *  ========================= */
function appendRowToWoDbFromSource_(srcSheet, srcRow, shWo) {
  const srcHdr = readHeader_(srcSheet).headers;
  const woHdr = readHeader_(shWo).headers;

  const srcVals = srcSheet.getRange(srcRow, 1, 1, srcSheet.getLastColumn()).getValues()[0];
  const newRow = new Array(woHdr.length).fill("");

  const srcMap = {};
  srcHdr.forEach((h, idx) => { if (h) srcMap[String(h).trim()] = idx; });

  woHdr.forEach((h, i) => {
    if (!h) return;
    if (srcMap[h] !== undefined) newRow[i] = srcVals[srcMap[h]];
  });

  const targetRow = shWo.getLastRow() + 1;
  shWo.getRange(targetRow, 1, 1, woHdr.length).setValues([newRow]);
  return targetRow;
}

/** =========================
 *  12) LOGGING (sheet)
 *  ========================= */
function appendLog_(ss, sheetName, row) {
  const sh = ss.getSheetByName(sheetName);
  if (!sh) return;

  ensureHeaders_(sh, ["TIMESTAMP", "WO_ID", "ACTION", "DETAIL", "BY"]);
  const next = sh.getLastRow() + 1;
  sh.getRange(next, 1, 1, row.length).setValues([row]);
}

/** =========================
 *  13) DATA EXTRACTION
 *  ========================= */
function getJenis_(headerMap, row) {
  return getAny_(headerMap, row, ["Jenis Pekerjaan", "JENIS PEKERJAAN", "Jenis", "JENIS"]) || "-";
}

function getAny_(headerMap, row, keys) {
  for (const k of keys) {
    if (headerMap[k] !== undefined) {
      const v = row[headerMap[k]];
      if (v !== null && v !== undefined && String(v).trim() !== "") return v;
    }
  }
  return "";
}

function getMapsUrl_(headerMap, row) {
  const keys = Object.keys(headerMap);

  const direct = ["MAPS_URL", "Maps_URL", "Maps Url", "Maps URL", "MAP_URL", "MAP URL"];
  for (const d of direct) {
    if (headerMap[d] !== undefined) {
      const v = String(row[headerMap[d]] || "").trim();
      if (v) return v;
    }
  }

  const mapKey = keys.find(k => /MAP/i.test(k) && /URL/i.test(k));
  if (mapKey) {
    const v = String(row[headerMap[mapKey]] || "").trim();
    if (v) return v;
  }

  const coordKey = keys.find(k => /KOORD/i.test(k));
  if (coordKey) {
    const coord = String(row[headerMap[coordKey]] || "").trim();
    if (coord) return `https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(coord)}`;
  }

  const lokasi = getAny_(headerMap, row, ["Lokasi", "LOKASI"]);
  if (lokasi) return `https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(String(lokasi))}`;

  return "";
}

function getScheduleText_(headerMap, row) {
  const date = getAny_(headerMap, row, ["Rencana Kerja", "RENCANA KERJA", "Tanggal", "TANGGAL"]);
  const startT = getAny_(headerMap, row, ["Waktu Mulai", "WAKTU MULAI", "Mulai", "MULAI"]);
  const endT = getAny_(headerMap, row, ["Waktu Selesai", "WAKTU SELESAI", "Selesai", "SELESAI"]);

  const d = date ? formatDate_(date) : "-";
  const s = startT ? formatTime_(startT) : "-";
  const e = endT ? formatTime_(endT) : "-";
  return `Rencana: ${d}\nMulai: ${s}\nSelesai: ${e}`;
}

/** FIXED: reads QTY headers like "QTY 1 (set)" correctly */
function buildMaterialsText_(headerMap, row) {
  const mats = [];
  const headers = Object.keys(headerMap);

  for (let i = 1; i <= 10; i++) {
    const mKey = findMaterialHeaderName_(headers, i);
    const qKey = findQtyHeaderName_(headers, i);

    const mVal = mKey ? String(row[headerMap[mKey]] || "").trim() : "";
    const qVal = qKey ? row[headerMap[qKey]] : "";

    if (!mVal) continue;
    const qty = Number(qVal || 0);
    mats.push(`- ${mVal} : ${qty} set`);
  }

  if (mats.length === 0) return "- (kosong)";
  return mats.join("\n");
}

function findMaterialHeaderName_(headers, i) {
  const patterns = [
    new RegExp(`^MATERIAL[\\s_]*${i}\\b`, "i"),
    new RegExp(`^MATERIAL${i}\\b`, "i")
  ];
  for (const h of headers) {
    for (const p of patterns) {
      if (p.test(String(h))) return h;
    }
  }
  return null;
}

function findQtyHeaderName_(headers, i) {
  const patterns = [
    new RegExp(`^QTY[\\s_]*${i}\\b`, "i"),
    new RegExp(`^QTY[\\s_]*${i}\\s*\\(.*\\)`, "i")
  ];
  for (const h of headers) {
    const hs = String(h).trim();
    for (const p of patterns) {
      if (p.test(hs)) return h;
    }
  }
  return null;
}

/** FIXED: sum qty from any header starting with QTY + number, including "(set)" */
function computeTotalSet_(headerMap, row) {
  let total = 0;
  const headers = Object.keys(headerMap);
  for (const h of headers) {
    const up = String(h).toUpperCase().trim();
    if (/^QTY[\s_]*\d+/.test(up)) {
      total += Number(row[headerMap[h]] || 0);
    }
  }
  return total;
}

/** =========================
 *  14) WO ID GENERATOR
 *  ========================= */
function nextWoId_() {
  const cfg = getConfig_();
  const props = PropertiesService.getScriptProperties();
  const prefix = props.getProperty("WO_PREFIX") || cfg.WO_PREFIX || "BL/";
  const pad = Number(props.getProperty("WO_PAD") || cfg.WO_PAD || 3);

  let counter = Number(props.getProperty("WO_COUNTER") || 0);
  counter += 1;
  props.setProperty("WO_COUNTER", String(counter));

  const num = String(counter).padStart(pad, "0");
  return `${prefix}${num}`;
}

/** =========================
 *  15) REQUIRED FIELDS
 *  ========================= */
function validateRequiredFields_(headerMap, row) {
  const missing = [];
  for (const f of REQUIRED_FIELDS) {
    const val = getAny_(headerMap, row, f.keys);
    if (!val || String(val).trim() === "") missing.push(f.label);
  }
  return missing;
}

/** =========================
 *  16) COMMAND PARSING (progress/cancel/reschedule)
 *  ========================= */
function parseCommand_(text) {
  const t = String(text || "").trim();
  if (!t) return null;

  const lower = t.toLowerCase();

  if (lower === "progress" || lower === "progres" || lower === "on progress" || lower === "onprogress") {
    return { type: "STATUS", status: "On Progress" };
  }
  if (lower === "done" || lower === "selesai") {
    return { type: "STATUS", status: "Done" };
  }

  if (lower.startsWith("cancel")) {
    const parts = t.split(":");
    const reason = parts.length >= 2 ? parts.slice(1).join(":").trim() : "-";
    return { type: "CANCEL", reason };
  }

  if (lower.startsWith("reschedule")) {
    const parts = t.split(":");
    if (parts.length < 2) return null;
    const dtStr = parts.slice(1).join(":").trim();
    const dt = parseDdMmYyyyHhMm_(dtStr);
    if (!dt) return null;
    return { type: "RESCHEDULE", datetime: dt };
  }

  return null;
}

function parseSupportTeams_(supportStr) {
  const s = String(supportStr || "").trim();
  if (!s || s === "-") return [];
  return s.split(",").map(x => x.trim()).filter(Boolean);
}

function isTerminalStatus_(s) {
  const x = String(s || "").trim().toLowerCase();
  return x === "done" || x === "cancel" || x === "canceled" || x === "cancelled";
}

function extractWoIdFromText_(text) {
  const t = String(text || "");
  const m = t.match(/WO[_\s-]*ID\s*:\s*([A-Z]{1,10}\/\d+)/i);
  if (m && m[1]) return m[1].trim();

  const m2 = t.match(/\b([A-Z]{1,10}\/\d+)\b/);
  if (m2 && m2[1]) return m2[1].trim();

  return "";
}

/** =========================
 *  17) TIME / FORMAT / PARSE
 *  ========================= */
function parseDdMmYyyyHhMm_(s) {
  const m = String(s || "").trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2}):(\d{2})$/);
  if (!m) return null;
  const dd = Number(m[1]), mm = Number(m[2]) - 1, yyyy = Number(m[3]);
  const hh = Number(m[4]), mi = Number(m[5]);
  const dt = new Date(yyyy, mm, dd, hh, mi, 0);
  if (isNaN(dt.getTime())) return null;
  return dt;
}

function formatDateTime_(d) {
  const dt = coerceDate_(d);
  if (!dt) return "-";
  return Utilities.formatDate(dt, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
}

function formatDate_(d) {
  const dt = coerceDate_(d);
  if (!dt) return "-";
  return Utilities.formatDate(dt, Session.getScriptTimeZone(), "dd/MM/yyyy");
}

function formatTime_(t) {
  const dt = coerceDate_(t);
  if (dt) return Utilities.formatDate(dt, Session.getScriptTimeZone(), "HH:mm");
  return String(t || "-");
}

function coerceDate_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) return v;
  const s = String(v || "").trim();
  if (!s) return null;
  const d = new Date(s);
  if (d instanceof Date && !isNaN(d.getTime())) return d;
  return null;
}

function coerceTime_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) return v;
  const s = String(v || "").trim();
  if (!s) return null;
  const m = s.match(/^(\d{1,2}):(\d{2})$/);
  if (m) {
    const hh = Number(m[1]), mm = Number(m[2]);
    return new Date(1899, 11, 30, hh, mm, 0);
  }
  const d = new Date(s);
  if (!isNaN(d.getTime())) return d;
  return null;
}

function combineDateTime_(date, time) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate(), time.getHours(), time.getMinutes(), 0);
}

function getStartEnd_(headerMap, row, defaultDurMin) {
  const dateVal = getAny_(headerMap, row, ["Rencana Kerja", "RENCANA KERJA", "Tanggal", "TANGGAL"]);
  const startVal = getAny_(headerMap, row, ["Waktu Mulai", "WAKTU MULAI", "Mulai", "MULAI"]);
  const endVal = getAny_(headerMap, row, ["Waktu Selesai", "WAKTU SELESAI", "Selesai", "SELESAI"]);

  const date = coerceDate_(dateVal);
  const startTime = coerceTime_(startVal);
  const endTime = coerceTime_(endVal);

  let start = null;
  let end = null;

  if (date && startTime) start = combineDateTime_(date, startTime);
  else if (date) start = new Date(date.getFullYear(), date.getMonth(), date.getDate(), 8, 0, 0);
  else start = new Date();

  if (date && endTime) end = combineDateTime_(date, endTime);

  if (!end || end.getTime() <= start.getTime()) {
    end = new Date(start.getTime() + (defaultDurMin || 60) * 60000);
  }

  return { start, end };
}

function setScheduleFields_(headerMap, row, dt) {
  const dateKeys = ["Rencana Kerja", "RENCANA KERJA", "Tanggal", "TANGGAL"];
  const startKeys = ["Waktu Mulai", "WAKTU MULAI", "Mulai", "MULAI"];

  for (const k of dateKeys) {
    if (headerMap[k] !== undefined) row[headerMap[k]] = new Date(dt.getFullYear(), dt.getMonth(), dt.getDate());
  }
  for (const k of startKeys) {
    if (headerMap[k] !== undefined) row[headerMap[k]] = new Date(1899, 11, 30, dt.getHours(), dt.getMinutes(), 0);
  }
}

/** =========================
 *  18) CALENDAR (create/update)
 *  ========================= */
function ensureCalendarEvent_(calendarId, woId, row, headerMap, assignment) {
  const cfg = getConfig_();
  if (!calendarId) throw new Error("CALENDAR_ID is empty.");

  const cal = CalendarApp.getCalendarById(calendarId);
  if (!cal) throw new Error("Calendar not found by ID.");

  const { start, end } = getStartEnd_(headerMap, row, cfg.DEFAULT_EVENT_DURATION_MIN);
  const title = buildCalendarTitle_(woId, row, headerMap);
  const desc = buildCalendarDescription_(woId, row, headerMap, assignment);

  let st = start, en = end;
  if (!(st instanceof Date) || isNaN(st.getTime())) throw new Error("Invalid start datetime.");
  if (!(en instanceof Date) || isNaN(en.getTime())) en = new Date(st.getTime() + (cfg.DEFAULT_EVENT_DURATION_MIN || 60) * 60000);
  if (en.getTime() <= st.getTime()) en = new Date(st.getTime() + (cfg.DEFAULT_EVENT_DURATION_MIN || 60) * 60000);

  const ev = cal.createEvent(title, st, en, { description: desc, location: getAny_(headerMap, row, ["Lokasi", "LOKASI"]) || "" });
  return ev.getId();
}

function updateCalendarEvent_(calendarId, eventId, woId, row, headerMap, assignment, opt) {
  const cfg = getConfig_();
  const cal = CalendarApp.getCalendarById(calendarId);
  if (!cal) throw new Error("Calendar not found by ID.");

  const ev = cal.getEventById(eventId);
  if (!ev) throw new Error("Event not found (EventId invalid or deleted).");

  const { start, end } = getStartEnd_(headerMap, row, cfg.DEFAULT_EVENT_DURATION_MIN);
  let st = start, en = end;
  if (!(st instanceof Date) || isNaN(st.getTime())) st = ev.getStartTime();
  if (!(en instanceof Date) || isNaN(en.getTime()) || en.getTime() <= st.getTime()) {
    en = new Date(st.getTime() + (cfg.DEFAULT_EVENT_DURATION_MIN || 60) * 60000);
  }

  ev.setTitle(buildCalendarTitle_(woId, row, headerMap));
  ev.setDescription(buildCalendarDescription_(woId, row, headerMap, assignment, opt));
  ev.setTime(st, en);

  const loc = getAny_(headerMap, row, ["Lokasi", "LOKASI"]) || "";
  if (loc) ev.setLocation(loc);
}

function buildCalendarTitle_(woId, row, headerMap) {
  const status = String(row[headerMap["WO_STATUS"]] || "Planned").trim();
  const jenis = getJenis_(headerMap, row) || "-";
  const pekerjaan = getAny_(headerMap, row, ["NAMA PEKERJAAN", "PEKERJAAN", "Nama Pekerjaan", "Nama PEKERJAAN"]) || "-";
  return `[${woId}] ${jenis} - ${pekerjaan} (${status})`;
}

function buildCalendarDescription_(woId, row, headerMap, assignment, opt) {
  const status = String(row[headerMap["WO_STATUS"]] || "Planned").trim();
  const jenis = getJenis_(headerMap, row) || "-";
  const pekerjaan = getAny_(headerMap, row, ["NAMA PEKERJAAN", "PEKERJAAN", "Nama Pekerjaan", "Nama PEKERJAAN"]) || "-";
  const lokasi = getAny_(headerMap, row, ["Lokasi", "LOKASI"]) || "-";
  const pengawas = getAny_(headerMap, row, ["Nama Pengawas", "PENGAWAS", "Pengawas"]) || "-";
  const penyulang = getAny_(headerMap, row, ["Penyulang", "PENYULANG"]) || "-";
  const up3 = getAny_(headerMap, row, ["UP3"]) || "-";
  const ulp = getAny_(headerMap, row, ["ULP"]) || "-";
  const mapsUrl = getMapsUrl_(headerMap, row);
  const totalSet = computeTotalSet_(headerMap, row);

  const assignedTo = assignment.assignedTo || "-";
  const supports = assignment.supportTeams && assignment.supportTeams.length ? assignment.supportTeams.join(", ") : "-";

  const lines = [];
  lines.push(`WO_ID: ${woId}`);
  lines.push(`Status: ${status}`);
  lines.push(`Jenis: ${jenis}`);
  lines.push(`Pekerjaan: ${pekerjaan}`);
  lines.push(`Penyulang: ${penyulang}`);
  lines.push(`UP3: ${up3}`);
  lines.push(`ULP: ${ulp}`);
  lines.push(`Lokasi: ${lokasi}`);
  lines.push(`Pengawas: ${pengawas}`);
  lines.push(`Maps: ${mapsUrl || "-"}`);
  lines.push("");
  lines.push(`Assigned: ${assignedTo}`);
  lines.push(`Support: ${supports}`);
  lines.push(`Total Set: ${Number(totalSet || 0)}`);
  if (opt && opt.cancelReason) {
    lines.push("");
    lines.push(`Cancel Reason: ${opt.cancelReason}`);
  }
  return lines.join("\n");
}

/** =========================
 *  19) MATERIAL NORMALIZATION (optional)
 *  ========================= */
function syncMaterialsToSheet_(woId, row, headerMap) {
  const cfg = getConfig_();
  const ss = getSpreadsheet_();
  const shMat = ss.getSheetByName(cfg.SHEET_MATERIAL);
  if (!shMat) return;

  ensureHeaders_(shMat, ["WO_ID", "MATERIAL_NAME", "QTY", "UNIT"]);

  const { headerMap: hm, values } = readTable_(shMat);
  const colWo = mustCol_(hm, "WO_ID");
  const colName = mustCol_(hm, "MATERIAL_NAME");
  const colQty = mustCol_(hm, "QTY");
  const colUnit = mustCol_(hm, "UNIT");

  const kept = values.filter(r => String(r[colWo] || "").trim() !== String(woId).trim());

  const headers = Object.keys(headerMap);
  const toAdd = [];
  for (let i = 1; i <= 10; i++) {
    const mKey = findMaterialHeaderName_(headers, i);
    const qKey = findQtyHeaderName_(headers, i);

    const mVal = mKey ? String(row[headerMap[mKey]] || "").trim() : "";
    const qVal = qKey ? Number(row[headerMap[qKey]] || 0) : 0;
    if (!mVal) continue;

    const newRow = new Array(shMat.getLastColumn()).fill("");
    newRow[colWo] = woId;
    newRow[colName] = mVal;
    newRow[colQty] = qVal;
    newRow[colUnit] = "set";
    toAdd.push(newRow);
  }

  writeTable_(shMat, kept.concat(toAdd));
}

/** =========================
 *  20) GANGGUAN PIKET MOVE (opsi 2, same as v4)
 *  ========================= */
function getLastUsedTeamForGangguanRotation_(assignment) {
  const supports = assignment.supportTeams || [];
  if (supports.length) return supports[supports.length - 1];
  const a = (assignment.assignedTo || "").trim();
  return a && a !== "-" ? a : "";
}

function computeNextGangguanTeam_(lastUsedTeamCode) {
  const cfg = getConfig_();
  const ss = getSpreadsheet_();
  const shTeam = ss.getSheetByName(cfg.SHEET_TEAM);
  if (!shTeam) throw new Error(`Sheet not found: ${cfg.SHEET_TEAM}`);

  const { headerMap, values } = readTable_(shTeam);
  const colTeam = mustCol_(headerMap, "TEAM_CODE");
  const colStatus = mustCol_(headerMap, "STATUS");
  const colMode = mustCol_(headerMap, "MODE");
  const colWo = mustCol_(headerMap, "ACTIVE_WO_COUNT");
  const colGg = mustCol_(headerMap, "ACTIVE_GANGGUAN_COUNT");

  const teams = values.map((r, i) => ({
    rowIndex: i,
    team: String(r[colTeam] || "").trim(),
    status: String(r[colStatus] || "").trim().toUpperCase(),
    mode: String(r[colMode] || "").trim().toUpperCase(),
    woCount: Number(r[colWo] || 0),
    ggCount: Number(r[colGg] || 0)
  })).filter(t => t.team);

  const cur = teams.find(t => t.team === String(lastUsedTeamCode || "").trim());
  if (!cur) return "";

  const next = nextAvailableTeamAfter_(teams, cur.rowIndex);
  return next ? next.team : "";
}

function setOnlyOnePiket_(teamCode) {
  const cfg = getConfig_();
  const ss = getSpreadsheet_();
  const shTeam = ss.getSheetByName(cfg.SHEET_TEAM);
  if (!shTeam) throw new Error(`Sheet not found: ${cfg.SHEET_TEAM}`);

  const { headerMap, values } = readTable_(shTeam);
  const colTeam = mustCol_(headerMap, "TEAM_CODE");
  const colMode = mustCol_(headerMap, "MODE");

  const target = String(teamCode || "").trim();
  if (!target) return;

  for (let i = 0; i < values.length; i++) {
    const t = String(values[i][colTeam] || "").trim();
    if (!t) continue;
    values[i][colMode] = (t === target) ? "PIKET" : "NORMAL";
  }

  writeTable_(shTeam, values);
}

/** =========================
 *  21) UTIL JSON
 *  ========================= */
function safeParseJson_(txt) {
  try { return JSON.parse(txt); } catch (e) { return { ok: false, raw: txt }; }
}
function safeJson_(obj) {
  try { return JSON.stringify(obj); } catch (e) { return String(obj); }
}
