/***************************************
 * WO Pilot Automation - v5 (Split Files)
 * File: WO_Pilot_Telegram.gs
 ***************************************/

/** =========================
 *  TELEGRAM API (low-level)
 *  ========================= */
function telegramApiRequest_(token, method, payload, httpMethod) {
  if (!token) throw new Error("TELEGRAM_BOT_TOKEN is empty.");
  const url = `https://api.telegram.org/bot${token}/${method}`;

  const opt = {
    method: (httpMethod || "post").toLowerCase(),
    muteHttpExceptions: true
  };

  if (opt.method === "post") {
    opt.contentType = "application/json";
    opt.payload = JSON.stringify(payload || {});
  }

  const res = UrlFetchApp.fetch(url, opt);
  return safeParseJson_(res.getContentText());
}

function telegramSendMessage_(token, chatId, text, opt) {
  if (!chatId) throw new Error("TELEGRAM chatId is empty.");

  const payload = {
    chat_id: String(chatId),
    text: String(text || ""),
    disable_web_page_preview: opt && opt.disablePreview === true ? true : false
  };

  if (opt && opt.parseMode) payload.parse_mode = opt.parseMode;
  if (opt && opt.replyToMessageId) payload.reply_to_message_id = Number(opt.replyToMessageId);
  if (opt && opt.replyMarkup) payload.reply_markup = opt.replyMarkup;

  return telegramApiRequest_(token, "sendMessage", payload, "post");
}

function telegramEditMessageText_(token, chatId, messageId, text, opt) {
  const payload = {
    chat_id: String(chatId),
    message_id: Number(messageId),
    text: String(text || "")
  };
  if (opt && opt.parseMode) payload.parse_mode = opt.parseMode;
  if (opt && opt.replyMarkup !== undefined) payload.reply_markup = opt.replyMarkup; // can be null to remove
  if (opt && opt.disablePreview === true) payload.disable_web_page_preview = true;

  return telegramApiRequest_(token, "editMessageText", payload, "post");
}

function telegramAnswerCallbackQuery_(token, callbackQueryId, text, showAlert) {
  const payload = {
    callback_query_id: String(callbackQueryId),
    text: String(text || ""),
    show_alert: showAlert ? true : false
  };
  return telegramApiRequest_(token, "answerCallbackQuery", payload, "post");
}

function telegramGetUpdates_(token, offset) {
  const url = `https://api.telegram.org/bot${token}/getUpdates?offset=${encodeURIComponent(String(offset || 0))}`;
  const res = UrlFetchApp.fetch(url, { method: "get", muteHttpExceptions: true });
  return safeParseJson_(res.getContentText());
}

function telegramGetMe_(token) {
  return telegramApiRequest_(token, "getMe", {}, "post");
}

/** =========================
 *  Media group (album)
 *  ========================= */
function telegramSendMediaGroupChunked_(token, chatId, fileIds, opt) {
  const ids = (fileIds || []).filter(Boolean);
  if (ids.length === 0) return { ok: false, error: "empty fileIds" };

  const chunks = [];
  for (let i = 0; i < ids.length; i += 10) chunks.push(ids.slice(i, i + 10));

  const results = [];
  for (let c = 0; c < chunks.length; c++) {
    const chunk = chunks[c];
    const media = chunk.map((fid, idx) => {
      const item = { type: "photo", media: fid };
      // optional caption only on first photo of first chunk (if you want)
      if (opt && opt.caption && c === 0 && idx === 0) {
        item.caption = String(opt.caption);
        if (opt.parseMode) item.parse_mode = opt.parseMode;
      }
      return item;
    });

    const payload = { chat_id: String(chatId), media: JSON.stringify(media) };
    const r = telegramApiRequest_(token, "sendMediaGroup", payload, "post");
    results.push(r);
  }
  return { ok: true, results };
}

/** =========================
 *  Inline keyboard builder
 *  ========================= */
function tgInlineKeyboard_(rows) {
  return { inline_keyboard: rows };
}
