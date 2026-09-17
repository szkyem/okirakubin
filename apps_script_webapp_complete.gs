const OKIRAKUBIN_SHARED_TOKEN = "okirakubin_drive_folder_20260527_K7mQp9xR2vL8sT4nY6bF";

function doGet(e) {
  try {
    const params = e && e.parameter ? e.parameter : {};
    if (params.action) {
      return routeAction_(params, false);
    }

    return json_({
      ok: true,
      message: "okirakubin Web API is running.",
      version: "2026-08-14-discord-file"
    });
  } catch (error) {
    return json_({
      ok: false,
      error: error && error.message ? error.message : String(error)
    });
  }
}

function testDiscordAuthorization() {
  const response = UrlFetchApp.fetch("https://discord.com", {
    method: "get",
    muteHttpExceptions: true
  });

  Logger.log(response.getResponseCode());
}

function authorizeOkirakubinGas() {
  const response = UrlFetchApp.fetch("https://www.google.com/generate_204", {
    muteHttpExceptions: true
  });

  DriveApp.getRootFolder().getName();
  const tempSpreadsheet = SpreadsheetApp.create("okirakubin_authorization_check");
  DriveApp.getFileById(tempSpreadsheet.getId()).setTrashed(true);
  ScriptApp.getOAuthToken();

  return {
    ok: true,
    status: response.getResponseCode(),
    message: "Authorization check completed."
  };
}

function doPost(e) {
  try {
    const body = parseBody_(e);
    return routeAction_(body, true);
  } catch (error) {
    return json_({
      ok: false,
      error: error && error.message ? error.message : String(error)
    });
  }
}

function routeAction_(body, requireToken) {
  if (requireToken && String(body.token || "") !== OKIRAKUBIN_SHARED_TOKEN) {
    return json_({ ok: false, error: "token mismatch." });
  }

  const action = String(body.action || "");

  if (action === "discordNotifyFile") {
    return handleDiscordNotifyFile_(body);
  }

  if (action === "discordNotify") {
    return handleDiscordNotify_(body);
  }

  if (action === "ocrTrackingScreenshot") {
    return handleOcrTrackingScreenshot_(body);
  }

  if (action === "createDriveFolder") {
    return handleCreateDriveFolder_(body);
  }

  if (action === "exportInspectionResultsToSheet") {
    return handleExportInspectionResultsToSheet_(body);
  }

  return json_({
    ok: false,
    error: "unsupported GAS action: " + (action || "empty")
  });
}

function parseBody_(e) {
  const params = e && e.parameter ? e.parameter : {};
  const raw = e && e.postData && e.postData.contents ? e.postData.contents : "{}";
  try {
    const body = JSON.parse(raw || "{}");
    return Object.assign({}, params, body);
  } catch (error) {
    throw new Error("Failed to parse JSON body.");
  }
}

function handleCreateDriveFolder_(body) {
  const driveBaseUrl = String(body.driveBaseUrl || "").trim();
  const folderName = sanitizeName_(body.folderName || "");

  if (!driveBaseUrl) {
    return json_({ ok: false, error: "Drive base URL is not set." });
  }

  if (!folderName) {
    return json_({ ok: false, error: "Folder name is empty." });
  }

  const rootFolderId = extractFolderId_(driveBaseUrl);
  if (!rootFolderId) {
    return json_({ ok: false, error: "Google Drive URL置き場には、実在するGoogle DriveのフォルダURLを設定してください。" });
  }

  let rootFolder;
  try {
    rootFolder = DriveApp.getFolderById(rootFolderId);
  } catch (error) {
    return json_({
      ok: false,
      error: "Google Drive URL置き場の親フォルダが見つかりません。管理者画面の会員設定で、実在するDriveフォルダURLを設定し、GASを実行しているGoogleアカウントに編集権限を付与してください。"
    });
  }

  const folder = getOrCreateSubFolder_(rootFolder, folderName);

  return json_({
    ok: true,
    folderName: folder.getName(),
    folderUrl: folder.getUrl(),
    folderId: folder.getId()
  });
}

function handleOcrTrackingScreenshot_(body) {
  const imageBase64 = String(body.imageBase64 || "").trim();
  const fileName = sanitizeName_(body.fileName || "yupack_tracking_screenshot.png") || "yupack_tracking_screenshot.png";
  const mimeType = String(body.mimeType || "image/png").trim();

  if (!imageBase64) {
    return json_({ ok: false, error: "OCR image is empty." });
  }

  let tempFileId = "";

  try {
    const bytes = Utilities.base64Decode(imageBase64);
    const blob = Utilities.newBlob(bytes, mimeType, fileName);

    const tempFile = createOcrGoogleDoc_(blob, fileName);

    tempFileId = tempFile.id;
    const doc = DocumentApp.openById(tempFileId);
    const text = doc.getBody().getText();
    const records = parseTrackingRecords_(text);

    return json_({
      ok: true,
      records: records,
      text: text
    });
  } catch (error) {
    return json_({
      ok: false,
      error: "OCR failed. Enable Drive API in Apps Script advanced services. detail: " + (error && error.message ? error.message : String(error))
    });
  } finally {
    if (tempFileId) {
      try {
        DriveApp.getFileById(tempFileId).setTrashed(true);
      } catch (trashError) {
        // Ignore temporary file cleanup failures.
      }
    }
  }
}

function createOcrGoogleDoc_(blob, fileName) {
  const title = "tmp_okirakubin_ocr_" + Date.now() + "_" + fileName;

  if (typeof Drive !== "undefined" && Drive.Files && Drive.Files.insert) {
    return Drive.Files.insert({
      title: title,
      mimeType: MimeType.GOOGLE_DOCS
    }, blob, {
      ocr: true,
      ocrLanguage: "ja"
    });
  }

  return createOcrGoogleDocWithUrlFetch_(blob, title);
}

function createOcrGoogleDocWithUrlFetch_(blob, title) {
  const boundary = "okirakubin_ocr_" + Date.now();
  const metadata = {
    name: title,
    mimeType: "application/vnd.google-apps.document"
  };

  const delimiter = "\r\n--" + boundary + "\r\n";
  const closeDelimiter = "\r\n--" + boundary + "--";

  const payload = Utilities.newBlob(
    delimiter +
    "Content-Type: application/json; charset=UTF-8\r\n\r\n" +
    JSON.stringify(metadata) +
    delimiter +
    "Content-Type: " + blob.getContentType() + "\r\n\r\n"
  ).getBytes()
    .concat(blob.getBytes())
    .concat(Utilities.newBlob(closeDelimiter).getBytes());

  const response = UrlFetchApp.fetch(
    "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&ocrLanguage=ja&fields=id,name",
    {
      method: "post",
      contentType: "multipart/related; boundary=" + boundary,
      headers: {
        Authorization: "Bearer " + ScriptApp.getOAuthToken()
      },
      payload: payload,
      muteHttpExceptions: true
    }
  );

  const status = response.getResponseCode();
  const text = response.getContentText() || "";

  if (status < 200 || status >= 300) {
    throw new Error("Drive API upload failed. status=" + status + " " + text.slice(0, 300));
  }

  const result = JSON.parse(text || "{}");
  if (!result.id) {
    throw new Error("Drive API upload did not return file id. " + text.slice(0, 300));
  }

  return result;
}

function parseTrackingRecords_(text) {
  const lines = String(text || "")
    .split(/\r?\n/)
    .map(function(line) { return normalizeSpace_(line); })
    .filter(Boolean);

  const records = [];

  lines.forEach(function(line, index) {
    const trackingNumbers = extractTrackingNumbers_(line);
    if (!trackingNumbers.length) return;

    const nearby = lines.slice(Math.max(0, index - 3), Math.min(lines.length, index + 4));
    const nearbyText = nearby.join(" ");
    const postalMatch = nearbyText.match(/\b\d{3}[-\s]?\d{4}\b/);
    const postalCode = postalMatch ? normalizePostalCode_(postalMatch[0]) : "";
    const name = inferName_(nearby, line);

    trackingNumbers.forEach(function(trackingNumber) {
      records.push({
        name: name,
        postalCode: postalCode,
        trackingNumber: trackingNumber
      });
    });
  });

  return uniqueRecords_(records);
}

function extractTrackingNumbers_(text) {
  const matches = String(text || "").match(/(?:\d[\s-]?){11,13}/g) || [];
  const values = matches
    .map(function(value) { return String(value || "").replace(/[^\d]/g, ""); })
    .filter(function(value) { return value.length >= 11 && value.length <= 13; });
  return Array.from(new Set(values));
}

function inferName_(nearbyLines, trackingLine) {
  const candidates = nearbyLines
    .filter(function(line) { return !extractTrackingNumbers_(line).length; })
    .filter(function(line) { return !/\b\d{3}[-\s]?\d{4}\b/.test(line); })
    .filter(function(line) { return !/\u767a\u9001|\u8ffd\u8de1|\u304a\u554f\u3044\u5408\u308f\u305b|\u9001\u308a\u72b6|\u756a\u53f7|\u90f5\u4fbf|\u3012|\u3086\u3046\u30d1\u30c3\u30af|\u30d7\u30ea\u30f3\u30c8|R/.test(line); })
    .map(function(line) { return normalizeSpace_(line).replace(/\u69d8$/, ""); })
    .filter(function(line) { return line.length >= 2 && line.length <= 30; });

  return candidates[0] || "";
}

function uniqueRecords_(records) {
  const seen = {};
  return records.filter(function(record) {
    const key = [record.trackingNumber, record.postalCode, record.name].join("_");
    if (seen[key]) return false;
    seen[key] = true;
    return true;
  });
}

function handleDiscordNotifyFile_(body) {
  const requestId = normalizeRequestId_(body.requestId || "");
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const sent = discordRequestStatus_(requestId);
    if (sent) {
      return json_({ ok: true, alreadySent: true, requestId: requestId });
    }

    const webhookUrl = String(body.webhookUrl || "").trim();
    const message = String(body.message || "").trim();
    const fileBase64 = String(body.fileBase64 || "").trim();
    const fileName = sanitizeName_(body.fileName || "label.png") || "label.png";
    const mimeType = String(body.mimeType || "image/png").trim();

    if (!/^https:\/\/(discord\.com|discordapp\.com)\/api\/webhooks\//.test(webhookUrl)) {
      return json_({ ok: false, error: "Invalid Discord Webhook URL." });
    }

    if (!message) {
      return json_({ ok: false, error: "Discord message is empty." });
    }

    if (!fileBase64) {
      return json_({ ok: false, error: "Attachment file is empty." });
    }

    const bytes = Utilities.base64Decode(fileBase64);
    const maxAttempts = 8;
    let lastStatus = 0;
    let lastText = "";
    let lastRetryAfter = 0;

    for (let attempt = 0; attempt < maxAttempts; attempt++) {
      const blob = Utilities.newBlob(bytes, mimeType, fileName);
      const res = UrlFetchApp.fetch(webhookUrl, {
        method: "post",
        payload: {
          payload_json: JSON.stringify({ content: message.slice(0, 2000) }),
          "files[0]": blob
        },
        muteHttpExceptions: true
      });

      lastStatus = res.getResponseCode();
      lastText = res.getContentText() || "";

      if (lastStatus === 200 || lastStatus === 204) {
        rememberDiscordRequest_(requestId);
        return json_({ ok: true, status: lastStatus, requestId: requestId });
      }

      if (lastStatus === 429 && attempt < maxAttempts - 1) {
        const retryAfter = discordRetryAfterSeconds_(res, lastText, attempt);
        lastRetryAfter = retryAfter;
        Utilities.sleep(Math.min(Math.ceil(retryAfter * 1000), 120000));
        continue;
      }

      break;
    }

    return json_({
      ok: false,
      status: lastStatus,
      retryAfter: lastRetryAfter,
      error: "Discord file notification failed. status=" + lastStatus + (lastText ? " " + lastText.slice(0, 240) : "")
    });
  } finally {
    lock.releaseLock();
  }
}

function handleDiscordNotify_(body) {
  const requestId = normalizeRequestId_(body.requestId || "");
  const sent = discordRequestStatus_(requestId);
  if (sent) {
    return json_({ ok: true, alreadySent: true, requestId: requestId });
  }

  const result = sendDiscordWebhook_(body.webhookUrl, body.message);
  if (result && result.ok) {
    rememberDiscordRequest_(requestId);
    result.requestId = requestId;
  }
  return json_(result);
}

function sendDiscordWebhook_(webhookUrlRaw, messageRaw) {
  const webhookUrl = String(webhookUrlRaw || "").trim();
  const message = String(messageRaw || "").trim();

  if (!/^https:\/\/(discord\.com|discordapp\.com)\/api\/webhooks\//.test(webhookUrl)) {
    return { ok: false, error: "Invalid Discord Webhook URL." };
  }

  if (!message) {
    return { ok: false, error: "Discord message is empty." };
  }

  const payload = JSON.stringify({
    content: message.slice(0, 2000)
  });

  const maxAttempts = 8;
  let lastStatus = 0;
  let lastText = "";
  let lastRetryAfter = 0;

  for (let attempt = 0; attempt < maxAttempts; attempt++) {
    const res = UrlFetchApp.fetch(webhookUrl, {
      method: "post",
      contentType: "application/json",
      payload: payload,
      muteHttpExceptions: true
    });

    lastStatus = res.getResponseCode();
    lastText = res.getContentText() || "";

    if (lastStatus === 200 || lastStatus === 204) {
      return { ok: true, status: lastStatus };
    }

    if (lastStatus === 429 && attempt < maxAttempts - 1) {
      const retryAfter = discordRetryAfterSeconds_(res, lastText, attempt);
      lastRetryAfter = retryAfter;
      Utilities.sleep(Math.min(Math.ceil(retryAfter * 1000), 60000));
      continue;
    }

    break;
  }

  return {
    ok: false,
    status: lastStatus,
    retryAfter: lastRetryAfter,
    error: "Discord notification failed. status=" + lastStatus + (lastText ? " " + lastText.slice(0, 240) : "")
  };
}

function discordRetryAfterSeconds_(response, text, attempt) {
  try {
    const headers = response.getAllHeaders ? response.getAllHeaders() : {};
    const headerValue = headers["Retry-After"] || headers["retry-after"];
    const headerSeconds = Number(headerValue);
    if (Number.isFinite(headerSeconds) && headerSeconds > 0) return headerSeconds;
  } catch (error) {
    // ignore header parse errors
  }

  try {
    const parsed = JSON.parse(text || "{}");
    const retryAfter = Number(parsed.retry_after || parsed.retryAfter || 0);
    if (Number.isFinite(retryAfter) && retryAfter > 0) return retryAfter;
  } catch (error) {
    // ignore body parse errors
  }

  return [3, 8, 15, 30, 45, 60][attempt] || 60;
}

function handleExportInspectionResultsToSheet_(body) {
  const spreadsheetUrl = String(body.spreadsheetUrl || "").trim();
  const sheetName = sanitizeName_(body.sheetName || "検品結果") || "検品結果";
  const rows = Array.isArray(body.rows) ? body.rows : [];
  const appendMode = String(body.appendMode || "").trim();
  const descriptionWebhookUrl = String(body.descriptionWebhookUrl || "").trim();

  if (!spreadsheetUrl) {
    return json_({ ok: false, error: "Spreadsheet URL is not set." });
  }

  let spreadsheet;
  try {
    spreadsheet = SpreadsheetApp.openByUrl(spreadsheetUrl);
  } catch (error) {
    const message = error && error.message ? String(error.message) : String(error);
    if (/SpreadsheetApp\.openByUrl|spreadsheets|承認|権限|permission/i.test(message)) {
      return json_({
        ok: false,
        error: "GASのスプレッドシート操作権限が未許可です。Apps Scriptで authorizeOkirakubinGas を実行して許可し、Webアプリを再デプロイしてください。詳細: " + message
      });
    }
    if (/Invalid argument|url/i.test(message)) {
      return json_({
        ok: false,
        error: "GoogleスプレッドシートURLを確認してください。詳細: " + message
      });
    }
    return json_({
      ok: false,
      error: "スプレッドシートを開けませんでした。会員GSSがGAS実行アカウントに共有されているか確認してください。詳細: " + message
    });
  }

  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  }
  const exportedAt = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss");
  const previousHeader = sheet.getLastRow() >= 1
    ? sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0].map(function(value) { return String(value || ""); })
    : [];
  const previousIdColumn = previousHeader.indexOf("_商品ID") + 1;
  const previousGeneratedColumn = previousHeader.indexOf("生成結果") + 1;
  const previousStatusColumn = previousHeader.indexOf("生成ステータス") + 1;
  const previousGeneratedAtColumn = previousHeader.indexOf("生成日時") + 1;
  const header = ["出力日", "商品名", "ナンバリング", "検品日", "確認結果", "写真URL", "生成結果", "_商品ID", "生成ステータス", "生成日時"];
  const generatedColumn = 7;
  const idColumn = 8;
  const statusColumn = 9;
  const generatedAtColumn = 10;
  const makeValue = function(row) {
    return [
      exportedAt,
      row.productName || "",
      row.numbering || "",
      row.inspectionDate || "",
      row.confirmationResult || "",
      row.photoUrl || "",
      "",
      row.productId || [row.numbering || "", row.productName || ""].join("_"),
      "",
      ""
    ];
  };

  if (appendMode === "upsert") {
    if (sheet.getLastRow() < 1) {
      sheet.getRange(1, 1, 1, header.length).setValues([header]);
    } else {
      sheet.getRange(1, 1, 1, header.length).setValues([header]);
    }

    const lastRow = sheet.getLastRow();
    const idColumnForLookup = previousIdColumn || header.length;
    const existingIds = lastRow > 1
      ? sheet.getRange(2, idColumnForLookup, lastRow - 1, 1).getValues().map(function(value) { return String(value[0] || ""); })
      : [];

    rows.forEach(function(row) {
      const value = makeValue(row);
      const productId = String(value[idColumn - 1] || "");
      const existingIndex = existingIds.indexOf(productId);
      if (existingIndex >= 0) {
        const targetRow = existingIndex + 2;
        const previousGeneratedResult = previousGeneratedColumn
          ? sheet.getRange(targetRow, previousGeneratedColumn).getValue()
          : "";
        const keepGeneratedResult = looksLikeUrl_(previousGeneratedResult) ? "" : previousGeneratedResult;
        const keepStatus = keepGeneratedResult && previousStatusColumn
          ? sheet.getRange(targetRow, previousStatusColumn).getValue()
          : "";
        const keepGeneratedAt = keepGeneratedResult && previousGeneratedAtColumn
          ? sheet.getRange(targetRow, previousGeneratedAtColumn).getValue()
          : "";
        sheet.getRange(targetRow, 1, 1, 6).setValues([value.slice(0, 6)]);
        sheet.getRange(targetRow, generatedColumn).setValue(keepGeneratedResult);
        sheet.getRange(targetRow, idColumn).setValue(productId);
        sheet.getRange(targetRow, statusColumn).setValue(keepStatus);
        sheet.getRange(targetRow, generatedAtColumn).setValue(keepGeneratedAt);
      } else {
        sheet.appendRow(value);
        existingIds.push(productId);
      }
    });
  } else {
    const values = [header].concat(rows.map(makeValue));
    sheet.clearContents();
    sheet.getRange(1, 1, values.length, header.length).setValues(values);
  }

  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, header.length).setFontWeight("bold").setBackground("#e9eef6");
  clearSurplusSheetColumns_(sheet, header.length);
  sheet.showColumns(1, header.length);
  sheet.showColumns(generatedColumn);
  sheet.hideColumns(idColumn);

  let descriptionNotificationOk = null;
  let descriptionNotificationError = "";
  if (descriptionWebhookUrl) {
    const notification = notifyDescriptionGenerator_(descriptionWebhookUrl, {
      spreadsheetUrl: spreadsheetUrl,
      sheetName: sheetName,
      memberName: String(body.memberName || ""),
      memberEmail: String(body.memberEmail || ""),
      productIds: rows.map(function(row) {
        return row.productId || [row.numbering || "", row.productName || ""].join("_");
      }).filter(Boolean)
    });
    descriptionNotificationOk = notification.ok === true;
    descriptionNotificationError = notification.error || "";
  }

  return json_({
    ok: true,
    rowCount: rows.length,
    sheetName: sheetName,
    spreadsheetUrl: spreadsheetUrl,
    descriptionNotificationOk: descriptionNotificationOk,
    descriptionNotificationError: descriptionNotificationError
  });
}

function clearSurplusSheetColumns_(sheet, headerLength) {
  const lastColumn = sheet.getLastColumn();
  const lastRow = Math.max(sheet.getLastRow(), 1);
  if (lastColumn <= headerLength) return;
  sheet.getRange(1, headerLength + 1, lastRow, lastColumn - headerLength).clearContent();
}

function looksLikeUrl_(value) {
  return /^https?:\/\//i.test(String(value || "").trim());
}

function notifyDescriptionGenerator_(webhookUrl, payload) {
  if (!/^https:\/\/script\.google\.com\/macros\/s\/[^/]+\/exec/.test(webhookUrl)) {
    return { ok: false, error: "商品文生成通知URLがApps ScriptのWebアプリURLではありません。" };
  }

  try {
    const response = UrlFetchApp.fetch(webhookUrl, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({
        action: "generateOkirakubinDescriptions",
        source: "okirakubin",
        ...payload
      }),
      muteHttpExceptions: true
    });
    const status = response.getResponseCode();
    const text = response.getContentText() || "";
    let parsed = {};

    try {
      parsed = text ? JSON.parse(text) : {};
    } catch (error) {
      return { ok: false, error: "商品文生成GASの返却がJSONではありません。status=" + status + " " + text.slice(0, 180) };
    }

    if (status < 200 || status >= 300 || parsed.ok === false) {
      return { ok: false, error: parsed.error || "商品文生成GASへの通知に失敗しました。status=" + status };
    }

    return {
      ok: true,
      generatedCount: parsed.generatedCount || 0,
      message: parsed.message || ""
    };
  } catch (error) {
    return { ok: false, error: error && error.message ? error.message : String(error) };
  }
}

function getOrCreateSubFolder_(parent, name) {
  const iterator = parent.getFoldersByName(name);
  return iterator.hasNext() ? iterator.next() : parent.createFolder(name);
}

function extractFolderId_(input) {
  const value = String(input || "").trim();
  let match = value.match(/\/folders\/([A-Za-z0-9_-]+)/);
  if (match) return match[1];
  match = value.match(/[?&]id=([A-Za-z0-9_-]+)/);
  if (match) return match[1];
  return /^[A-Za-z0-9_-]{20,}$/.test(value) ? value : "";
}

function sanitizeName_(value) {
  return String(value || "")
    .trim()
    .replace(/[\/\\<>:"|?*\u0000-\u001F]/g, "_")
    .replace(/\s{2,}/g, " ")
    .slice(0, 180);
}

function normalizePostalCode_(value) {
  return String(value || "").replace(/[^\d]/g, "");
}

function normalizeSpace_(value) {
  return String(value || "").replace(/[\s\u3000]+/g, " ").trim();
}

function normalizeRequestId_(value) {
  return String(value || "").trim().replace(/[^A-Za-z0-9_-]/g, "").slice(0, 120);
}

function discordRequestStatus_(requestId) {
  if (!requestId) return false;
  try {
    return CacheService.getScriptCache().get("discord_sent_" + requestId) === "1";
  } catch (error) {
    return false;
  }
}

function rememberDiscordRequest_(requestId) {
  if (!requestId) return;
  try {
    CacheService.getScriptCache().put("discord_sent_" + requestId, "1", 21600);
  } catch (error) {
    // Cache failures should not make a successful Discord send look failed.
  }
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}



/**
 * 17時のDiscord定時連絡トリガーと、送信漏れ確認用のバックアップトリガーを作り直します。
 * Apps Script画面でこの関数を1回だけ実行してください。
 */
function setupOkirakubinDailyDiscordTrigger() {
  const handlers = {
    sendOkirakubinScheduledDiscordNotifications: true,
    watchOkirakubinScheduledDiscordNotifications: true
  };

  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    const handler = trigger.getHandlerFunction ? trigger.getHandlerFunction() : "";
    if (handlers[handler]) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  const hour = Number(getScriptProperty_("OKIRAKUBIN_SCHEDULED_DISCORD_HOUR", "17"));
  const timezone = okirakubinTimezone_();

  ScriptApp.newTrigger("sendOkirakubinScheduledDiscordNotifications")
    .timeBased()
    .everyDays(1)
    .atHour(hour)
    .nearMinute(0)
    .inTimezone(timezone)
    .create();

  ScriptApp.newTrigger("watchOkirakubinScheduledDiscordNotifications")
    .timeBased()
    .everyMinutes(5)
    .inTimezone(timezone)
    .create();

  return {
    ok: true,
    message: "Daily Discord trigger and backup watcher were set.",
    scheduledHour: hour,
    timezone: timezone,
    triggerCount: getOkirakubinScheduledDiscordTriggerStatus().triggerCount
  };
}

/**
 * 時間トリガー用。今日の未送信Discord通知だけを送ります。
 */
function sendOkirakubinScheduledDiscordNotifications() {
  return runOkirakubinScheduledDiscordNotifications_("daily-trigger");
}

/**
 * 送信漏れ確認用。17時以降に未送信が残っていれば送ります。
 */
function watchOkirakubinScheduledDiscordNotifications() {
  const hour = Number(getScriptProperty_("OKIRAKUBIN_SCHEDULED_DISCORD_HOUR", "17"));
  const now = new Date();
  const currentHour = Number(Utilities.formatDate(now, okirakubinTimezone_(), "H"));
  const today = okirakubinTodayString_();

  if (currentHour < hour) {
    return {
      ok: true,
      skipped: true,
      reason: "Before scheduled hour.",
      currentHour: currentHour,
      scheduledHour: hour
    };
  }

  const lastResult = parseStoredJson_(getScriptProperty_("OKIRAKUBIN_LAST_SCHEDULED_DISCORD_RESULT", ""));
  if (lastResult && lastResult.ok === true && lastResult.date === today) {
    return {
      ok: true,
      skipped: true,
      reason: "Scheduled notification already completed today.",
      date: today,
      completedBy: lastResult.source || "scheduled"
    };
  }

  return runOkirakubinScheduledDiscordNotifications_("backup-watcher");
}

function runOkirakubinScheduledDiscordNotifications_(source) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) {
    return {
      ok: true,
      skipped: true,
      reason: "Another scheduled notification process is running.",
      date: okirakubinTodayString_(),
      source: source
    };
  }

  try {
    const result = sendOkirakubinDiscordNotificationsByDate_(okirakubinTodayString_(), "scheduled");
    rememberOkirakubinScheduledDiscordResult_(Object.assign({}, result, {
      source: source
    }));
    return result;
  } catch (error) {
    const result = {
      ok: false,
      source: source,
      date: okirakubinTodayString_(),
      error: error && error.message ? error.message : String(error)
    };
    rememberOkirakubinScheduledDiscordResult_(result);
    throw error;
  } finally {
    lock.releaseLock();
  }
}

function rememberOkirakubinScheduledDiscordResult_(result) {
  const payload = Object.assign({
    ranAt: new Date().toISOString(),
    timezone: okirakubinTimezone_()
  }, result || {});
  PropertiesService.getScriptProperties().setProperty("OKIRAKUBIN_LAST_SCHEDULED_DISCORD_RESULT", JSON.stringify(payload));
}

/**
 * 手動テスト用。Apps Script画面から実行できます。
 */
function testOkirakubinScheduledDiscordNotificationsToday() {
  return sendOkirakubinDiscordNotificationsByDate_(okirakubinTodayString_(), "scheduled");
}

/**
 * トリガー設定と直近実行結果の確認用。
 */
function getOkirakubinScheduledDiscordTriggerStatus() {
  const triggers = ScriptApp.getProjectTriggers()
    .map(function(trigger) {
      return trigger.getHandlerFunction ? trigger.getHandlerFunction() : "";
    })
    .filter(function(handler) {
      return handler === "sendOkirakubinScheduledDiscordNotifications"
        || handler === "watchOkirakubinScheduledDiscordNotifications";
    });

  return {
    ok: true,
    triggerCount: triggers.length,
    handlers: triggers,
    scheduledHour: Number(getScriptProperty_("OKIRAKUBIN_SCHEDULED_DISCORD_HOUR", "17")),
    timezone: okirakubinTimezone_(),
    lastResult: getScriptProperty_("OKIRAKUBIN_LAST_SCHEDULED_DISCORD_RESULT", "")
  };
}

function diagnoseOkirakubinScheduledDiscordNotifications() {
  const date = okirakubinTodayString_();
  const status = getOkirakubinScheduledDiscordTriggerStatus();
  const props = PropertiesService.getScriptProperties();
  const result = {
    ok: true,
    checkedAt: new Date().toISOString(),
    date: date,
    currentHour: Number(Utilities.formatDate(new Date(), okirakubinTimezone_(), "H")),
    timezone: okirakubinTimezone_(),
    triggerStatus: status,
    hasDatabaseSecret: !!props.getProperty("OKIRAKUBIN_DATABASE_SECRET"),
    hasServiceAccountJson: !!props.getProperty("OKIRAKUBIN_SERVICE_ACCOUNT_JSON"),
    hasDatabaseUrl: !!props.getProperty("OKIRAKUBIN_DATABASE_URL"),
    firebaseReadable: false,
    memberCount: 0,
    targetMemberCount: 0,
    unsentItemCount: 0,
    missingWebhookMemberCount: 0,
    errors: []
  };

  try {
    const queue = firebaseGet_("discordNotificationQueue") || {};
    const members = firebaseGet_("members") || {};
    const settings = firebaseGet_("memberSettings") || {};
    result.firebaseReadable = true;
    Object.keys(members).forEach(function(uid) {
      if (members[uid] !== true) return;
      result.memberCount += 1;
      const items = Object.values(queue && queue[uid] && queue[uid][date] ? queue[uid][date] : {})
        .filter(function(item) { return item && item.sent !== true; });
      if (!items.length) return;
      result.targetMemberCount += 1;
      result.unsentItemCount += items.length;
      if (!String((settings[uid] || {}).discordWebhookUrl || "").trim()) {
        result.missingWebhookMemberCount += 1;
      }
    });
  } catch (error) {
    result.ok = false;
    result.errors.push(error && error.message ? error.message : String(error));
  }

  PropertiesService.getScriptProperties().setProperty("OKIRAKUBIN_LAST_SCHEDULED_DISCORD_DIAGNOSIS", JSON.stringify(result));
  return result;
}

function sendOkirakubinDiscordNotificationsByDate_(date, mode) {
  const queue = firebaseGet_("discordNotificationQueue") || {};
  const members = firebaseGet_("members") || {};
  const settings = firebaseGet_("memberSettings") || {};
  const loginUsers = firebaseGet_("loginUsers") || {};
  const updates = {};
  const results = [];
  let sentMemberCount = 0;
  let sentItemCount = 0;
  let skippedItemCount = 0;

  Object.keys(members).forEach(function(uid) {
    if (members[uid] !== true) return;

    const dayMap = queue && queue[uid] && queue[uid][date] ? queue[uid][date] : {};
    const items = Object.keys(dayMap)
      .map(function(id) { return dayMap[id]; })
      .filter(function(item) { return item && item.sent !== true; })
      .sort(function(a, b) { return Number(a.createdAt || 0) - Number(b.createdAt || 0); });

    if (!items.length) return;

    const memberSettings = settings[uid] || {};
    const loginUser = loginUsers[uid] || {};
    const webhookUrl = String(memberSettings.discordWebhookUrl || "").trim();

    if (!webhookUrl) {
      skippedItemCount += items.length;
      results.push({ uid: uid, ok: false, error: "Webhook URL is not set.", count: items.length });
      return;
    }

    const member = {
      uid: uid,
      name: loginUser.name || memberSettings.name || loginUser.email || uid,
      email: loginUser.email || memberSettings.email || "",
      managementNumber: memberSettings.managementNumber || ""
    };

    const message = buildOkirakubinDiscordDailyMessage_(member, date, items, mode);
    const requestId = discordBatchRequestId_(mode, date, uid, items);
    const sent = handleDiscordNotify_({
      webhookUrl: webhookUrl,
      message: message,
      requestId: requestId
    });
    const parsed = parseGasJsonOutput_(sent);

    if (!parsed.ok) {
      skippedItemCount += items.length;
      results.push({ uid: uid, ok: false, error: parsed.error || "Discord send failed.", count: items.length });
      return;
    }

    items.forEach(function(item) {
      const itemId = item.id || "";
      if (!itemId) return;
      updates["discordNotificationQueue/" + uid + "/" + date + "/" + itemId + "/sent"] = true;
      updates["discordNotificationQueue/" + uid + "/" + date + "/" + itemId + "/sentAt"] = Date.now();
      updates["discordNotificationQueue/" + uid + "/" + date + "/" + itemId + "/sentBy"] = "gas-scheduled";
      updates["discordNotificationQueue/" + uid + "/" + date + "/" + itemId + "/sentMode"] = mode;
    });

    sentMemberCount += 1;
    sentItemCount += items.length;
    results.push({ uid: uid, ok: true, count: items.length });
    Utilities.sleep(1500);
  });

  if (Object.keys(updates).length) {
    firebasePatch_("", updates);
  }

  return {
    ok: results.every(function(item) { return item.ok === true; }),
    date: date,
    mode: mode,
    sentMemberCount: sentMemberCount,
    sentItemCount: sentItemCount,
    skippedItemCount: skippedItemCount,
    results: results
  };
}

function discordBatchRequestId_(mode, date, uid, items) {
  const source = [
    mode,
    date,
    uid,
    items.map(function(item) { return String(item.id || ""); }).sort().join("|")
  ].join("_");
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, source, Utilities.Charset.UTF_8)
    .map(function(value) { return (value + 256).toString(16).slice(-2); })
    .join("");
  return "discord_batch_" + digest;
}

function parseStoredJson_(value) {
  try {
    return value ? JSON.parse(value) : null;
  } catch (error) {
    return null;
  }
}

function buildOkirakubinDiscordDailyMessage_(member, date, items, mode) {
  const storedItems = items.filter(function(item) { return item.eventType === "stored"; });
  const shippedItems = items.filter(function(item) { return item.eventType === "shipped"; });
  const title = mode === "scheduled" ? "【お気楽便 本日の定時連絡】" : "【お気楽便 本日の追加連絡】";
  const lines = [
    title,
    mode === "manual" ? "定時連絡後に完了した内容を追加でお知らせします。" : "",
    mode === "manual" ? "" : "",
    "■ 保管商品へ移動：" + storedItems.length + "件",
    "",
    "■ 発送完了：" + shippedItems.length + "件"
  ];

  return lines.filter(function(line, index) { return line || index > 0; }).join("\n").trim();
}

function firebaseGet_(path) {
  const response = firebaseFetch_(path, "get");
  return response && response.ok ? response.value : null;
}

function firebasePatch_(path, value) {
  return firebaseFetch_(path, "patch", value);
}

function firebaseFetch_(path, method, value) {
  const url = firebaseUrl_(path);
  const options = {
    method: method,
    muteHttpExceptions: true
  };

  if (method !== "get") {
    options.contentType = "application/json";
    options.payload = JSON.stringify(value || {});
  }

  const response = UrlFetchApp.fetch(url, options);
  const status = response.getResponseCode();
  const text = response.getContentText() || "";
  let parsed = null;

  try {
    parsed = text ? JSON.parse(text) : null;
  } catch (error) {
    throw new Error("Firebase response is not JSON. status=" + status + " " + text.slice(0, 240));
  }

  if (status < 200 || status >= 300) {
    throw new Error("Firebase request failed. status=" + status + " " + text.slice(0, 240));
  }

  return { ok: true, status: status, value: parsed };
}

function firebaseUrl_(path) {
  const databaseUrl = getScriptProperty_("OKIRAKUBIN_DATABASE_URL", "https://okirakubin-manager-default-rtdb.asia-southeast1.firebasedatabase.app").replace(/\/$/, "");
  const cleanPath = String(path || "").replace(/^\/+|\/+$/g, "");
  const url = databaseUrl + "/" + (cleanPath ? cleanPath.split("/").map(encodeURIComponent).join("/") : "") + ".json";
  const token = firebaseAccessToken_();
  const separator = url.indexOf("?") >= 0 ? "&" : "?";
  return url + separator + token.name + "=" + encodeURIComponent(token.value);
}

function firebaseAccessToken_() {
  const databaseSecret = getScriptProperty_("OKIRAKUBIN_DATABASE_SECRET", "");
  if (databaseSecret) return { name: "auth", value: databaseSecret };

  const serviceAccountJson = getScriptProperty_("OKIRAKUBIN_SERVICE_ACCOUNT_JSON", "");
  if (serviceAccountJson) return { name: "access_token", value: serviceAccountAccessToken_(serviceAccountJson) };

  throw new Error("Set OKIRAKUBIN_DATABASE_SECRET or OKIRAKUBIN_SERVICE_ACCOUNT_JSON in Script Properties.");
}

function serviceAccountAccessToken_(serviceAccountJson) {
  const cache = CacheService.getScriptCache();
  const cached = cache.get("okirakubin_firebase_access_token");
  if (cached) return cached;

  const serviceAccount = JSON.parse(serviceAccountJson);
  const now = Math.floor(Date.now() / 1000);
  const header = { alg: "RS256", typ: "JWT" };
  const claim = {
    iss: serviceAccount.client_email,
    scope: "https://www.googleapis.com/auth/firebase.database https://www.googleapis.com/auth/userinfo.email",
    aud: "https://oauth2.googleapis.com/token",
    iat: now,
    exp: now + 3600
  };
  const unsigned = base64Url_(JSON.stringify(header)) + "." + base64Url_(JSON.stringify(claim));
  const signature = Utilities.computeRsaSha256Signature(unsigned, serviceAccount.private_key);
  const jwt = unsigned + "." + Utilities.base64EncodeWebSafe(signature).replace(/=+$/g, "");

  const response = UrlFetchApp.fetch("https://oauth2.googleapis.com/token", {
    method: "post",
    payload: {
      grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
      assertion: jwt
    },
    muteHttpExceptions: true
  });

  const status = response.getResponseCode();
  const text = response.getContentText() || "";
  const result = JSON.parse(text || "{}");
  if (status < 200 || status >= 300 || !result.access_token) {
    throw new Error("Failed to get service account access token. status=" + status + " " + text.slice(0, 240));
  }

  cache.put("okirakubin_firebase_access_token", result.access_token, 3300);
  return result.access_token;
}

function base64Url_(value) {
  return Utilities.base64EncodeWebSafe(value).replace(/=+$/g, "");
}

function getScriptProperty_(key, fallback) {
  const value = PropertiesService.getScriptProperties().getProperty(key);
  return value == null || value === "" ? fallback : value;
}

function okirakubinTodayString_() {
  return Utilities.formatDate(new Date(), okirakubinTimezone_(), "yyyy-MM-dd");
}

function okirakubinTimezone_() {
  return getScriptProperty_("OKIRAKUBIN_SCHEDULED_DISCORD_TIMEZONE", "Asia/Tokyo");
}

function parseGasJsonOutput_(output) {
  try {
    if (output && typeof output.getContent === "function") {
      return JSON.parse(output.getContent() || "{}");
    }
    if (typeof output === "string") return JSON.parse(output || "{}");
    return output || {};
  } catch (error) {
    return { ok: false, error: "GAS response parse failed." };
  }
}
