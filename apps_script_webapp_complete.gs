const GAS_SHARED_TOKEN = "okirakubin_drive_folder_20260527_K7mQp9xR2vL8sT4nY6bF";

function doGet() {
  return json_({
    ok: true,
    message: "お気楽便 Web API is running."
  });
}

function doPost(e) {
  try {
    const body = parseBody_(e);

    if (String(body.token || "") !== GAS_SHARED_TOKEN) {
      return json_({ ok: false, error: "tokenが一致しません。" });
    }

    const action = String(body.action || "createDriveFolder");

    if (action === "discordNotify") {
      return handleDiscordNotify_(body);
    }

    if (action === "ocrTrackingScreenshot") {
      return handleOcrTrackingScreenshot_(body);
    }

    return handleCreateDriveFolder_(body);
  } catch (error) {
    return json_({
      ok: false,
      error: error && error.message ? error.message : String(error)
    });
  }
}

function parseBody_(e) {
  const raw = e && e.postData && e.postData.contents ? e.postData.contents : "{}";
  try {
    return JSON.parse(raw);
  } catch (error) {
    throw new Error("JSONの解析に失敗しました。");
  }
}

function handleCreateDriveFolder_(body) {
  const driveBaseUrl = String(body.driveBaseUrl || "").trim();
  const folderName = sanitizeName_(body.folderName || "");

  if (!driveBaseUrl) {
    return json_({ ok: false, error: "Drive URL置き場が未設定です。" });
  }

  if (!folderName) {
    return json_({ ok: false, error: "フォルダ名が空です。" });
  }

  const rootFolderId = extractFolderId_(driveBaseUrl);
  if (!rootFolderId) {
    return json_({ ok: false, error: "DriveフォルダIDを取得できませんでした。" });
  }

  const rootFolder = DriveApp.getFolderById(rootFolderId);
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
    return json_({ ok: false, error: "OCR対象の画像がありません。" });
  }

  let tempFileId = "";

  try {
    const bytes = Utilities.base64Decode(imageBase64);
    const blob = Utilities.newBlob(bytes, mimeType, fileName);

    const resource = {
      title: "tmp_okirakubin_ocr_" + Date.now() + "_" + fileName,
      mimeType: MimeType.GOOGLE_DOCS
    };

    const tempFile = Drive.Files.insert(resource, blob, {
      ocr: true,
      ocrLanguage: "ja"
    });

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
      error: "OCR処理に失敗しました。Apps Scriptの高度なGoogleサービスでDrive APIを有効化してください。詳細: " + (error && error.message ? error.message : String(error))
    });
  } finally {
    if (tempFileId) {
      try {
        DriveApp.getFileById(tempFileId).setTrashed(true);
      } catch (trashError) {
        // 一時ファイルの削除失敗はOCR結果返却を妨げない。
      }
    }
  }
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
    .filter(function(line) { return !/発送|追跡|お問い合わせ|送り状|番号|郵便|〒|ゆうパック|プリント|R/.test(line); })
    .map(function(line) { return normalizeSpace_(line).replace(/様$/, ""); })
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

function handleDiscordNotify_(body) {
  const webhookUrl = String(body.webhookUrl || "").trim();
  const message = String(body.message || "").trim();

  if (!/^https:\/\/(discord\.com|discordapp\.com)\/api\/webhooks\//.test(webhookUrl)) {
    return json_({ ok: false, error: "Discord Webhook URLの形式が正しくありません。" });
  }

  if (!message) {
    return json_({ ok: false, error: "Discord通知メッセージが空です。" });
  }

  const payload = JSON.stringify({
    content: message.slice(0, 2000)
  });

  const maxAttempts = 5;
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
      return json_({ ok: true });
    }

    if (lastStatus === 429 && attempt < maxAttempts - 1) {
      const retryAfter = discordRetryAfterSeconds_(res, lastText, attempt);
      lastRetryAfter = retryAfter;
      Utilities.sleep(Math.min(Math.ceil(retryAfter * 1000), 30000));
      continue;
    }

    break;
  }

  return json_({
    ok: false,
    status: lastStatus,
    retryAfter: lastRetryAfter,
    error: "Discord notification failed. status=" + lastStatus + (lastText ? " " + lastText.slice(0, 240) : "")
  });
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

  return [2.5, 6, 12, 20][attempt] || 20;
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
  return value;
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

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
