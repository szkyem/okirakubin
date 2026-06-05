/**
 * Discord通知用のApps Script側パッチ。
 * 既存のdoPostで action === "discordNotify" の時に、この関数を呼び出してください。
 */
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
      payload,
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
  } catch (e) {
    // ignore header parse errors
  }

  try {
    const parsed = JSON.parse(text || "{}");
    const retryAfter = Number(parsed.retry_after || parsed.retryAfter || 0);
    if (Number.isFinite(retryAfter) && retryAfter > 0) return retryAfter;
  } catch (e) {
    // ignore body parse errors
  }

  return [2.5, 6, 12, 20][attempt] || 20;
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
