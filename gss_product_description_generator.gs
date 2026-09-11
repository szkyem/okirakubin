/**
 * お気楽便 検品結果GSS用 商品文生成スクリプト
 *
 * 使い方:
 * 1. Googleスプレッドシートで「拡張機能」→「Apps Script」を開く
 * 2. このファイルの内容を貼り付ける
 * 3. スクリプトプロパティに OPENAI_API_KEY を設定する
 * 4. setupOkirakubinDescriptionGenerator を1回実行する
 *
 * シート構成:
 * - 検品結果
 *   A: 出力日
 *   B: 商品名
 *   C: ナンバリング
 *   D: 検品日
 *   E: 確認結果
 *   F: 写真URL
 *   G: 生成結果
 *   H: _商品ID（お気楽便側の更新管理用。触らない）
 *   I: 生成ステータス
 *   J: 生成日時
 *
 * - 商品文生成設定
 *   1行目 B列以降: 参考タイトル
 *   2行目 B列以降: 参考文
 *   B3: 出品ルール
 *   B4: 書き方の指示
 *   B5: 禁止事項
 *   B6: 必ず入れる文言
 */

const OKIRAKUBIN_INSPECTION_SHEET_NAME = "検品結果";
const OKIRAKUBIN_PROMPT_SHEET_NAME = "商品文生成設定";
const OKIRAKUBIN_OPENAI_MODEL = "gpt-4.1-mini";

const COL_OUTPUT_DATE = 1;
const COL_PRODUCT_NAME = 2;
const COL_NUMBERING = 3;
const COL_INSPECTION_DATE = 4;
const COL_INSPECTION_RESULT = 5;
const COL_PHOTO_URL = 6;
const COL_GENERATED_TEXT = 7;
const COL_PRODUCT_ID = 8;
const COL_GENERATION_STATUS = 9;
const COL_GENERATED_AT = 10;

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("お気楽便")
    .addItem("未生成の商品文を生成", "generateOkirakubinPendingDescriptions")
    .addItem("自動生成トリガーを設定", "setupOkirakubinDescriptionGenerator")
    .addToUi();
}

function doGet() {
  try {
    const ss = getOkirakubinSpreadsheet_();
    return json_({
      ok: true,
      message: "お気楽便 商品文生成GAS is running.",
      spreadsheetName: ss.getName(),
      hasOpenAiKey: Boolean(PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY"))
    });
  } catch (error) {
    return json_({
      ok: false,
      error: errorMessage_(error)
    });
  }
}

function doPost(e) {
  try {
    const body = parsePostBody_(e);
    if (body.action && body.action !== "generateOkirakubinDescriptions") {
      return json_({ ok: false, error: "unsupported action: " + body.action });
    }

    const expectedToken = PropertiesService.getScriptProperties().getProperty("OKIRAKUBIN_DESCRIPTION_WEBHOOK_TOKEN");
    if (expectedToken && String(body.token || "") !== expectedToken) {
      return json_({ ok: false, error: "token mismatch." });
    }

    const result = generateOkirakubinPendingDescriptions();
    return json_({
      ok: result.ok !== false,
      source: "gss-description-generator",
      generatedCount: result.generatedCount || 0,
      message: result.message || ""
    });
  } catch (error) {
    return json_({
      ok: false,
      error: errorMessage_(error)
    });
  }
}

function setupOkirakubinDescriptionGenerator() {
  ensureOkirakubinDescriptionSheets_();

  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction && trigger.getHandlerFunction() === "generateOkirakubinPendingDescriptions") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger("generateOkirakubinPendingDescriptions")
    .timeBased()
    .everyMinutes(10)
    .create();

  return {
    ok: true,
    message: "お気楽便の商品文生成トリガーを設定しました。"
  };
}

function generateOkirakubinPendingDescriptions() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return { ok: false, error: "別の生成処理が実行中です。" };

  try {
    ensureOkirakubinDescriptionSheets_();

    const ss = getOkirakubinSpreadsheet_();
    const sheet = ss.getSheetByName(OKIRAKUBIN_INSPECTION_SHEET_NAME);
    const lastRow = sheet.getLastRow();

    if (lastRow < 2) {
      return { ok: true, generatedCount: 0, message: "生成対象の行がありません。" };
    }

    const settings = readOkirakubinPromptSettings_(ss);
    const values = sheet.getRange(2, 1, lastRow - 1, COL_GENERATED_AT).getValues();
    let generatedCount = 0;

    values.forEach(function(row, index) {
      const rowNumber = index + 2;
      const productName = String(row[COL_PRODUCT_NAME - 1] || "").trim();
      const numbering = String(row[COL_NUMBERING - 1] || "").trim();
      const inspectionResult = String(row[COL_INSPECTION_RESULT - 1] || "").trim();
      const generatedText = String(row[COL_GENERATED_TEXT - 1] || "").trim();

      if (!productName || !inspectionResult || generatedText) return;

      sheet.getRange(rowNumber, COL_GENERATION_STATUS).setValue("生成中");
      SpreadsheetApp.flush();

      try {
        const description = createOkirakubinProductDescription_({
          productName: productName,
          numbering: numbering,
          inspectionDate: row[COL_INSPECTION_DATE - 1],
          inspectionResult: inspectionResult,
          photoUrl: row[COL_PHOTO_URL - 1],
          settings: settings
        });

        sheet.getRange(rowNumber, COL_GENERATED_TEXT).setValue(description);
        sheet.getRange(rowNumber, COL_GENERATION_STATUS).setValue("生成完了");
        sheet.getRange(rowNumber, COL_GENERATED_AT).setValue(new Date());
        generatedCount += 1;
      } catch (error) {
        sheet.getRange(rowNumber, COL_GENERATION_STATUS).setValue("生成失敗: " + errorMessage_(error).slice(0, 180));
      }

      Utilities.sleep(1200);
    });

    return {
      ok: true,
      generatedCount: generatedCount
    };
  } finally {
    lock.releaseLock();
  }
}

function ensureOkirakubinDescriptionSheets_() {
  const ss = getOkirakubinSpreadsheet_();

  let inspectionSheet = ss.getSheetByName(OKIRAKUBIN_INSPECTION_SHEET_NAME);
  if (!inspectionSheet) inspectionSheet = ss.insertSheet(OKIRAKUBIN_INSPECTION_SHEET_NAME);

  migrateOkirakubinInspectionColumns_(inspectionSheet);

  const headers = [
    "出力日",
    "商品名",
    "ナンバリング",
    "検品日",
    "確認結果",
    "写真URL",
    "生成結果",
    "_商品ID",
    "生成ステータス",
    "生成日時"
  ];
  inspectionSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  inspectionSheet.setFrozenRows(1);

  let settingsSheet = ss.getSheetByName(OKIRAKUBIN_PROMPT_SHEET_NAME);
  if (!settingsSheet) settingsSheet = ss.insertSheet(OKIRAKUBIN_PROMPT_SHEET_NAME);

  const labels = [
    ["参考タイトル"],
    ["参考文"],
    ["出品ルール"],
    ["書き方の指示"],
    ["禁止事項"],
    ["必ず入れる文言"]
  ];
  settingsSheet.getRange(1, 1, labels.length, 1).setValues(labels);
  settingsSheet.setFrozenRows(1);
}

function migrateOkirakubinInspectionColumns_(sheet) {
  if (sheet.getLastRow() < 2 || sheet.getLastColumn() < 1) return;

  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  const previousHeader = sheet.getRange(1, 1, 1, lastColumn).getValues()[0]
    .map(function(value) { return String(value || "").trim(); });
  const previousPhotoColumn = previousHeader.indexOf("写真URL") + 1;
  const previousGeneratedColumn = previousHeader.indexOf("生成結果") + 1;

  if (previousPhotoColumn && previousPhotoColumn !== COL_PHOTO_URL) {
    moveColumnValuesIfTargetBlank_(sheet, previousPhotoColumn, COL_PHOTO_URL, lastRow);
    if (previousPhotoColumn === COL_GENERATED_TEXT) {
      sheet.getRange(2, COL_GENERATED_TEXT, lastRow - 1, 1).clearContent();
    }
  }

  if (previousGeneratedColumn && previousGeneratedColumn !== COL_GENERATED_TEXT) {
    moveColumnValuesIfTargetBlank_(sheet, previousGeneratedColumn, COL_GENERATED_TEXT, lastRow);
  }

  clearUrlLikeGeneratedText_(sheet, lastRow);
}

function moveColumnValuesIfTargetBlank_(sheet, fromColumn, toColumn, lastRow) {
  if (!fromColumn || fromColumn === toColumn || lastRow < 2) return;

  const rowCount = lastRow - 1;
  const fromValues = sheet.getRange(2, fromColumn, rowCount, 1).getValues();
  const toValues = sheet.getRange(2, toColumn, rowCount, 1).getValues();
  let changed = false;

  const merged = toValues.map(function(row, index) {
    const current = String(row[0] || "").trim();
    const fallback = fromValues[index] ? fromValues[index][0] : "";
    if (!current && fallback) changed = true;
    return [current ? row[0] : fallback];
  });

  if (changed) sheet.getRange(2, toColumn, rowCount, 1).setValues(merged);
}

function clearUrlLikeGeneratedText_(sheet, lastRow) {
  if (lastRow < 2) return;

  const rowCount = lastRow - 1;
  const photoValues = sheet.getRange(2, COL_PHOTO_URL, rowCount, 1).getValues();
  const generatedValues = sheet.getRange(2, COL_GENERATED_TEXT, rowCount, 1).getValues();
  let changed = false;

  const cleanedGenerated = generatedValues.map(function(row, index) {
    const generatedText = String(row[0] || "").trim();
    if (!looksLikeUrl_(generatedText)) return [row[0]];

    if (!String(photoValues[index][0] || "").trim()) {
      photoValues[index][0] = row[0];
    }
    changed = true;
    return [""];
  });

  if (changed) {
    sheet.getRange(2, COL_PHOTO_URL, rowCount, 1).setValues(photoValues);
    sheet.getRange(2, COL_GENERATED_TEXT, rowCount, 1).setValues(cleanedGenerated);
  }
}

function looksLikeUrl_(value) {
  return /^https?:\/\//i.test(String(value || "").trim());
}

function readOkirakubinPromptSettings_(ss) {
  const sheet = ss.getSheetByName(OKIRAKUBIN_PROMPT_SHEET_NAME);
  if (!sheet) throw new Error("商品文生成設定シートがありません。");

  const lastColumn = Math.max(sheet.getLastColumn(), 2);
  const titleSamples = sheet.getRange(1, 2, 1, lastColumn - 1).getValues()[0]
    .map(function(value) { return String(value || "").trim(); })
    .filter(Boolean);
  const descriptionSamples = sheet.getRange(2, 2, 1, lastColumn - 1).getValues()[0]
    .map(function(value) { return String(value || "").trim(); })
    .filter(Boolean);

  return {
    titleSamples: titleSamples,
    descriptionSamples: descriptionSamples,
    listingRules: String(sheet.getRange("B3").getValue() || "").trim(),
    writingInstructions: String(sheet.getRange("B4").getValue() || "").trim(),
    prohibited: String(sheet.getRange("B5").getValue() || "").trim(),
    requiredText: String(sheet.getRange("B6").getValue() || "").trim()
  };
}

function createOkirakubinProductDescription_(input) {
  const apiKey = PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");
  if (!apiKey) throw new Error("スクリプトプロパティ OPENAI_API_KEY が未設定です。");

  const prompt = buildOkirakubinPrompt_(input);
  const response = UrlFetchApp.fetch("https://api.openai.com/v1/responses", {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: "Bearer " + apiKey
    },
    payload: JSON.stringify({
      model: PropertiesService.getScriptProperties().getProperty("OPENAI_MODEL") || OKIRAKUBIN_OPENAI_MODEL,
      input: prompt,
      temperature: 0.3
    }),
    muteHttpExceptions: true
  });

  const status = response.getResponseCode();
  const text = response.getContentText() || "";
  const json = JSON.parse(text || "{}");

  if (status < 200 || status >= 300) {
    throw new Error("OpenAI API error status=" + status + " " + text.slice(0, 200));
  }

  const outputText = extractOpenAiText_(json).trim();
  if (!outputText) throw new Error("生成結果が空です。");
  return outputText;
}

function buildOkirakubinPrompt_(input) {
  const settings = input.settings || {};
  return [
    "あなたはカメラ・レンズ等の商品説明文を作成する担当者です。",
    "以下の検品結果に書かれている事実だけを使って、出品用の商品文を作成してください。",
    "推測で状態を良く見せたり、検品結果にない付属品・動作・状態を追加しないでください。",
    "",
    "【商品情報】",
    "商品名: " + (input.productName || "-"),
    "ナンバリング: " + (input.numbering || "-"),
    "検品日: " + formatMaybeDate_(input.inspectionDate),
    "写真URL: " + (input.photoUrl || "-"),
    "",
    "【検品結果】",
    input.inspectionResult || "-",
    "",
    "【参考タイトル】",
    (settings.titleSamples || []).join("\n---\n") || "-",
    "",
    "【参考文】",
    (settings.descriptionSamples || []).join("\n---\n") || "-",
    "",
    "【出品ルール】",
    settings.listingRules || "-",
    "",
    "【書き方の指示】",
    settings.writingInstructions || "-",
    "",
    "【禁止事項】",
    settings.prohibited || "-",
    "",
    "【必ず入れる文言】",
    settings.requiredText || "-",
    "",
    "【出力】",
    "商品文のみを出力してください。"
  ].join("\n");
}

function extractOpenAiText_(json) {
  if (json.output_text) return json.output_text;

  const output = json.output || [];
  const parts = [];
  output.forEach(function(item) {
    (item.content || []).forEach(function(content) {
      if (content.text) parts.push(content.text);
    });
  });
  return parts.join("\n");
}

function formatMaybeDate_(value) {
  if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone() || "Asia/Tokyo", "yyyy-MM-dd");
  }
  return String(value || "-");
}

function errorMessage_(error) {
  return error && error.message ? error.message : String(error);
}

function getOkirakubinSpreadsheet_() {
  const spreadsheetId = PropertiesService.getScriptProperties().getProperty("OKIRAKUBIN_SPREADSHEET_ID");
  if (spreadsheetId) return SpreadsheetApp.openById(spreadsheetId);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error("対象スプレッドシートを取得できません。スクリプトプロパティ OKIRAKUBIN_SPREADSHEET_ID にスプレッドシートIDを設定してください。");
  return ss;
}

function parsePostBody_(e) {
  const raw = e && e.postData && e.postData.contents ? e.postData.contents : "{}";
  try {
    return JSON.parse(raw || "{}");
  } catch (error) {
    throw new Error("POST内容がJSONではありません。");
  }
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
