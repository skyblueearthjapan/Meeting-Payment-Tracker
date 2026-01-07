/***********************
 * Webアプリ：設定
 ***********************/
const SHEET_MASTER = "名簿マスター";
const SHEET_EVENTS = "イベント管理";
const SHEET_EVENT_TEMPLATE = "会シート";
const EVENT_SHEET_PREFIX = "会_";

const MASTER_HEADERS = ["回答者ID", "氏名", "メールアドレス", "有効フラグ", "備考"];
const EVENTS_HEADERS = ["イベントID", "イベント名", "開催日", "回答締切", "会シート名", "配信用URL", "メール送信ステータス", "作成日時", "備考"];
const EVENT_SHEET_HEADERS = ["回答者ID", "氏名", "メールアドレス", "個別URL", "出欠", "回答日時", "入金状況", "入金確認日時", "要確認フラグ", "備考"];

/***********************
 * WebApp Entry
 ***********************/
function doGet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureBaseSheets_(ss);

  const t = HtmlService.createTemplateFromFile("Index");
  t.appName = "出欠・入金管理（デモWebアプリ）";
  return t.evaluate()
    .setTitle("出欠・入金管理")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include_(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/***********************
 * API：名簿
 ***********************/
function api_getRoster() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureBaseSheets_(ss);
  const sh = ss.getSheetByName(SHEET_MASTER);
  const values = sh.getDataRange().getValues();
  const headers = values[0].map(h => String(h).trim());
  const idx = indexMap_(headers);

  const out = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    if (row.every(v => v === "" || v === null)) continue;
    out.push({
      responderId: String(row[idx["回答者ID"]] || ""),
      name: String(row[idx["氏名"]] || ""),
      email: String(row[idx["メールアドレス"]] || ""),
      active: row[idx["有効フラグ"]],
      note: String(row[idx["備考"]] || ""),
      rowNumber: r + 1
    });
  }
  return { ok: true, headers, rows: out };
}

/**
 * rows: [{responderId,name,email,active,note,rowNumber?}]
 * rowNumberがあれば更新、無ければ新規追加
 */
function api_saveRoster(rows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureBaseSheets_(ss);
  const sh = ss.getSheetByName(SHEET_MASTER);

  const values = sh.getDataRange().getValues();
  const headers = values[0].map(h => String(h).trim());
  const idx = indexMap_(headers);

  // 既存ID重複チェック用
  const existingIds = new Set();
  for (let r = 1; r < values.length; r++) {
    const id = String(values[r][idx["回答者ID"]] || "").trim();
    if (id) existingIds.add(id);
  }

  for (const item of rows) {
    const payload = [
      item.responderId || "",
      item.name || "",
      item.email || "",
      item.active === true || String(item.active).toUpperCase() === "TRUE" ? true : false,
      item.note || ""
    ];

    if (item.rowNumber) {
      // 更新（rowNumberの行を上書き）
      sh.getRange(item.rowNumber, 1, 1, MASTER_HEADERS.length).setValues([payload]);
    } else {
      // 新規追加
      sh.appendRow(payload);
    }
  }

  return { ok: true };
}

function api_assignResponderIdsIfEmpty() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureBaseSheets_(ss);
  const sh = ss.getSheetByName(SHEET_MASTER);

  const values = sh.getDataRange().getValues();
  const headers = values[0].map(h => String(h).trim());
  const idx = indexMap_(headers);
  if (idx["回答者ID"] === undefined) throw new Error("名簿マスターに「回答者ID」列がありません。");

  let counter = 1;
  for (let r = 1; r < values.length; r++) {
    const v = String(values[r][idx["回答者ID"]] || "");
    const m = v.match(/^U(\d+)$/);
    if (m) counter = Math.max(counter, parseInt(m[1], 10) + 1);
  }

  let filled = 0;
  for (let r = 1; r < values.length; r++) {
    const cell = sh.getRange(r + 1, idx["回答者ID"] + 1);
    const cur = String(cell.getValue() || "").trim();
    const name = String(values[r][idx["氏名"]] || "").trim();
    const email = String(values[r][idx["メールアドレス"]] || "").trim();
    if (!name || !email) continue;
    if (cur) continue;

    const newId = "U" + String(counter).padStart(3, "0");
    cell.setValue(newId);
    counter++;
    filled++;
  }

  return { ok: true, filled };
}

/***********************
 * API：イベント
 ***********************/
function api_getEvents() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureBaseSheets_(ss);
  const sh = ss.getSheetByName(SHEET_EVENTS);
  const values = sh.getDataRange().getValues();
  const headers = values[0].map(h => String(h).trim());
  const idx = indexMap_(headers);

  const out = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    if (row.every(v => v === "" || v === null)) continue;
    out.push({
      eventId: String(row[idx["イベントID"]] || ""),
      eventName: String(row[idx["イベント名"]] || ""),
      eventDate: String(row[idx["開催日"]] || ""),
      deadline: String(row[idx["回答締切"]] || ""),
      eventSheetName: String(row[idx["会シート名"]] || ""),
      distributionUrl: String(row[idx["配信用URL"]] || ""),
      mailStatus: String(row[idx["メール送信ステータス"]] || ""),
      createdAt: String(row[idx["作成日時"]] || ""),
      note: String(row[idx["備考"]] || ""),
      rowNumber: r + 1
    });
  }
  // 新しい順っぽく
  out.reverse();
  return { ok: true, rows: out };
}

/**
 * 新規イベント作成（フォーム生成→会シート生成→個別URL生成→イベント管理追加）
 */
function api_createEvent(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureBaseSheets_(ss);

  const eventName = String(payload.eventName || "").trim();
  if (!eventName) throw new Error("イベント名は必須です。");

  const eventId = "EVT" + Utilities.getUuid().slice(0, 6).toUpperCase();
  const eventDate = payload.eventDate ? String(payload.eventDate) : "";
  const deadline = payload.deadline ? String(payload.deadline) : "";

  // フォーム作成
  const form = buildForm_(ss, eventName);

  // 会シート生成
  const eventSheetName = EVENT_SHEET_PREFIX + eventId;
  const eventSheet = createEventSheetFromTemplate_(ss, eventSheetName);

  // イベント管理追加
  const eventRow = appendEventRow_(ss, {
    eventId,
    eventName,
    eventDate,
    deadline,
    eventSheetName,
    distributionUrl: form.getPublishedUrl(),
  });

  // 名簿→会シート + 個別URL
  const members = getActiveMembers_(ss);
  const { eventIdItem, responderIdItem } = getFormInternalItems_(form);
  const linkMap = writeMembersAndLinks_(eventSheet, form, eventId, members, eventIdItem, responderIdItem);

  // フォーム回答反映トリガー
  installOnFormSubmitTrigger_(form);

  // ステータス
  updateEventStatus_(ss, eventRow, "未送信");

  return {
    ok: true,
    event: {
      eventId,
      eventName,
      eventDate,
      deadline,
      eventSheetName,
      distributionUrl: form.getPublishedUrl()
    },
    membersCount: members.length
  };
}

/**
 * イベント指定：出欠メール一斉送信
 */
function api_sendAttendanceMail(eventId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureBaseSheets_(ss);

  const eventInfo = findEventInfo_(ss, eventId);
  if (!eventInfo) throw new Error("イベントが見つかりません。");

  const sh = ss.getSheetByName(eventInfo.eventSheetName);
  if (!sh) throw new Error("会シートが見つかりません。");

  // 会シートから送信（個別URL列を使う）
  const values = sh.getDataRange().getValues();
  const headers = values[0].map(h => String(h).trim());
  const idx = indexMap_(headers);

  const required = ["氏名", "メールアドレス", "個別URL"];
  for (const k of required) if (idx[k] === undefined) throw new Error(`会シートに「${k}」列がありません：${k}`);

  let sent = 0;
  for (let r = 1; r < values.length; r++) {
    const name = String(values[r][idx["氏名"]] || "").trim();
    const email = String(values[r][idx["メールアドレス"]] || "").trim();
    const url = String(values[r][idx["個別URL"]] || "").trim();
    if (!name || !email || !url) continue;

    const subject = `出欠のご回答のお願い（${eventInfo.eventName}）`;
    const body =
`${name} 様

お世話になっております。
「${eventInfo.eventName}」の出欠確認のご案内です。
お手数ですが、下記のリンクよりご回答をお願いいたします。

▼ご回答リンク（あなた専用）
${url}

開催日：${eventInfo.eventDate || "（未設定）"}
回答締切：${eventInfo.deadline || "（未設定）"}

どうぞよろしくお願いいたします。`;

    MailApp.sendEmail({ to: email, subject, body });
    sent++;
  }

  // ステータス更新（イベント管理）
  if (eventInfo.rowNumber) {
    updateEventStatus_(ss, eventInfo.rowNumber, "送信済");
  }

  return { ok: true, sent };
}

/***********************
 * API：会シート（イベント詳細）
 ***********************/
function api_getEventSheet(eventId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureBaseSheets_(ss);

  const eventInfo = findEventInfo_(ss, eventId);
  if (!eventInfo) throw new Error("イベントが見つかりません。");
  const sh = ss.getSheetByName(eventInfo.eventSheetName);
  if (!sh) throw new Error("会シートが見つかりません。");

  const values = sh.getDataRange().getValues();
  const headers = values[0].map(h => String(h).trim());
  const idx = indexMap_(headers);

  const rows = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    if (row.every(v => v === "" || v === null)) continue;

    rows.push({
      responderId: String(row[idx["回答者ID"]] || ""),
      name: String(row[idx["氏名"]] || ""),
      email: String(row[idx["メールアドレス"]] || ""),
      personalUrl: String(row[idx["個別URL"]] || ""),
      attendance: String(row[idx["出欠"]] || ""),
      answeredAt: String(row[idx["回答日時"]] || ""),
      payStatus: String(row[idx["入金状況"]] || ""),
      paidAt: String(row[idx["入金確認日時"]] || ""),
      flag: String(row[idx["要確認フラグ"]] || ""),
      note: String(row[idx["備考"]] || ""),
      rowNumber: r + 1
    });
  }

  return { ok: true, event: eventInfo, headers, rows };
}

/**
 * 入金済にする（複数ID対応）＋入金確認日時を確定入力（トリガーに依存しない）
 */
function api_markPaid(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureBaseSheets_(ss);

  const eventId = String(payload.eventId || "").trim();
  const ids = payload.responderIds || [];
  if (!eventId || !Array.isArray(ids) || ids.length === 0) throw new Error("eventId / responderIds が必要です。");

  const eventInfo = findEventInfo_(ss, eventId);
  if (!eventInfo) throw new Error("イベントが見つかりません。");

  const sh = ss.getSheetByName(eventInfo.eventSheetName);
  const values = sh.getDataRange().getValues();
  const headers = values[0].map(h => String(h).trim());
  const idx = indexMap_(headers);

  const idCol = idx["回答者ID"] + 1;
  const payCol = idx["入金状況"] + 1;
  const paidAtCol = idx["入金確認日時"] + 1;

  const now = formatDateTime_(new Date());
  let updated = 0;

  const idSet = new Set(ids.map(x => String(x)));
  for (let r = 2; r <= sh.getLastRow(); r++) {
    const rid = String(sh.getRange(r, idCol).getValue() || "");
    if (!idSet.has(rid)) continue;

    sh.getRange(r, payCol).setValue("入金済");
    sh.getRange(r, paidAtCol).setValue(now);
    updated++;
  }

  return { ok: true, updated, paidAt: now };
}

/**
 * 未入金メール送信（出席者のみ）
 */
function api_sendUnpaidMail(eventId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureBaseSheets_(ss);

  const eventInfo = findEventInfo_(ss, eventId);
  if (!eventInfo) throw new Error("イベントが見つかりません。");

  const sh = ss.getSheetByName(eventInfo.eventSheetName);
  const values = sh.getDataRange().getValues();
  const headers = values[0].map(h => String(h).trim());
  const idx = indexMap_(headers);

  const required = ["氏名", "メールアドレス", "入金状況", "出欠", "備考"];
  for (const k of required) if (idx[k] === undefined) throw new Error(`会シートに「${k}」列がありません：${k}`);

  const subject = `【ご確認】未入金のご案内（${eventInfo.eventName}）`;
  const today = formatDate_(new Date());
  const now = formatDateTime_(new Date());

  let sent = 0;
  for (let r = 1; r < values.length; r++) {
    const name = String(values[r][idx["氏名"]] || "").trim();
    const email = String(values[r][idx["メールアドレス"]] || "").trim();
    const pay = String(values[r][idx["入金状況"]] || "").trim();
    const att = String(values[r][idx["出欠"]] || "").trim();

    if (!name || !email) continue;
    if (att !== "出席") continue;
    if (!(pay === "" || pay === "未入金")) continue;

    const body =
`${name} 様

お世話になっております。
本日（${today}）時点で、入金状況が「未入金」となっております。

お手数ですが、ご確認のうえお支払い手続きをお願いいたします。
※すでにお支払い済みの場合は、本メールは行き違いとなりますためご容赦ください。

どうぞよろしくお願いいたします。`;

    MailApp.sendEmail({ to: email, subject, body });
    sent++;

    // 備考に記録
    sh.getRange(r + 1, idx["備考"] + 1).setValue(`催促メール送信：${now}`);
  }

  return { ok: true, sent };
}

/***********************
 * 既存ロジック（フォーム生成・トリガー等）
 ***********************/
function ensureBaseSheets_(ss) {
  ensureSheetWithHeaders_(ss, SHEET_MASTER, MASTER_HEADERS);
  ensureSheetWithHeaders_(ss, SHEET_EVENTS, EVENTS_HEADERS);
  ensureSheetWithHeaders_(ss, SHEET_EVENT_TEMPLATE, EVENT_SHEET_HEADERS);
}

function ensureSheetWithHeaders_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  const firstRow = sh.getRange(1, 1, 1, headers.length).getValues()[0];
  const isEmpty = firstRow.every(v => v === "" || v === null);
  if (isEmpty) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
}

function buildForm_(ss, eventName) {
  const form = FormApp.create("出欠確認フォーム（" + eventName + "）");
  form.setDescription("以下の出欠確認にご回答ください。入力は1分ほどで完了します。");

  form.addMultipleChoiceItem()
    .setTitle("出欠を選択してください")
    .setChoiceValues(["出席", "欠席"])
    .setRequired(true);

  form.addTextItem()
    .setTitle("イベントID")
    .setHelpText("※管理用項目です（通常は自動入力されます）")
    .setRequired(false);

  form.addTextItem()
    .setTitle("回答者ID")
    .setHelpText("※管理用項目です（通常は自動入力されます）")
    .setRequired(false);

  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  try { form.setLimitOneResponsePerUser(false); } catch (e) {}
  return form;
}

function getFormInternalItems_(form) {
  const items = form.getItems(FormApp.ItemType.TEXT);
  let eventIdItem = null;
  let responderIdItem = null;
  for (const it of items) {
    const title = it.getTitle();
    if (title === "イベントID") eventIdItem = it.asTextItem();
    if (title === "回答者ID") responderIdItem = it.asTextItem();
  }
  if (!eventIdItem || !responderIdItem) {
    throw new Error("フォームに「イベントID」「回答者ID」のText項目が見つかりません。項目名を確認してください。");
  }
  return { eventIdItem, responderIdItem };
}

function createEventSheetFromTemplate_(ss, newName) {
  let name = newName;
  let i = 2;
  while (ss.getSheetByName(name)) {
    name = newName + "_" + i;
    i++;
  }
  const template = ss.getSheetByName(SHEET_EVENT_TEMPLATE);
  const copied = template.copyTo(ss).setName(name);
  copied.getRange(2, 1, copied.getMaxRows() - 1, copied.getMaxColumns()).clearContent();
  return copied;
}

function appendEventRow_(ss, { eventId, eventName, eventDate, deadline, eventSheetName, distributionUrl }) {
  const sh = ss.getSheetByName(SHEET_EVENTS);
  const row = [
    eventId,
    eventName,
    eventDate,
    deadline,
    eventSheetName,
    distributionUrl,
    "未送信",
    formatDateTime_(new Date()),
    ""
  ];
  sh.appendRow(row);
  return sh.getLastRow();
}

function updateEventStatus_(ss, eventRow, status) {
  const sh = ss.getSheetByName(SHEET_EVENTS);
  sh.getRange(eventRow, 7).setValue(status);
}

function getActiveMembers_(ss) {
  const sh = ss.getSheetByName(SHEET_MASTER);
  const values = sh.getDataRange().getValues();
  const headers = values[0].map(h => String(h).trim());
  const idx = indexMap_(headers);

  const out = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const id = row[idx["回答者ID"]];
    const name = row[idx["氏名"]];
    const mail = row[idx["メールアドレス"]];
    const active = row[idx["有効フラグ"]];

    if (!id || !name || !mail) continue;
    const isActive = (active === true) || (String(active).toUpperCase() === "TRUE") || (String(active) === "有効");
    if (!isActive) continue;

    out.push({ responderId: String(id), name: String(name), email: String(mail) });
  }
  if (out.length === 0) {
    throw new Error("名簿マスターに有効なデータがありません。回答者ID/氏名/メール/有効フラグ を確認してください。");
  }
  return out;
}

function writeMembersAndLinks_(eventSheet, form, eventId, members, eventIdItem, responderIdItem) {
  const startRow = 2;
  const rows = [];
  const linkMap = {};

  for (const m of members) {
    const resp = form.createResponse()
      .withItemResponse(eventIdItem.createResponse(String(eventId)))
      .withItemResponse(responderIdItem.createResponse(String(m.responderId)));

    const url = resp.toPrefilledUrl();
    linkMap[m.responderId] = url;

    rows.push([
      m.responderId,
      m.name,
      m.email,
      url,
      "未回答",
      "",
      "未入金",
      "",
      "",
      ""
    ]);
  }

  eventSheet.getRange(startRow, 1, rows.length, EVENT_SHEET_HEADERS.length).setValues(rows);
  return linkMap;
}

function installOnFormSubmitTrigger_(form) {
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) {
    if (t.getHandlerFunction() === "onDemoFormSubmit_") {
      ScriptApp.deleteTrigger(t);
    }
  }
  ScriptApp.newTrigger("onDemoFormSubmit_")
    .forForm(form)
    .onFormSubmit()
    .create();
}

function onDemoFormSubmit_(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const itemResponses = e.response.getItemResponses();
  const byTitle = {};
  for (const ir of itemResponses) {
    byTitle[ir.getItem().getTitle()] = ir.getResponse();
  }

  const attendance = byTitle["出欠を選択してください"];
  const eventId = byTitle["イベントID"];
  const responderId = byTitle["回答者ID"];
  const ts = e.response.getTimestamp();

  if (!eventId || !responderId) return;

  const eventInfo = findEventInfo_(ss, eventId);
  if (!eventInfo) return;

  const sh = ss.getSheetByName(eventInfo.eventSheetName);
  if (!sh) return;

  const values = sh.getDataRange().getValues();
  const headers = values[0].map(h => String(h).trim());
  const idx = indexMap_(headers);

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    if (String(row[idx["回答者ID"]]) === String(responderId)) {
      sh.getRange(r + 1, idx["出欠"] + 1).setValue(attendance || "");
      sh.getRange(r + 1, idx["回答日時"] + 1).setValue(formatDateTime_(ts));
      sh.getRange(r + 1, idx["要確認フラグ"] + 1).setValue("");
      return;
    }
  }
}

function findEventInfo_(ss, eventId) {
  const sh = ss.getSheetByName(SHEET_EVENTS);
  const values = sh.getDataRange().getValues();
  const headers = values[0].map(h => String(h).trim());
  const idx = indexMap_(headers);

  for (let r = 1; r < values.length; r++) {
    if (String(values[r][idx["イベントID"]]) === String(eventId)) {
      return {
        eventId: String(values[r][idx["イベントID"]] || ""),
        eventName: String(values[r][idx["イベント名"]] || ""),
        eventDate: String(values[r][idx["開催日"]] || ""),
        deadline: String(values[r][idx["回答締切"]] || ""),
        eventSheetName: String(values[r][idx["会シート名"]] || ""),
        distributionUrl: String(values[r][idx["配信用URL"]] || ""),
        mailStatus: String(values[r][idx["メール送信ステータス"]] || ""),
        rowNumber: r + 1
      };
    }
  }
  return null;
}

/***********************
 * Utility
 ***********************/
function indexMap_(headers) {
  const map = {};
  headers.forEach((h, i) => {
    if (!h) return;
    map[String(h).trim()] = i;
  });
  return map;
}

function formatDate_(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy/MM/dd");
}

function formatDateTime_(d) {
  return Utilities.formatDate(new Date(d), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm");
}
