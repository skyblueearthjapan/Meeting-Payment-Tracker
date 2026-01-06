/***********************
 * デモ設定
 ***********************/
const SHEET_MASTER = "名簿マスター";
const SHEET_EVENTS = "イベント管理";
const SHEET_EVENT_TEMPLATE = "会シート"; // テンプレ（見出しのみ）
const EVENT_SHEET_PREFIX = "会_";        // 自動生成する会シート名の接頭辞（例：会_EVT001）

// 列ヘッダ（あなたの見出しに合わせて固定）
const MASTER_HEADERS = ["回答者ID", "氏名", "メールアドレス", "有効フラグ", "備考"];
const EVENTS_HEADERS = ["イベントID", "イベント名", "開催日", "回答締切", "会シート名", "配信用URL", "メール送信ステータス", "作成日時", "備考"];
const EVENT_SHEET_HEADERS = ["回答者ID", "氏名", "メールアドレス", "個別URL", "出欠", "回答日時", "入金状況", "入金確認日時", "要確認フラグ", "備考"];

/***********************
 * 入口：デモ実行（これだけ押せばOK）
 * 1) フォーム自動生成
 * 2) 会シート自動生成（テンプレコピー）
 * 3) 個別URL自動生成
 * 4) メール自動送信（本文に個別URL差し込み）
 * 5) 回答反映トリガー作成（任意だがデモに強い）
 ***********************/
function demo_createEvent_and_sendMails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureBaseSheets_(ss);

  // 1) イベント情報（デモ用に固定。必要ならここだけ書き換え）
  const eventId = "EVT" + Utilities.getUuid().slice(0, 6).toUpperCase(); // 例: EVT1A2B3C
  const eventName = "デモ説明会";
  const eventDate = formatDate_(new Date(Date.now() + 7 * 24 * 60 * 60 * 1000));   // 1週間後
  const deadline = formatDate_(new Date(Date.now() + 5 * 24 * 60 * 60 * 1000));    // 5日後

  // 2) フォーム作成
  const form = buildForm_(ss, eventName);

  // 3) 会シート生成（テンプレコピー）
  const eventSheetName = EVENT_SHEET_PREFIX + eventId;
  const eventSheet = createEventSheetFromTemplate_(ss, eventSheetName);

  // 4) イベント管理に1行追加
  const eventRow = appendEventRow_(ss, {
    eventId,
    eventName,
    eventDate,
    deadline,
    eventSheetName,
    distributionUrl: form.getPublishedUrl(), // 参考URL（共通）
  });

  // 5) 名簿から有効メンバー取得 → 会シートへ転記 & 個別URL生成
  const members = getActiveMembers_(ss);
  const { eventIdItem, responderIdItem } = getFormInternalItems_(form);
  const personalLinks = writeMembersAndLinks_(eventSheet, form, eventId, members, eventIdItem, responderIdItem);

  // 6) メール送信（本文に個別URL差し込み）
  sendPersonalLinksByEmail_(members, personalLinks, {
    eventName,
    eventDate,
    deadline,
  });

  // 7) フォーム回答→会シート反映トリガー（デモに強い）
  installOnFormSubmitTrigger_(form);

  // ステータス更新
  updateEventStatus_(ss, eventRow, "送信済");

  SpreadsheetApp.getUi().alert(
    "デモ準備完了！\n\n" +
    "イベントID: " + eventId + "\n" +
    "会シート: " + eventSheetName + "\n\n" +
    "次は、送られた個別URLを開いて「出席/欠席」を送信してみてください。\n" +
    "送信後、会シートの「出欠」「回答日時」が自動で埋まります。"
  );
}

/***********************
 * フォーム作成
 ***********************/
function buildForm_(ss, eventName) {
  const form = FormApp.create("出欠確認フォーム（" + eventName + "）");

  form.setDescription("以下の出欠確認にご回答ください。入力は1分ほどで完了します。");

  // Q1: 出欠（必須）
  form.addMultipleChoiceItem()
    .setTitle("出欠を選択してください")
    .setChoiceValues(["出席", "欠席"])
    .setRequired(true);

  // Q2: イベントID（内部管理用・短文）
  form.addTextItem()
    .setTitle("イベントID")
    .setHelpText("※管理用項目です（通常は自動入力されます）")
    .setRequired(false);

  // Q3: 回答者ID（内部管理用・短文）
  form.addTextItem()
    .setTitle("回答者ID")
    .setHelpText("※管理用項目です（通常は自動入力されます）")
    .setRequired(false);

  // 回答先スプレッドシートに紐付け（自動で「フォームの回答 1」等が作成される）
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

  // 1人1回制限（ログイン不要運用なのでOFF推奨）
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

/***********************
 * シート整備
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

function createEventSheetFromTemplate_(ss, newName) {
  // 同名があれば削除せず末尾に連番
  let name = newName;
  let i = 2;
  while (ss.getSheetByName(name)) {
    name = newName + "_" + i;
    i++;
  }
  const template = ss.getSheetByName(SHEET_EVENT_TEMPLATE);
  const copied = template.copyTo(ss).setName(name);
  copied.getRange(2, 1, copied.getMaxRows() - 1, copied.getMaxColumns()).clearContent();

  // ★追加：擬似ボタンエリアを作る
  setupActionArea_(copied);

  return copied;
}

/***********************
 * 会シートに擬似ボタンエリア（チェックボックス）を設置
 ***********************/
function setupActionArea_(sheet) {
  // タイトル
  sheet.getRange("L1").setValue("デモ操作（チェックで実行）");
  sheet.getRange("L2").setValue("☑ 未入金メール送信（出席者のみ）");
  sheet.getRange("L3").setValue("☑ 入金済にした行へ入金日時を反映");
  sheet.getRange("L4").setValue("☑ 入金トリガー設定（B）");

  // チェックボックス
  sheet.getRange("M2:M4").insertCheckboxes().setValue(false);

  // 見た目（任意）
  sheet.setColumnWidth(12, 220); // L
  sheet.setColumnWidth(13, 30);  // M
}

/***********************
 * 名簿マスターに擬似ボタンエリア（チェックボックス）を設置
 ***********************/
function setupMasterActionArea_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_MASTER);
  if (!sh) throw new Error("名簿マスターが見つかりません。");

  // 表示文言
  sh.getRange("L1").setValue("名簿マスター：デモ操作（チェックで実行）");
  sh.getRange("L2").setValue("☑ 回答者IDを自動発行（空欄のみ）");
  sh.getRange("L3").setValue("☑ 有効フラグを一括TRUE（空欄のみ）");
  sh.getRange("L4").setValue("☑ テストメール送信（有効な全員）");

  // チェックボックス
  sh.getRange("M2:M4").insertCheckboxes().setValue(false);

  // 見た目
  sh.setColumnWidth(12, 260); // L
  sh.setColumnWidth(13, 30);  // M

  SpreadsheetApp.getUi().alert("名簿マスターに操作エリアを設置しました。");
}

/***********************
 * 名簿読み込み（有効フラグ=TRUE/有効 を対象）
 ***********************/
function getActiveMembers_(ss) {
  const sh = ss.getSheetByName(SHEET_MASTER);
  const values = sh.getDataRange().getValues();
  const headers = values[0];
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

/***********************
 * イベント管理 行追加
 ***********************/
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
    "デモ用"
  ];
  sh.appendRow(row);
  return sh.getLastRow(); // 行番号
}

function updateEventStatus_(ss, eventRow, status) {
  const sh = ss.getSheetByName(SHEET_EVENTS);
  // G列：メール送信ステータス
  sh.getRange(eventRow, 7).setValue(status);
}

/***********************
 * 会シートに名簿コピー + 個別URL生成
 ***********************/
function writeMembersAndLinks_(eventSheet, form, eventId, members, eventIdItem, responderIdItem) {
  const startRow = 2;
  const rows = [];
  const linkMap = {}; // responderId -> url

  for (const m of members) {
    // eventId / responderId を事前入力した個別URLを生成
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

/***********************
 * メール送信（本文に個別URL差し込み）
 ***********************/
function sendPersonalLinksByEmail_(members, linkMap, { eventName, eventDate, deadline }) {
  for (const m of members) {
    const url = linkMap[m.responderId];
    const subject = `出欠のご回答のお願い（${eventName}）`;
    const body =
`${m.name} 様

お世話になっております。
「${eventName}」の出欠確認のご案内です。
お手数ですが、下記のリンクよりご回答をお願いいたします。

▼ご回答リンク（あなた専用）
${url}

開催日：${eventDate}
回答締切：${deadline}

どうぞよろしくお願いいたします。`;

    // デモなので MailApp でOK（GmailAppでも可）
    MailApp.sendEmail({
      to: m.email,
      subject,
      body
    });
  }
}

/***********************
 * フォーム送信 → 会シートへ自動反映（デモに最強）
 ***********************/
function installOnFormSubmitTrigger_(form) {
  // 既存の同関数トリガーが増えすぎないよう、同名は消す
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

// フォーム送信時に呼ばれる
function onDemoFormSubmit_(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 回答から値取得
  const itemResponses = e.response.getItemResponses();
  const byTitle = {};
  for (const ir of itemResponses) {
    byTitle[ir.getItem().getTitle()] = ir.getResponse();
  }

  const attendance = byTitle["出欠を選択してください"];
  const eventId = byTitle["イベントID"];
  const responderId = byTitle["回答者ID"];
  const ts = e.response.getTimestamp();

  if (!eventId || !responderId) {
    // 管理用IDがない＝要確認（デモではログだけ）
    return;
  }

  // イベント管理から会シート名を探す
  const eventSheetName = findEventSheetName_(ss, eventId);
  if (!eventSheetName) return;

  const sh = ss.getSheetByName(eventSheetName);
  if (!sh) return;

  // 会シートで該当回答者ID行を探して更新
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  const idx = indexMap_(headers);

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    if (String(row[idx["回答者ID"]]) === String(responderId)) {
      // E列：出欠、F列：回答日時、I列：要確認フラグ
      sh.getRange(r + 1, idx["出欠"] + 1).setValue(attendance || "");
      sh.getRange(r + 1, idx["回答日時"] + 1).setValue(formatDateTime_(ts));
      sh.getRange(r + 1, idx["要確認フラグ"] + 1).setValue(""); // 正常なら空
      return;
    }
  }

  // 見つからない場合は要確認扱いに（会シート末尾に追記でも可）
}

/***********************
 * イベント管理 → 会シート名検索
 ***********************/
function findEventSheetName_(ss, eventId) {
  const sh = ss.getSheetByName(SHEET_EVENTS);
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  const idx = indexMap_(headers);

  for (let r = 1; r < values.length; r++) {
    if (String(values[r][idx["イベントID"]]) === String(eventId)) {
      return values[r][idx["会シート名"]];
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

/***********************
 * デモ：未入金者へ催促メール送信
 * - 最新の「会_」シートを対象
 * - 出欠が「出席」かつ入金状況が「未入金」の人だけに送信
 ***********************/
function demo_sendPaymentReminders_latestEventSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getLatestEventSheet_(ss);
  if (!sheet) throw new Error("会_ で始まる会シートが見つかりません。");

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) throw new Error("会シートにデータがありません。");

  const headers = values[0].map(h => String(h).trim());
  const idx = indexMap_(headers);

  // 必要列チェック
  const required = ["氏名", "メールアドレス", "入金状況", "出欠"];
  for (const k of required) {
    if (idx[k] === undefined) throw new Error(`会シートに「${k}」列が見つかりません。見出しを確認してください。`);
  }

  const unpaidRows = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const name = row[idx["氏名"]];
    const email = row[idx["メールアドレス"]];
    const pay = String(row[idx["入金状況"]] || "").trim();
    const attendance = String(row[idx["出欠"]] || "").trim(); // ★追加

    if (!name || !email) continue;

    // ★「出席」かつ「未入金（または空）」だけ送信
    const isUnpaid = (pay === "" || pay === "未入金");
    const isAttending = (attendance === "出席");

    if (isAttending && isUnpaid) {
      unpaidRows.push({ rowNumber: r + 1, name: String(name), email: String(email) });
    }
  }

  if (unpaidRows.length === 0) {
    SpreadsheetApp.getUi().alert("未入金者がいません（入金状況が未入金の行がありません）。");
    return;
  }

  const subject = "【ご確認】未入金のご案内";
  const today = formatDate_(new Date());

  // 送信（デモなので同一メールでもOK）
  for (const u of unpaidRows) {
    const body =
`${u.name} 様

お世話になっております。
本日（${today}）時点で、入金状況が「未入金」となっております。

お手数ですが、ご確認のうえお支払い手続きをお願いいたします。
※すでにお支払い済みの場合は、本メールは行き違いとなりますためご容赦ください。

どうぞよろしくお願いいたします。`;

    MailApp.sendEmail({ to: u.email, subject, body });

    // 送信記録を備考に残す（デモで見せやすい）
    const noteCol = idx["備考"] !== undefined ? idx["備考"] + 1 : null;
    if (noteCol) {
      sheet.getRange(u.rowNumber, noteCol).setValue(`催促メール送信：${formatDateTime_(new Date())}`);
    }
  }

  SpreadsheetApp.getUi().alert(
    `未入金者 ${unpaidRows.length} 名に催促メールを送信しました。\n対象シート：${sheet.getName()}`
  );
}

/***********************
 * 最新の会シート（会_ で始まる）を取得
 ***********************/
function getLatestEventSheet_(ss) {
  const sheets = ss.getSheets().filter(s => s.getName().startsWith("会_"));
  if (sheets.length === 0) return null;

  // 最終更新で厳密には取れないので、名前順の末尾 or 最後に追加された順に近い方法として「末尾」を採用
  // もし複数ある場合でもデモなら十分
  sheets.sort((a, b) => a.getName().localeCompare(b.getName()));
  return sheets[sheets.length - 1];
}

/***********************
 * デモ：入金状況が「入金済」になったら
 * 入金確認日時を自動で入れる（会シート全て対象）
 *
 * 使い方：
 * 1) installPaymentOnEditTrigger_() を1回実行してトリガーを作る
 * 2) 会_ で始まる会シートの「入金状況」を編集すると自動反映
 ***********************/
function installPaymentOnEditTrigger_() {
  // 既存の同名トリガーが増えないように削除
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) {
    if (t.getHandlerFunction() === "onPaymentStatusEdit_") {
      ScriptApp.deleteTrigger(t);
    }
  }
  ScriptApp.newTrigger("onPaymentStatusEdit_")
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();

  SpreadsheetApp.getUi().alert("入金状況の自動日時入力トリガーを設定しました。\n会_ シートの「入金状況」を編集して動作確認してください。");
}

/**
 * onEdit トリガーで動く関数
 * - 会_ で始まるシートのみ対象
 * - 編集列が「入金状況」列の時だけ処理
 * - チェックボックス擬似ボタンの処理も担当
 * - 名簿マスターのチェックボックス操作も担当
 */
function onPaymentStatusEdit_(e) {
  // ★名簿マスターのチェック操作
  if (handleMasterActionCheckbox_(e)) return;

  // ★会シートのチェック操作
  if (handleActionCheckbox_(e)) return;

  // --- ここから既存の「入金状況 → 入金確認日時」処理 ---
  const range = e.range;
  const sheet = range.getSheet();

  // 会シートのみ
  if (!sheet.getName().startsWith("会_")) return;

  // ヘッダ行は無視
  if (range.getRow() === 1) return;

  // ヘッダ取得して列位置特定
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
  const idx = indexMap_(headers);

  // 必要な列が無い場合は何もしない
  if (idx["入金状況"] === undefined || idx["入金確認日時"] === undefined) return;

  const payCol = idx["入金状況"] + 1;
  const confirmedCol = idx["入金確認日時"] + 1;

  // 編集された列が「入金状況」以外なら無視
  if (range.getColumn() !== payCol) return;

  const newValue = String(range.getValue() || "").trim();
  const confirmedCell = sheet.getRange(range.getRow(), confirmedCol);

  // 入金済 → 日時を自動入力（デモ映えのため常に更新）
  if (newValue === "入金済" || newValue === "入金") {
    confirmedCell.setValue(formatDateTime_(new Date()));
    return;
  }

  // 未入金（または空）→ 日時をクリア（運用に合わせてON/OFF）
  if (newValue === "未入金" || newValue === "") {
    confirmedCell.clearContent();
    return;
  }

  // それ以外（例：要確認など）→ 何もしない
}

/***********************
 * チェックボックス擬似ボタンの処理
 * M2:M4 のチェックで各機能を実行
 ***********************/
function handleActionCheckbox_(e) {
  const range = e.range;
  const sheet = range.getSheet();
  if (!sheet.getName().startsWith("会_")) return false;

  // M2:M4 が操作チェックボックス（列M=13）
  const col = range.getColumn();
  const row = range.getRow();
  if (col !== 13) return false; // M列以外は無視
  if (![2, 3, 4].includes(row)) return false;

  const checked = range.getValue() === true;
  if (!checked) return true; // OFFは何もしない（処理済扱い）

  try {
    if (row === 2) {
      // 未入金メール送信（出席者のみ）
      demo_sendPaymentReminders_forSheet_(sheet);
    } else if (row === 3) {
      // ここは「入金済にした行へ入金日時反映」など別処理に差し替え可
      SpreadsheetApp.getUi().alert("入金状況を変更すると自動で入金確認日時が入ります（G列を編集してください）。");
    } else if (row === 4) {
      // トリガー設定
      installPaymentOnEditTrigger_();
    }
  } finally {
    // チェックを戻す（ボタンっぽくする）
    range.setValue(false);
  }
  return true;
}

/***********************
 * 名簿マスターのチェックボックス擬似ボタン処理
 * M2:M4 のチェックで各機能を実行
 ***********************/
function handleMasterActionCheckbox_(e) {
  const range = e.range;
  const sheet = range.getSheet();

  // 名簿マスターのみ
  if (sheet.getName() !== SHEET_MASTER) return false;

  // M2:M4 が操作チェックボックス（列M=13）
  const col = range.getColumn();
  const row = range.getRow();
  if (col !== 13) return false;
  if (![2, 3, 4].includes(row)) return false;

  const checked = range.getValue() === true;
  if (!checked) return true; // OFFは何もしない（処理済扱い）

  try {
    if (row === 2) {
      // 回答者IDを自動発行（空欄のみ）
      fillResponderIdsIfEmpty_(sheet);
    } else if (row === 3) {
      // 有効フラグを一括TRUE（空欄のみ）
      fillActiveFlagIfEmpty_(sheet);
    } else if (row === 4) {
      // テストメール送信（全員）
      sendTestMailToAllFromMaster_(sheet);
    }
  } finally {
    // チェックを戻す（ボタンっぽく）
    range.setValue(false);
  }
  return true;
}

/***********************
 * 回答者IDを空欄だけ自動発行（U001形式）
 ***********************/
function fillResponderIdsIfEmpty_(sheet) {
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(h => String(h).trim());
  const idx = indexMap_(headers);

  if (idx["回答者ID"] === undefined) throw new Error("名簿マスターに「回答者ID」列がありません。");

  let counter = 1;
  // 既存IDから最大を拾う（U001の最大を見て次から）
  for (let r = 1; r < values.length; r++) {
    const v = String(values[r][idx["回答者ID"]] || "");
    const m = v.match(/^U(\d+)$/);
    if (m) counter = Math.max(counter, parseInt(m[1], 10) + 1);
  }

  let filled = 0;
  for (let r = 1; r < values.length; r++) {
    const cell = sheet.getRange(r + 1, idx["回答者ID"] + 1);
    const cur = String(cell.getValue() || "").trim();
    if (cur) continue;

    const newId = "U" + String(counter).padStart(3, "0");
    cell.setValue(newId);
    counter++;
    filled++;
  }

  SpreadsheetApp.getUi().alert(`回答者IDを ${filled} 件発行しました。`);
}

/***********************
 * 有効フラグを空欄だけTRUEにする
 ***********************/
function fillActiveFlagIfEmpty_(sheet) {
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(h => String(h).trim());
  const idx = indexMap_(headers);

  if (idx["有効フラグ"] === undefined) throw new Error("名簿マスターに「有効フラグ」列がありません。");

  let filled = 0;
  for (let r = 1; r < values.length; r++) {
    const cell = sheet.getRange(r + 1, idx["有効フラグ"] + 1);
    const cur = String(cell.getValue() || "").trim();
    if (cur) continue;

    cell.setValue(true);
    filled++;
  }

  SpreadsheetApp.getUi().alert(`有効フラグを ${filled} 件 TRUE にしました。`);
}

/***********************
 * テストメールを名簿全員に送る（デモ用）
 ***********************/
function sendTestMailToAllFromMaster_(sheet) {
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(h => String(h).trim());
  const idx = indexMap_(headers);

  const required = ["氏名", "メールアドレス", "有効フラグ"];
  for (const k of required) {
    if (idx[k] === undefined) throw new Error(`名簿マスターに「${k}」列がありません。`);
  }

  let sent = 0;
  for (let r = 1; r < values.length; r++) {
    const name = values[r][idx["氏名"]];
    const email = values[r][idx["メールアドレス"]];
    const active = values[r][idx["有効フラグ"]];

    const isActive = (active === true) || (String(active).toUpperCase() === "TRUE") || (String(active) === "有効");
    if (!isActive) continue;
    if (!name || !email) continue;

    const subject = "【デモ】テストメール（名簿マスター）";
    const body =
`${name} 様

こちらはデモ用のテストメールです。
名簿マスターから対象者を抽出し、自動送信できることの確認です。

どうぞよろしくお願いいたします。`;

    MailApp.sendEmail({ to: String(email), subject, body });
    sent++;
  }

  SpreadsheetApp.getUi().alert(`テストメールを ${sent} 件送信しました。`);
}

/***********************
 * 特定の会シートだけを対象に未入金メール送信（出席者のみ）
 * 既存の demo_sendPaymentReminders_latestEventSheet のシート指定版
 ***********************/
function demo_sendPaymentReminders_forSheet_(sheet) {
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(h => String(h).trim());
  const idx = indexMap_(headers);

  const required = ["氏名", "メールアドレス", "入金状況", "出欠"];
  for (const k of required) {
    if (idx[k] === undefined) throw new Error(`会シートに「${k}」列が見つかりません：${k}`);
  }

  const subject = "【ご確認】未入金のご案内";
  const today = formatDate_(new Date());

  let count = 0;
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const name = row[idx["氏名"]];
    const email = row[idx["メールアドレス"]];
    const pay = String(row[idx["入金状況"]] || "").trim();
    const attendance = String(row[idx["出欠"]] || "").trim();

    if (!name || !email) continue;
    if (attendance !== "出席") continue;
    if (!(pay === "" || pay === "未入金")) continue;

    const body =
`${name} 様

お世話になっております。
本日（${today}）時点で、入金状況が「未入金」となっております。
お手数ですが、ご確認のうえお支払い手続きをお願いいたします。
※すでにお支払い済みの場合は、本メールは行き違いとなりますためご容赦ください。

どうぞよろしくお願いいたします。`;

    MailApp.sendEmail({ to: String(email), subject, body });
    count++;

    // 備考に記録
    if (idx["備考"] !== undefined) {
      sheet.getRange(r + 1, idx["備考"] + 1).setValue(`催促メール送信：${formatDateTime_(new Date())}`);
    }
  }

  SpreadsheetApp.getUi().alert(`未入金（出席者のみ）へ ${count} 件送信しました。\n対象シート：${sheet.getName()}`);
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("デモ操作")
    .addItem("入金トリガー設定（B）", "installPaymentOnEditTrigger_")
    .addItem("未入金メール送信（A）", "demo_sendPaymentReminders_latestEventSheet")
    .addItem("回作成＋送信（フォーム・個別URL）", "demo_createEvent_and_sendMails")
    .addSeparator()
    .addItem("名簿マスターに操作エリア設置", "setupMasterActionArea_")
    .addToUi();
}
