const SPREADSHEET_NAME = '忠義的零用錢記錄';
const RECORD_SHEET = '收支紀錄';
const GOAL_SHEET = '儲蓄目標';

function getSpreadsheet_() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('SPREADSHEET_ID');
  if (id) {
    try { return SpreadsheetApp.openById(id); } catch (e) {}
  }
  const files = DriveApp.getFilesByName(SPREADSHEET_NAME);
  const ss = files.hasNext() ? SpreadsheetApp.open(files.next()) : SpreadsheetApp.create(SPREADSHEET_NAME);
  props.setProperty('SPREADSHEET_ID', ss.getId());
  setupSheets_(ss);
  return ss;
}

function setupSheets_(ss) {
  let records = ss.getSheetByName(RECORD_SHEET);
  if (!records) records = ss.insertSheet(RECORD_SHEET);
  if (records.getLastRow() === 0) records.appendRow(['ID', '日期', '類型', '項目', '金額']);

  let goal = ss.getSheetByName(GOAL_SHEET);
  if (!goal) goal = ss.insertSheet(GOAL_SHEET);
  if (goal.getLastRow() === 0) goal.appendRow(['名稱', '目標金額']);
}

function doGet(e) {
  try {
    const data = getData_();
    const json = JSON.stringify(data);
    const callback = e && e.parameter && e.parameter.callback;
    if (callback && /^[A-Za-z_$][0-9A-Za-z_$]*$/.test(callback)) {
      return ContentService.createTextOutput(callback + '(' + json + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ok:false,error:String(err)}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    const action = e.parameter.action;
    const data = JSON.parse(e.parameter.data || '{}');
    if (action === 'addRecord') addRecord_(data);
    else if (action === 'deleteRecord') deleteRecord_(data.id);
    else if (action === 'saveGoal') saveGoal_(data);
    else throw new Error('未知操作：' + action);
    return ContentService.createTextOutput(JSON.stringify({ok:true})).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ok:false,error:String(err)}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getData_() {
  const ss = getSpreadsheet_();
  const recordsSheet = ss.getSheetByName(RECORD_SHEET);
  const values = recordsSheet.getDataRange().getValues();
  const records = values.slice(1).filter(r => r[0] !== '').map(r => ({
    id: String(r[0]), date: formatDate_(r[1]), type: String(r[2]), item: String(r[3]), amount: Number(r[4])
  }));

  const goalSheet = ss.getSheetByName(GOAL_SHEET);
  const goalValues = goalSheet.getDataRange().getValues();
  const goal = goalValues.length > 1 && goalValues[1][0] !== ''
    ? {name:String(goalValues[1][0]), amount:Number(goalValues[1][1])}
    : null;

  return {ok:true, records, goal, spreadsheetName:ss.getName(), spreadsheetUrl:ss.getUrl()};
}

function addRecord_(data) {
  const sheet = getSpreadsheet_().getSheetByName(RECORD_SHEET);
  sheet.appendRow([String(data.id), String(data.date), String(data.type), String(data.item), Number(data.amount)]);
}

function deleteRecord_(id) {
  const sheet = getSpreadsheet_().getSheetByName(RECORD_SHEET);
  const values = sheet.getDataRange().getValues();
  for (let i = values.length - 1; i >= 1; i--) {
    if (String(values[i][0]) === String(id)) { sheet.deleteRow(i + 1); return; }
  }
}

function saveGoal_(data) {
  const sheet = getSpreadsheet_().getSheetByName(GOAL_SHEET);
  if (sheet.getLastRow() < 2) sheet.appendRow([String(data.name), Number(data.amount)]);
  else sheet.getRange(2, 1, 1, 2).setValues([[String(data.name), Number(data.amount)]]);
}

function formatDate_(value) {
  if (value instanceof Date) return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(value);
}

// 可選：第一次設定時可手動執行 setup()，它會建立「忠義的零用錢記錄」。
function setup() {
  const ss = getSpreadsheet_();
  setupSheets_(ss);
  Logger.log(ss.getUrl());
}
