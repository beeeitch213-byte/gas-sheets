// ==== 決算情報取得銘柄 ====
// 決算情報を取得したい銘柄を、ダイアログ(KessanTargetDialog)で登録銘柄からチェックで選び、
// 別のスプレッドシート「決算情報」の「決算情報取得銘柄」シートに書き出す。
// 決算スケジュールの作成や決算の評価は Claude の会話で行い、そのときの対象銘柄としてこのシートを読む想定
// (株式情報は大きいため、小さな別ファイルにしておく)。
const KESSAN_INFO_FILE_NAME = '決算情報';

const KESSAN_TARGET_SHEET_NAME = '決算情報取得銘柄';

const KESSAN_INFO_FILE_ID_KEY = 'KESSAN_INFO_SPREADSHEET_ID'; // 「決算情報」のファイルIDをスクリプトプロパティに覚えておく

const KESSAN_TARGET_HEADERS = ['コード', '銘柄名', '業種', '追加日'];

const KESSAN_TARGET_FIRST_ROW = 3; // 1行目=最終更新日時、2行目=見出し

/** 「機能 → 決算情報取得銘柄」 */
function kessanTargetUI() {
  const html = HtmlService.createHtmlOutputFromFile('KessanTargetDialog').setWidth(680).setHeight(640);
  SpreadsheetApp.getUi().showModalDialog(html, '決算情報取得銘柄');
}

/**
 * 「決算情報」のスプレッドシートを返す。create=trueなら、無いときに作る(無ければnull)。
 * 覚えているIDのファイルがゴミ箱に入っているなどで開けないときは、同じ名前のファイルを探し、それも無ければ作り直す。
 */
function getKessanInfoSpreadsheet_(create) {
  const props = PropertiesService.getScriptProperties();
  const savedId = props.getProperty(KESSAN_INFO_FILE_ID_KEY);
  if (savedId) {
    try {
      if (!DriveApp.getFileById(savedId).isTrashed()) return SpreadsheetApp.openById(savedId);
    } catch (e) {}
  }
  const found = DriveApp.getFilesByName(KESSAN_INFO_FILE_NAME);
  while (found.hasNext()) {
    const file = found.next();
    if (file.getMimeType() === MimeType.GOOGLE_SHEETS && !file.isTrashed()) {
      props.setProperty(KESSAN_INFO_FILE_ID_KEY, file.getId());
      return SpreadsheetApp.openById(file.getId());
    }
  }
  if (!create) return null;
  const ss = SpreadsheetApp.create(KESSAN_INFO_FILE_NAME);
  props.setProperty(KESSAN_INFO_FILE_ID_KEY, ss.getId());
  return ss;
}

/** 「決算情報取得銘柄」シートの銘柄を [{code, name, industry, addedAt}] で返す(シートが無ければ空) */
function readKessanTargets_(ss) {
  const sheet = ss ? ss.getSheetByName(KESSAN_TARGET_SHEET_NAME) : null;
  if (!sheet || sheet.getLastRow() < KESSAN_TARGET_FIRST_ROW) return [];
  return sheet.getRange(KESSAN_TARGET_FIRST_ROW, 1, sheet.getLastRow() - KESSAN_TARGET_FIRST_ROW + 1, KESSAN_TARGET_HEADERS.length)
    .getValues()
    .filter(r => String(r[0]).trim() !== '')
    .map(r => ({ code: String(r[0]).trim().toUpperCase(), name: String(r[1]), industry: String(r[2]), addedAt: r[3] }));
}

/**
 * KessanTargetDialogを開いたときに呼ばれる。
 * stocks: 当月シートの登録銘柄 [{code, name, industry}](シートの並び順)
 * selected: 「決算情報取得銘柄」シートに今ある銘柄コード
 */
function getKessanTargetDialogData() {
  const sheet = getTargetSheet(new Date());
  const stocks = [];
  if (sheet) {
    const lastCol = sheet.getLastColumn();
    const headRows = sheet.getRange(2, 1, 2, lastCol).getValues(); // 2行目=銘柄コード・業種、3行目=銘柄名
    for (let col = STOCK_START_COL; col <= lastCol; col += STOCK_COL_WIDTH) {
      const code = headRows[0][col - 1];
      if (!code) continue;
      stocks.push({ code: String(code).trim().toUpperCase(), name: String(headRows[1][col - 1]), industry: String(headRows[0][col] || '') });
    }
  }
  const selected = readKessanTargets_(getKessanInfoSpreadsheet_(false)).map(t => t.code);
  return { stocks: stocks, selected: selected };
}

/**
 * KessanTargetDialogの「更新」から呼ばれる。選んだ銘柄で「決算情報」の「決算情報取得銘柄」シートを書き直す。
 * 前から入っていた銘柄は追加日を引き継ぐ。ファイルやシートが無ければ作る。
 * targets: [{code, name, industry}](ダイアログのリストの順)
 */
function saveKessanTargets(targets) {
  const ss = getKessanInfoSpreadsheet_(true);
  let sheet = ss.getSheetByName(KESSAN_TARGET_SHEET_NAME);
  if (!sheet) {
    // 新しく作ったファイルの最初の空のシート(シート1)は、名前を変えてそのまま使う
    const sheets = ss.getSheets();
    const blank = sheets.length === 1 && sheets[0].getLastRow() === 0 ? sheets[0] : null;
    sheet = blank ? blank.setName(KESSAN_TARGET_SHEET_NAME) : ss.insertSheet(KESSAN_TARGET_SHEET_NAME);
  }

  const addedAtByCode = {};
  readKessanTargets_(ss).forEach(t => { addedAtByCode[t.code] = t.addedAt; });
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  sheet.clear();
  const fmt = (d, f) => Utilities.formatDate(d, 'Asia/Tokyo', f);
  sheet.getRange(1, 1).setValue('最終更新: ' + fmt(new Date(), 'yyyy/MM/dd HH:mm') + '(株式情報の「機能 → 決算情報取得銘柄」から更新)')
    .setFontColor('#666666');
  sheet.getRange(2, 1, 1, KESSAN_TARGET_HEADERS.length).setValues([KESSAN_TARGET_HEADERS]).setFontWeight('bold').setBackground('#f1f3f4');
  sheet.setFrozenRows(2);

  if (targets.length > 0) {
    const values = targets.map(t => {
      const code = String(t.code).trim().toUpperCase();
      return [code, t.name || '', t.industry || '', addedAtByCode[code] instanceof Date ? addedAtByCode[code] : today];
    });
    // 表示形式は値より先に設定する(銘柄コードを文字列にしておかないと、7203などが数値になる)
    sheet.getRange(KESSAN_TARGET_FIRST_ROW, 1, values.length, 1).setNumberFormat('@');
    sheet.getRange(KESSAN_TARGET_FIRST_ROW, 4, values.length, 1).setNumberFormat('yyyy/MM/dd');
    sheet.getRange(KESSAN_TARGET_FIRST_ROW, 1, values.length, KESSAN_TARGET_HEADERS.length).setValues(values);
  }
  sheet.autoResizeColumns(1, KESSAN_TARGET_HEADERS.length);

  return { message: '「' + KESSAN_INFO_FILE_NAME + '」の「' + KESSAN_TARGET_SHEET_NAME + '」を更新しました(' + targets.length + '銘柄)', url: ss.getUrl() };
}
