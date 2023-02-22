/**
 * データからアンケートFormを作成
 */
function createForm() {
  // *** データを取得する対象のSpreadSheet ***
  // const ss = SpreadsheetApp.openById("xxxx");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETNAME_DATA);

  // アンケートFormの情報
  const formInfo = getFormInfo();
  // *** 作成先アンケートとなるForm ***
  let form = null;
  // アンケート用質問
  let mainQuestion = null;

  const formId = customVal(sheet, CUSTOM_CELL_LABEL.FORM_ID);
  if (formId) {
    // ID指定あれば既存のFormを流用
    form = FormApp.openById(formId);
    // 既存の質問の中から対象を検索
    let itemsNum = form.getItems();
    for (let i = 0; i < itemsNum.length; i++) {
      let c = form.getItems()[i];
      if (c.getTitle() == formInfo.mainSelection.title) {
        mainQuestion = c.asCheckboxItem();
      }
    }
  } else {
    // ID指定なければFormを新規作成
    form = FormApp.create(formInfo.title);
    // このスクリプトの設定されているSpreadSheetに回答を保存
    form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  }
  if (mainQuestion == null) {
    mainQuestion = form.addCheckboxItem();
  }

  // アンケートの基本データを設定
  form.setTitle(formInfo.title);
  form.setDescription(formInfo.desc);

  // ******************************
  // *** アンケート質問 ***
  // ******************************
  mainQuestion
    .setTitle(formInfo.mainSelection.title)
    .setHelpText(formInfo.mainSelection.desc)
    .setChoices(
      formInfo.mainSelection.selections.map((obj) =>
        mainQuestion.createChoice(createSelectText(obj))
      )
    );
}

/**
 * getMoviesDataの情報から、画面表示用の選択肢文字列を生成
 * @param {*} obj getMoviesDataの情報
 * @returns 画面表示用の選択肢文字列
 */
function createSelectText(obj) {
  return (
    `【${obj.title}】 ${obj.desc}` +
    (obj.count && obj.count > 0 ? ` （投票数:${obj.count}）` : "")
  );
}

/**
 * アンケート生成用の情報を取得
 * @returns アンケート生成用の情報
 */
function getFormInfo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETNAME_DATA);

  let info = {
    title: customVal(sheet, CUSTOM_CELL_LABEL.FORM_TITLE),
    desc: customVal(sheet, CUSTOM_CELL_LABEL.FORM_DESCRIPTION),
    mainSelection: {
      title: customVal(sheet, CUSTOM_CELL_LABEL.SELECTION_TITLE),
      desc: customVal(sheet, CUSTOM_CELL_LABEL.SELECTION_DESC),
      selections: getMoviesData(sheet),
    },
  };

  return info;
}

function getMoviesData(sheet) {
  let row = findKeyFromA(sheet, CUSTOM_CELL_LABEL.SELECTION_DATA_START);
  // ヘッダ分の次の行
  row += 1;
  let table = customTable(sheet, row, 1, 4);
  return (
    table
      .map((line) => {
        return {
          title: line[0],
          desc: line[1],
          image: line[2],
          count: line[3],
        };
      })
      // 投票数で降順ソートする
      .sort((data1, data2) => {
        return data2.count - data1.count;
      })
  );
}

////////////////////////////////
// 以下、ライブラリ（システムF仕様を含まない）
/**
 * A列から対象文字列を検索し、行数を返す
 * @param {*} sheet 対象シート
 * @param {*} key 対象文字列
 * @returns 行
 */
function findKeyFromA(sheet, key) {
  // A1から探す
  for (let row = 1; row <= CUSTOM_SEARCH_ROW_MAX; row++) {
    if (sheet.getRange(row, 1).getDisplayValue() == key) {
      return row;
    }
  }
  // 見つからなかった
  return 0;
}

/**
 * カスタマイズ設定を取得
 * A列からkeyを検索し、見つかった行のB列の値を返す
 * @param {*} sheet 対象シート
 * @param {*} key 設定の種類
 * @returns 設定値
 */
function customVal(sheet, key) {
  let row = findKeyFromA(sheet, key);
  if (row) {
    return sheet.getRange(row, 2).getDisplayValue();
  }
  // 見つからなかった
  return "";
}
/**
 * カスタマイズ設定を設定
 * A列からkeyを検索し、見つかった行に値を設定(デフォルトはB列)
 * @param {*} sheet 対象シート
 * @param {*} key 設定の種類
 * @param {*} val 設定値
 */
function customSet(sheet, key, val, col = 2) {
  let row = findKeyFromA(sheet, key);
  if (row) {
    return sheet.getRange(row, col).setValue(val);
  }
}

/**
 * カスタマイズ設定(表形式)を取得
 * 指定されたfirstRow,firstColから有効なデータを表形式として取得
 * @param {*} sheet 対象シート
 * @returns 二重配列
 */
function customTable(sheet, firstRow, firstCol, colNum, isChange = true) {
  // 指定された行から、firstCol列が空になる行まで
  for (let r = firstRow; r - firstRow < CUSTOM_SEARCH_ROW_MAX; r++) {
    let value = sheet.getRange(r, firstCol).getDisplayValue();
    if (value == "") {
      let rowNum = r - firstRow;
      console.log(
        `get table from SHEET[${sheet.getSheetName()}] ` +
          `at range(row:${firstRow}-${r}, col:${firstCol}-${firstCol + colNum})`
      );
      if (rowNum == 0) {
        // データなし
        return [];
      }

      // 見つかった範囲を二重配列で返す
      return sheet
        .getRange(firstRow, firstCol, rowNum, colNum)
        .getDisplayValues();
    } else {
      if (isChange) {
        // Form作成時に違いが出ないようにタイトルを修正
        sheet
          .getRange(r, firstCol)
          .setValue(value.trim().replaceAll(/[\s　]+/g, " "));
      }
    }
  }
  return [];
}
