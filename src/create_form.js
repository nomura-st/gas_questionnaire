// 対象シート名
const SHEETNAME_DATA = "データ";
// 設定ラベル
const CUSTOM_CELL_LABEL = {
  FORM_ID: "アンケート用FormID",
  FORM_TITLE: "タイトル",
  FORM_DESCRIPTION: "説明",
  SELECTION_TITLE: "選択肢のタイトル",
  SELECTION_DESC: "選択肢の説明",
  SELECTION_DATA_START: "候補",
};
// 設定ラベルを探す最大行
const CUSTOM_SEARCH_ROW_MAX = 100;

/**
 * WEBアプリ用のエントリポイント
 */
function doGet() {
  createForm();
}

/**
 * データからアンケートFormを作成
 */
function createForm() {
  // *** データを取得する対象のSpreadSheet ***
  // const ss = SpreadsheetApp.openById("xxxx");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETNAME_DATA);
  // *** 作成先アンケートとなるForm ***
  const form = FormApp.openById(customVal(sheet, CUSTOM_CELL_LABEL.FORM_ID));

  // その他情報
  const formInfo = getFormInfo();

  // アンケートの基本データを設定
  form.setTitle(formInfo.title);
  form.setDescription(formInfo.desc);

  // アンケートを作成
  // 既存データから、対象を検索
  var mainQuestion = null;
  let itemsNum = form.getItems();
  // 最初に現れた複数選択質問を対象とする
  for (let i = 0; i < itemsNum.length; i++) {
    let c = form.getItems()[i];
    if (c.getTitle() == formInfo.mainSelection.title) {
      mainQuestion = c.asCheckboxItem();
    }
  }
  if (mainQuestion == null) {
    console.log("対象質問なし");
    mainQuestion = form.addCheckboxItem();
  }

  // ******************************
  // *** アンケート質問 ***
  // ******************************
  mainQuestion
    .setTitle(formInfo.mainSelection.title)
    .setHelpText(formInfo.mainSelection.desc)
    .setChoices(
      formInfo.mainSelection.selections.map((obj) =>
        mainQuestion.createChoice(obj.title)
      )
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
    desc: customVal(sheet, CUSTOM_CELL_LABEL.CUSTOM_CELL_LABEL),
    mainSelection: {
      title: customVal(sheet, CUSTOM_CELL_LABEL.SELECTION_TITLE),
      desc: customVal(sheet, CUSTOM_CELL_LABEL.SELECTION_DESC),
      selections: [],
    },
  };
  // アンケート選択肢selectionsを生成
  {
    let row = findKeyFromA(sheet, CUSTOM_CELL_LABEL.SELECTION_DATA_START);
    // ヘッダ分の次の行
    row += 2;
    let table = customTable(sheet, row, 1, 4);
    info.mainSelection.selections = table.map((line) => {
      return {
        title: line[0],
        year: line[1],
        desc: line[2],
        count: line[3],
      };
    });
  }

  return info;
}

////////////////////////////////
// 以下、ライブラリ
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
 * カスタマイズ設定(表形式)を取得
 * 指定されたfirstRow,firstColから有効なデータを表形式として取得
 * @param {*} sheet 対象シート
 * @returns 二重配列
 */
function customTable(sheet, firstRow, firstCol, colNum) {
  // 指定された行から、firstCol列が空になる行まで
  for (let r = firstRow; r - firstRow < CUSTOM_SEARCH_ROW_MAX; r++) {
    if (sheet.getRange(r, firstCol).getDisplayValue() == "") {
      let rowNum = r - firstRow;
      if (rowNum == 0) {
        // データなし
        return [];
      }

      // 見つかった範囲を二重配列で返す
      return sheet
        .getRange(firstRow, firstCol, rowNum, colNum)
        .getDisplayValues();
    }
  }
  return [];
}

////////////////////////////////
// 以下、ただのメモ
// スプレッドシートに紐づけ
// form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

// クイズにする
// form.setDescription(formDescription).setIsQuiz(true);

// 画像要素
// var img = UrlFetchApp.fetch('https://www.google.com/images/srpr/logo4w.png');
// form.addImageItem()
//     .setTitle('Google')
//     .setHelpText('Google Logo') // The help text is the image description
//     .setImage(img);
