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
  for (let i = 0; i < itemsNum.length; i++) {
    let c = form.getItems()[i];
    if (c.getTitle() == formInfo.mainSelection.title) {
      mainQuestion = c.asCheckboxItem();
    }
  }

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

  // // ******************************
  // // *** ヒント ***
  // // ******************************
  // // 一旦すべて削除
  // form
  //   .getItems(FormApp.ItemType.IMAGE)
  //   .forEach((item) => form.deleteItem(item));
  // // 作成
  // formInfo.mainSelection.selections.forEach((obj) => {
  //   if (obj.desc.trim() == "" && obj.image.trim() == "") {
  //     return;
  //   }

  //   let item = form.addImageItem();
  //   // タイトル
  //   item.setTitle(obj.title);
  //   // 画像
  //   if (obj.image.trim() != "") {
  //     let img = UrlFetchApp.fetch(obj.image);
  //     item.setImage(img);
  //     item.setWidth(300);
  //   } else {
  //     // 画像なし
  //   }
  //   // 説明文
  //   if (obj.desc.trim() != "") {
  //     item.setHelpText(obj.desc);
  //   }
  // });
}

function createSelectText(obj) {
  return (
    `【${obj.title}】  ${obj.desc}  ` +
    (obj.count > 0 ? `(投票数:${obj.count})` : "★NEW★")
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
      selections: getMoviesData(sheet),
    },
  };

  return info;
}

function getMoviesData(sheet) {
  let row = findKeyFromA(sheet, CUSTOM_CELL_LABEL.SELECTION_DATA_START);
  // ヘッダ分の次の行
  row += 2;
  let table = customTable(sheet, row, 1, 4);
  return table.map((line) => {
    return {
      title: line[0],
      desc: line[1],
      image: line[2],
      count: line[3],
    };
  });
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

// let check = document.querySelector("input[type='checkbox']");
// check.addEventListener("click", (event) => {
//   let target = event.currentTarget;
//   let table = document.createElement("table");
//   target.appendChild(table);
//   fetch("xxxxxx")
//     .then((response) => response.json())
//     .then((data) => {
//       console.log(data);
//       target.appendChild(data);
//     });
// });

/**
 * WEBアプリ用のエントリポイント
 */
function doGet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETNAME_DATA);
  let payload = getMoviesData(sheet)
    .map((data) => {
      let ret = "";
      if (data.desc) {
        ret += `<h3>${data.title}</h3><span>${data.desc.replace(
          "\\n",
          "<br>"
        )}</span>`;
      }
      if (data.image) {
        ret += `<image src="${data.image}">`;
      }
      if (ret) {
        ret = `<div>${ret}</div>`;
      }
      return ret;
    })
    .join("");

  var output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.TEXT);
  output.setContent(payload);

  return output;
}
