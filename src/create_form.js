const SHEETNAME_CUSTOM = "設定";
const KEY_QUESTION_TITLE="集計用キー(編集不要)";
const CUSTOM_CELL_POSITON = {
  FORM_ID: "B1",

  DATA_SHEETNAME: "D1",
  FORM_DATA_KEY_POS: "D4",
  FORM_TITLE_POS: "D2",
  FORM_DESCRIPTION_POS: "D3",

  SELECTION_TITLE_POS: "D5",
  SELECTION_DESC_POS: "D6",
  SELECTION_DATA_START_POS: "D7",
};

function createForm() {
  // *** データを取得する対象のSpreadSheet ***
  // const ss = SpreadsheetApp.openById("xxxx");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const datasheet = ss.getSheetByName(customVal("DATA_SHEETNAME"));
  // *** 作成先アンケートとなるForm ***
  const form = FormApp.openById(customVal("FORM_ID"));

  // その他情報
  const formInfo = getFormInfo();

  // アンケートの基本データを設定
  form.setTitle(formInfo.title);
  form.setDescription(formInfo.desc);

  // アンケートを作成
  // 既存データから、対象を検索
  var mainQuestion = null;
  var keyQuestion = null;
  let itemsNum = form.getItems();
  // 最初に現れた複数選択質問を対象とする
  for (let i = 0; i < itemsNum.length; i++) {
    let c = form.getItems()[i];
    if(c.getTitle() == KEY_QUESTION_TITLE){
      keyQuestion = c.asMultipleChoiceItem();
    }
    else if(c.getType() == FormApp.ItemType.MULTIPLE_CHOICE) {
      mainQuestion = c.asMultipleChoiceItem();
    }
  }

  // ******************************
  // *** 必須キー ***
  // ******************************
  if(keyQuestion == null){
    keyQuestion = form.addMultipleChoiceItem();
  }
  keyQuestion
    .setTitle(KEY_QUESTION_TITLE)
    .setChoices([keyQuestion.createChoice(formInfo.key)]);
  keyQuestion.setRequired(true);
  
  // ******************************
  // *** アンケート質問 ***
  // ******************************
  if(mainQuestion == null){
    console.log("対象質問なし");
    mainQuestion = form.addMultipleChoiceItem();
  }

  const firstRow = 7;
  const lastRow = 9;

  const dataRows = lastRow - (firstRow - 1);

  let choiceList = [];
  datasheet
    .getRange(firstRow, 1, dataRows, 3)
    .getDisplayValues()
    .map((question) => {
      return {
        title: question[0],
        helpText: question[1],
        point: question[2],
      };
    })
    .forEach((choice) => {
      if (choice.title != "") {
        let choiceObj = mainQuestion.createChoice(choice.title);

        choiceList.push(choiceObj);
      }
    });

  mainQuestion
    .setTitle(formInfo.selections[0].title)
    .setHelpText(formInfo.selections[0].desc)
    .setChoices(choiceList);

    
var img = UrlFetchApp.fetch('https://www.google.com/images/srpr/logo4w.png');
form.addImageItem()
    .setTitle('Google')
    .setHelpText('Google Logo') // The help text is the image description
    .setImage(img);
}

/**
 * カスタマイズ設定を取得
 * @param {*} key 設定の種類(CUSTOM_CELL_POSITONのキー)
 * @returns SHEETNAME_CUSTOMシートに設定された値
 */
function customVal(key) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const customSheet = ss.getSheetByName(SHEETNAME_CUSTOM);
  return customSheet.getRange(CUSTOM_CELL_POSITON[key]).getValue();
}

/**
 * アンケート生成用の情報を取得
 * @returns アンケート生成用の情報
 */
function getFormInfo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const datasheet = ss.getSheetByName(customVal("DATA_SHEETNAME"));

  let info = {
    key: datasheet.getRange(customVal("FORM_DATA_KEY_POS")).getValue().toString(),
    title: datasheet.getRange(customVal("FORM_TITLE_POS")).getValue().toString(),
    desc: datasheet.getRange(customVal("FORM_DESCRIPTION_POS")).getValue().toString(),
    selections: [],
  };

  info.selections.push({
    title: datasheet.getRange(customVal("SELECTION_TITLE_POS")).getValue().toString(),
    desc: datasheet.getRange(customVal("SELECTION_DESC_POS")).getValue().toString(),
    selections: [],
  });

  return info;
}

// 以下、ただのメモ
// スプレッドシートに紐づけ
// form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

// クイズにする
// form.setDescription(formDescription).setIsQuiz(true);
