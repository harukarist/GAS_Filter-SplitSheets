const ss = SpreadsheetApp.getActiveSpreadsheet();
const sheet = ss.getSheetByName('List'); //シート名
const startCol = 1; //開始列
const numCols = 6; //開始行からの対象列数
const sortCol = 6; //ソート列


// グループごとにシートを分割する関数
function createSheets() {
  // グループ一覧を取得
  let groupList = getGroups();

  // グループごとに処理
  for(let i in groupList){
    let groupId = groupList[i][0];
    let groupName = groupList[i][1];
    let newFileName = groupName + '_一覧リスト'

    // フィルタを実行
    sheetFilter(groupId);

    // フィルタリングされた範囲の値を取得
    let values = sheet.getDataRange().getValues().filter(function(_, i) {return !sheet.isRowHiddenByFilter(i + 1)}); //見出しを含めてコピー
    // let values = sheet.getDataRange().getValues().filter(function(_, i) {return i > 1 && !sheet.isRowHiddenByFilter(i + 1)}); //見出しを除いてコピー
    // Logger.log(values);

    // 新規スプレッドシートにコピー
    if (values.length > 0) {
      copyToNewSheet(newFileName,values);
    } 
  }
}


// グループ一覧を取得
function getGroups() {
  const sheet = ss.getSheetByName('Groups');
  const startCol = 1; //開始列
  const startRow = 2; //開始行
  const numCols = 2; //開始行からの対象列数
  let lastRow = sheet.getLastRow();
  // let lastCol = sheet.getLastColumn();

  let range = sheet.getRange(startRow,startCol,lastRow-startRow+1,numCols); //対象範囲
  let values = range.getValues();
  return values;
}


// フィルタを実行
function sheetFilter(flg) {
  let rule = SpreadsheetApp.newFilterCriteria()
    .whenTextEqualTo(flg)　//条件を指定
    .build();
  if(sheet.getFilter() != null) {
     sheet.getFilter().remove(); //シートのフィルター条件を削除する
  }
  sheet.getDataRange().createFilter()
  .setColumnFilterCriteria(sortCol, rule);
}


// 新しいシートにコピー
function copyToNewSheet(newSheetName,values) {
  // 新規シートを作成
  let newFile = DriveApp.getFileById(SpreadsheetApp.create(newSheetName).getId());
  let newSpreadsheet = SpreadsheetApp.openById(newFile.getId());

  // 新規シートにペースト
  let newSheet = newSpreadsheet.getSheets()[0];
  newSheet.getRange(newSheet.getLastRow() + 1, 1, values.length, values[0].length).setValues(values);
  Logger.log(newSheetName +' を作成しました');
}


// メニュー表示
function onOpen() {
  const ui = SpreadsheetApp.getUi(); // UIクラス取得
  const menu = ui.createMenu('GASメニュー'); // スプレッドシートにメニューを追加
  menu.addItem('シート生成','createSheets'); // 関数をセット
  menu.addToUi(); // スプレッドシートに反映
}
