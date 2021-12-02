// グループごとにシートを分割する関数
function createSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('List'); //シート名
  const startRow = 1; //開始行
  const startCol = 1; //開始列
  const numCols = 6; //開始行からの対象列数
  let results = [];

  // グループ一覧を取得
  let groupList = getGroups(ss);

  // グループごとに処理
  for(let i in groupList){
    let groupId = groupList[i][0];
    let groupName = groupList[i][1];


    // フィルタを実行
    let flg = [groupId, 1]
    sheetFilter(sheet,flg);

    // フィルタリングされた範囲の値を取得
    let lastRow = sheet.getLastRow();

    // 表示セルのみ値を取得
    let values = sheet.getRange(startRow, startCol, lastRow, numCols).getValues().filter(function(_, i) {
      return i > 0 && !sheet.isRowHiddenByFilter(i + 1) //見出しを除いてコピー
      // return !sheet.isRowHiddenByFilter(i + 1)}); //見出しを含めてコピー
    });

    // // 新規スプレッドシートにコピー
    // if (values.length > 0) {
    //   copyToNewTemplate(groupName,values);
    // }
    results.push([groupName, values.length]);

  }
  Logger.log(results);
  const resultSheet = ss.getSheetByName('result'); //結果出力シート名
  resultSheet.getRange(2, 1, results.length, results[0].length).setValues(results);

  sheet.getFilter().remove(); //シートのフィルター条件を削除する
}


// グループ一覧を取得
function getGroups(ss) {
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
function sheetFilter(sheet,flg) {
  const sortCol1 = 6; //ソート列
  const sortCol2 = 7; //ソート列
  // ビルダを作成
  let rule1 = SpreadsheetApp.newFilterCriteria()
    .whenTextEqualTo(flg[0])　//条件を指定
    .build();
  let rule2 = SpreadsheetApp.newFilterCriteria()
  .whenTextEqualTo(flg[1])　//条件を指定
  .build();
  // フィルタ済みの場合
  if(sheet.getFilter() != null) {
     sheet.getFilter().remove(); //シートのフィルター条件を削除する
  }
  // フィルタを実行
  sheet.getDataRange().createFilter()
    .setColumnFilterCriteria(sortCol1, rule1)
    .setColumnFilterCriteria(sortCol2, rule2);
}


// // 新しいシートにコピー
// function copyToNewSheet(groupName,values) {
//   let newFileName = groupName + '_一覧';
//   // 新規シートを作成
//   let newFile = DriveApp.getFileById(SpreadsheetApp.create(newFileName).getId());
//   let newSpreadsheet = SpreadsheetApp.openById(newFile.getId());

//   // 新規シートにペースト
//   let newSheet = newSpreadsheet.getSheets()[0];
//   newSheet.getRange(newSheet.getLastRow() + 1, 1, values.length, values[0].length).setValues(values);
//   // セルの整形
//   formatCells(newSheet);
//   Logger.log(newFileName +' を作成しました');
// }


// 雛形を複製してコピー
function copyToNewTemplate(groupName,values) {
  // テンプレートファイル
  const templateFile = DriveApp.getFileById('1ew2DZe2825GM_bsuvXSt12HPZPXe4wl6Ozm_M95uvjk');
  // 出力先
  const OutputFolder = DriveApp.getFolderById('1VlBQ-xxig36V-c4SHbhYS4E9GL_n_ZFA');

  // 出力ファイル名
  const OutputFileName = templateFile.getName().replace('雛形', '') + '一覧表_' + groupName
  +'_'+Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd');

  // コピー作成
  let newFile = templateFile.makeCopy(OutputFileName, OutputFolder);
  let newSpreadsheet = SpreadsheetApp.openById(newFile.getId());
  // 新規シートにペースト
  let newSheet = newSpreadsheet.getSheets()[0];
  newSheet.getRange(newSheet.getLastRow() + 1, 1, values.length, values[0].length).setValues(values);
  // セルの整形
  formatCells(newSheet);
  Logger.log(OutputFileName +' を作成しました');
}



// 対象範囲のセルを整形する
function formatCells(sheet){
  const startRow = 1; //開始行
  const startCol = 1; //開始列
  let lastRow = sheet.getLastRow();
  let lastCol = sheet.getLastColumn();

  let range = sheet.getRange(startRow,startCol,lastRow,lastCol); //対象範囲
  // range.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP); //文字列を折り返す
  range.setBorder(true,true,true,true,true,true,'#000000',SpreadsheetApp.BorderStyle.SOLID); //枠線を引く
}


// メニュー表示
function onOpen() {
  const ui = SpreadsheetApp.getUi(); // UIクラス取得
  const menu = ui.createMenu('GASメニュー'); // スプレッドシートにメニューを追加
  menu.addItem('シート生成','createSheets'); // 関数をセット
  menu.addToUi(); // スプレッドシートに反映
}
