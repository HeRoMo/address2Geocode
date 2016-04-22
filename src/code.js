"use strict"
var APP_NAME = 'Address2Geocode'

function onOpen(){
  var ui = SpreadsheetApp.getUi();
  var addon = ui.createAddonMenu();
  addon.addItem('サイドバーを表示', 'showSidebar');
  addon.addToUi();
}

function onInstall() {
  onOpen();
}

/*
 * サイドバーの表示
 */
function showSidebar(){
  var ui = SpreadsheetApp.getUi();
  var sidebar = HtmlService.createHtmlOutputFromFile('sidebar');
  sidebar.setSandboxMode(HtmlService.SandboxMode.IFRAME)
  sidebar.setTitle(APP_NAME);
  ui.showSidebar(sidebar);
}

/**
 *  文字列から緯度経度を求める関数（シンプルジオコーディング実験 API版）
 *  @param 住所を表す文字列
 *  @return 緯度、経度の配列
 */
function geocode(address){
  if (!address||address == ''){return ''}
  var url = "http://newspat.csis.u-tokyo.ac.jp/cgi-bin/simple_geocode.cgi?addr=" + address;
  var res = UrlFetchApp.fetch(url)
  if(res.getResponseCode()!=200){return 'ERROR'}
  var xml = XmlService.parse(res.getContentText());
  var root = xml.getRootElement()
  var candidate = root.getChild("candidate")
  if(!candidate) {return 'No Result.'}
  var latitude = candidate.getChildText('latitude');
  var longitude = candidate.getChildText('longitude');
  return [ [ latitude , longitude ] ];
}

/**
 * 選択したセル(複数のカラムを選択していれば左端のカラム)に記載されている住所で
 * 緯度経度を求め、出力カラムに出力する。
 * @param outCol [integer] 緯度を書き込むカラムのインデックス。経度はoutCol+1のカラムに書き込む
 */
function batchGeocode(outCol){
  var outNumCols = 2 //出力のカラム幅。緯度と経度で2カラム

  //選択されているレンジから住所を取得し、変換する
  var addrRange = reselectRange();
  var addrRow = addrRange.getRow(); //レンジの左上のセルの行インデックス
  var addrNumRows = addrRange.getNumRows(); //レンジの行数
  var values = addrRange.getValues();

  var sheet = SpreadsheetApp.getActiveSheet()
  var outRange = sheet.getRange(addrRow, outCol,addrNumRows,outNumCols) //出力レンジ
  var outRangeVals = outRange.getValues();
  var outRangeFoms = outRange.getFormulas();
  var d = data(outRangeVals,outRangeFoms)
  var outVals = []
  for(var i = 0; i < values.length; i++){
    var code = d(i); // 出力セルの値を一旦格納
    if (code == null){ //出力セルが空の時だけ処理
      code = geocode(values[i][0]) //レンジの左端の列をループ
      if(!Array.isArray(code)){code = ['','']}else{code = code[0]}
    }
    outVals.push(code)
  }
  outRange.setValues(outVals)

  var ui = SpreadsheetApp.getUi();
  ui.alert("処理を終了しました" , ui.ButtonSet.OK);
}

/*
 * value と formulaを比べて formulaがあればformulaを
 * なければvalueを返す関数を返す
 * @param values [object[][]] 出力レンジのvalue
 * @param formulas [object[][]] 出力レンジのformula
 * @return [Function] 関数
 */
function data(values, formulas){
  return function(i){
    if(!values[i][0] && !values[i][1] ){
      return null
    } else {
      var a = !formulas[i][0] ? values[i][0] : formulas[i][0]
      var b = !formulas[i][1] ? values[i][1] : formulas[i][1]
      return [a,b]
    }
  }
}

/**
 * カラム等が選択された時に実際に値が入っているレンジを選択し直す
 * @return {[Range]} 実際に値の入っているレンジ
 */
function reselectRange(){
  var range = SpreadsheetApp.getActiveRange();
  var row = range.getRowIndex();
  var col = range.getColumn()
  var sheet = range.getSheet()
  var dRange = sheet.getDataRange()
  var dRow = Math.min(dRange.getLastRow(),range.getLastRow())+1;
  Logger.log("rowindex: %s, column: %s, lastrow %s", row, col, dRow)
  var activeRange = sheet.getRange(row, col, dRow-row)
  sheet.setActiveRange(activeRange)
  return activeRange
}
