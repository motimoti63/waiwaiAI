/**
 * ========== HTML操作.gs ==========
 */
//カスタムメニュー
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('カスタムメニュー')
    .addItem('都道府県市区選択', 'showDialog')
    .addItem('条件マッチング','performMatching')
    .addItem('種別とジャンル', 'showMultiSelectDialog')
    .addToUi();
}

//種別とジャンル
function showMultiSelectDialog() {
  const html = HtmlService.createHtmlOutputFromFile('ルール')
      .setWidth(400)
      .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, '種別とジャンル');
}

function getRuleSheetData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("ルール");
  const dataRange = sheet.getRange(6, 1, sheet.getLastRow() - 5, sheet.getLastColumn());
  const values = dataRange.getValues();

  // データをNo列でソート
  values.sort((a, b) => a[0] - b[0]);

  const structuredData = values.map(row => ({
    category: row[1],
    items: row.slice(2).filter(item => item !== '')
  }));
  console.log(structuredData);
  return structuredData;
}

function saveSelections(category, item) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const cell = sheet.getActiveCell();

  sheet.getRange(cell.getRow(), 4).setValue(category);
  sheet.getRange(cell.getRow(), 5).setValue(item);
}

//都道府県市区
function showDialog() {
  const html = HtmlService.createHtmlOutputFromFile('Page')
    .setWidth(400)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, '都道府県、市、区を選択');
}

// スプレッドシートから階層的なデータを取得する
function getStructuredData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("都道府県市区");
  const protection = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET)[0];

  const range = sheet.getDataRange();
  const values = range.getValues();
  let structuredData = [];

  // ヘッダー行をスキップ
  values.slice(1).forEach((row, index) => {
    const [no, prefecture, city, district] = row;

    // 都道府県レベルでデータを追加
    let prefectureObj = structuredData.find(p => p.name === prefecture);
    if (!prefectureObj) {
      prefectureObj = { name: prefecture, cities: [] };
      structuredData.push(prefectureObj);
    }

    // 市レベルでデータを追加
    if (city) {
      let cityObj = prefectureObj.cities.find(c => c.name === city);
      if (!cityObj) {
        cityObj = { name: city, districts: [] };
        prefectureObj.cities.push(cityObj);
      }

      // 区レベルでデータを追加
      if (district && !cityObj.districts.includes(district)) {
        cityObj.districts.push(district);
      }
    }
  });

  // 再保護
  console.log(structuredData);
  return structuredData;
}

// 選択された都道府県、市、区をスプレッドシートに保存するApps Script関数
function saveSelectionsToSheet(selectedPrefectures, selectedCities, selectedDistricts) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // 配列をカンマ区切りの文字列に変換
  const prefecturesString = selectedPrefectures.join(", ");
  const citiesString = selectedCities.join(", ");
  const districtsString = selectedDistricts.join(", ");
  var row = sheet.getActiveCell().getRow();
  sheet.getRange(row, 6).setValue(prefecturesString); // 都道府県
  sheet.getRange(row, 7).setValue(citiesString);      // 市
  sheet.getRange(row, 8).setValue(districtsString);  // 区
  SpreadsheetApp.flush(); // 変更を即時反映させる
}
