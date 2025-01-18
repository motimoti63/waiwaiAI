function createJsonFromRow6() {
  // シート名を指定してシートを取得
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("買い情報赤川");
  
  // カラムDからZまでのヘッダーを配列として定義
  var headers = [
    "種別 プルダウン",
    "ジャンル プルダウン",
    "都道府県 プルダウン（1個）",
    "市",
    "場所 選択式（複数可）",
    "駅",
    "分以内",
    "金額以上 (万円)",
    "金額以下 (万円)",
    "一種 単価 以下",
    "坪 単価 以下",
    "坪以上 (土地)",
    "坪以下 (土地)",
    "延床面積 坪以上",
    "容積 以上",
    "容積 以下",
    "前面道路 幅員以上 (広い方)",
    "間口以上 (広い方)",
    "表面 利回り 以上",
    "築年数 以内",
    "検査 済証 要・不要",
    "全空",
    "その他"
  ];
  
  // ヘッダーの数がカラムDからZの数と一致しているか確認
  var expectedColumnCount = 23; // カラムDからZは23列
  if (headers.length !== expectedColumnCount) {
    Logger.log("ヘッダーの数がカラムDからZの数と一致していません。");
    return;
  }
  
  // 6行目のデータを取得（カラムD=4から開始、23列取得）
  var dataRange = sheet.getRange(6, 4, 1, expectedColumnCount);
  var data = dataRange.getValues()[0];
  
  // ヘッダーとデータをマッピングしてJSONオブジェクトを作成
  var jsonObject = {};
  for (var i = 0; i < headers.length; i++) {
    jsonObject[headers[i]] = data[i];
  }
  
  // JSONオブジェクトを文字列に変換
  var jsonString = JSON.stringify(jsonObject, null, 2); // 見やすいようにインデントを追加
  
  // ログに出力（必要に応じて他の処理に利用可能）
  Logger.log(jsonString);
  
  // 必要に応じてJSONを返す
  return jsonObject;
}



function createSellInfoJson() {
  // シート名を指定してシートを取得
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("売り情報赤川");
  
  if (!sheet) {
    Logger.log("シート「売り情報赤川」が見つかりません。");
    return;
  }
  
  // カラムAからXまでのヘッダーを配列として定義
  var headers = [
    "会社",
    "先数 コード",
    "名前",
    "種別 プルダウン",
    "ジャンル プルダウン",
    "都道府県 プルダウン（1個）",
    "市 プルダウン（1個）",
    "場所 プルダウン（1個）",
    "駅",
    "分",
    "金額 (万円)",
    "一種 単価",
    "坪 単価",
    "㎡ (土地)",
    "坪 (土地)",
    "容積",
    "用途地域",
    "前面道路 幅員(広い方)",
    "間口 (広い方)",
    "表面 利回り or粗利",
    "築年数",
    "検査 済証",
    "全空",
    "その他"
  ];
  
  // ヘッダーの数がカラムAからXの数と一致しているか確認
  var expectedColumnCount = 24; // カラムAからXは24列
  if (headers.length !== expectedColumnCount) {
    Logger.log("ヘッダーの数がカラムAからXの数と一致していません。");
    return;
  }
  
  // 対象となる行番号を配列で定義
  var targetRows = [121, 136, 153];
  
  // JSONオブジェクトを格納する配列
  var jsonArray = [];
  
  // 各対象行に対して処理を実行
  for (var i = 0; i < targetRows.length; i++) {
    var rowNumber = targetRows[i];
    
    // 行番号がシートの範囲内か確認
    if (rowNumber > sheet.getLastRow()) {
      Logger.log("行番号 " + rowNumber + " はシートの最終行を超えています。");
      continue;
    }
    
    // 対象行のデータを取得（カラムA=1から開始、24列取得）
    var dataRange = sheet.getRange(rowNumber, 1, 1, expectedColumnCount);
    var data = dataRange.getValues()[0];
    
    // ヘッダーとデータをマッピングしてJSONオブジェクトを作成
    var jsonObject = {};
    for (var j = 0; j < headers.length; j++) {
      jsonObject[headers[j]] = data[j];
    }
    
    // 作成したJSONオブジェクトを配列に追加
    jsonArray.push(jsonObject);
  }
  
  // JSON配列を文字列に変換（見やすいようにインデントを追加）
  var jsonString = JSON.stringify(jsonArray, null, 2);
  
  // ログに出力（必要に応じて他の処理に利用可能）
  Logger.log(jsonString);
  
  // 必要に応じてJSONを返す
  return jsonArray;
}

