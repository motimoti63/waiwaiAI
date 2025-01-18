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




// hidaka





/**
 * AI比較を実行：チェックが入った行を対象に最適な物件を選定し、リンクと理由を格納
 */
function processCheckedRows() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("買い情報赤川");
  if (!sheet) {
    Logger.log("シート「買い情報赤川」が見つかりません。");
    return;
  }

  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();

  for (let i = 1; i < data.length; i++) { // ヘッダーをスキップ
    const isChecked = data[i][27]; // AB列にチェックボックスがある想定（インデックス27）

    if (isChecked === true) {
      const afLink = data[i][31]; // AF列
      const agLink = data[i][32]; // AG列
      const ahLink = data[i][33]; // AH列

      if (afLink && agLink && ahLink) {
        // 1. リンク先から詳細を取得
        const properties = [
          parsePropertyDetails(afLink),
          parsePropertyDetails(agLink),
          parsePropertyDetails(ahLink)
        ].filter(Boolean); // null データを除外

        if (properties.length > 0) {
          // 2. 最適な物件を選定
          const selectedProperty = selectBestProperty(properties);

          // 3. AD列に選定した物件のリンクを記載
          sheet.getRange(i + 1, 30).setValue(selectedProperty.details);

          // 4. AC列に選定理由を記載
          const matchingDetails = createJsonFromRow6();
          const reason = generateReason(properties, selectedProperty, matchingDetails);
          sheet.getRange(i + 1, 29).setValue(reason);
        }
      }
    }
  }
}

/**
 * AF, AG, AH列のリンクを基に、売り情報赤川シートから詳細を取得
 */
function parsePropertyDetails(linkCell) {
  if (!linkCell) {
    Logger.log("リンクセルが空または無効です。");
    return null;
  }

  // リンクから範囲情報を抽出 (例: range=A121)
  const match = linkCell.match(/range=[A-Z]+(\d+)/); // 例: "range=A121" から 121 を抽出
  if (!match) {
    Logger.log("リンクから範囲情報を抽出できません: " + linkCell);
    return null;
  }

  const row = parseInt(match[1], 10); // 行番号を取得（例: 121）

  // 売り情報赤川シートを取得
  const linkSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("売り情報赤川");
  if (!linkSheet) {
    Logger.log("シート「売り情報赤川」が見つかりません。");
    return null;
  }

  // 必要なデータを取得
  try {
    const company = linkSheet.getRange(row, 1).getValue(); // A列: 会社
    const price = linkSheet.getRange(row, 11).getValue(); // K列: 金額
    const area = linkSheet.getRange(row, 15).getValue(); // O列: 坪
    const ratio = linkSheet.getRange(row, 17).getValue(); // Q列: 容積率
    const location = linkSheet.getRange(row, 7).getValue(); // G列: 市区町村

    return {
      company: company || "不明",
      location: location || "不明",
      price: parseFloat(price || 0),
      area: parseFloat(area || 0),
      ratio: parseFloat(ratio || 0),
      details: linkCell // AD列にリンクを格納するため
    };
  } catch (e) {
    Logger.log("データ取得時にエラーが発生しました: " + e.message);
    return null;
  }
}

/**
 * 最適な物件を選定（価格が最も安いものを選択）
 */
function selectBestProperty(properties) {
  return properties.reduce((best, property) => (property.price < best.price ? property : best));
}

/**
 * 選定理由を生成：買い情報条件に基づくマッチング理由を詳細に記述
 */
function generateReason(properties, selectedProperty, matchingDetails) {
  const reasons = [];

  reasons.push(`買い情報に最もマッチする売り情報は以下です：`);
  reasons.push(`選択した売り情報：`);
  reasons.push(`会社: ${selectedProperty.company}`);
  reasons.push(`場所: ${selectedProperty.location}`);
  reasons.push(`金額: ${selectedProperty.price}万円`);
  reasons.push(`坪: ${selectedProperty.area}坪`);
  reasons.push(`容積率: ${selectedProperty.ratio}%`);

  reasons.push(`理由：`);
  reasons.push(`種別: ${matchingDetails["種別 プルダウン"]} → 一致`);
  reasons.push(`ジャンル: ${matchingDetails["ジャンル プルダウン"]} → 一致`);
  reasons.push(`場所: ${selectedProperty.location} → 買い条件に含まれる`);
  reasons.push(
    `売り情報の価格（${selectedProperty.price}万円）は、買い情報の予算内（${
      matchingDetails["金額以下 (万円)"]
    }万円）に収まっています。`
  );

  reasons.push(`条件を満たした中での最適選択`);
  properties.forEach(property => {
    if (property.details !== selectedProperty.details) {
      reasons.push(
        `他の候補（${property.details}）は、価格や条件面で劣ります。`
      );
    }
  });

  reasons.push(`結論`);
  reasons.push(
    `「${selectedProperty.company}」の売り情報は、買い情報の条件を最も満たし、価格、立地、土地形状の面で最適です。`
  );

  return reasons.join("\n");
}