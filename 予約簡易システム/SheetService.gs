/**
 * Reservations シートにデータを保存
 * @param {Object} data - 保存するデータ
 */
function saveToSheet(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Reservations");

  if (!sheet) {
    throw new Error("Reservations シートが見つかりません");
  }

  // データをシートに追加
  sheet.appendRow([
    data.userName,        // 予約者名
    data.userEmail,       // 予約者のメール
    data.selectedTime,    // 予約された時間
    data.calendarId,      // カレンダーID
    data.meetLink,        // Google Meetリンク
    new Date(),           // 作成日時
  ]);
}

/**
 * Counselors シートからカウンセラー情報を取得
 * @returns {Array} - カウンセラー情報の配列
 */
function getCounselors() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Counselors");

  if (!sheet) {
    throw new Error("Counselors シートが見つかりません");
  }

  const rows = sheet.getDataRange().getValues();

  // カウンセラー情報をオブジェクト形式で返す
  return rows.slice(1).map(row => ({
    id: row[0],          // counselor_id
    name: row[1],        // counselor_name
    calendarId: row[2],  // calendar_id
    isActive: row[3],    // is_active
  })).filter(counselor => counselor.isActive); // is_active が TRUE のみ返す
}
