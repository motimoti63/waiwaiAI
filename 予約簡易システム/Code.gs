function doGet() {
  return HtmlService.createHtmlOutputFromFile('index.html');
}

// 1週間分の空き時間を取得
function searchAvailability() {
  const calendarId = "vjnkmjmwom0603@gmail.com"; // カレンダーIDを直接ハードコード
  Logger.log("Using calendarId: " + calendarId); // カレンダーIDをログに出力

  if (!calendarId) {
    throw new Error("カレンダーIDが渡されていません");
  }

  const now = new Date();
  const oneWeekLater = new Date();
  oneWeekLater.setDate(now.getDate() + 7);

  Logger.log("Fetching events from: " + now + " to: " + oneWeekLater); // 範囲ログを出力

  const events = CalendarApp.getCalendarById(calendarId).getEvents(now, oneWeekLater);
  const weekdays = ["日", "月", "火", "水", "木", "金", "土"];
  const availability = [];

  for (let dayOffset = 0; dayOffset <= 7; dayOffset++) {
    const currentDate = new Date(now.getFullYear(), now.getMonth(), now.getDate() + dayOffset);
    Logger.log("Checking date: " + currentDate); // 日付ログ

    for (let hour = 9; hour < 18; hour++) {
      const start = new Date(currentDate.getFullYear(), currentDate.getMonth(), currentDate.getDate(), hour, 0);
      const end = new Date(currentDate.getFullYear(), currentDate.getMonth(), currentDate.getDate(), hour + 1, 0);

      if (start < now) continue;

      const isAvailable = !events.some(event => {
        const eventStart = event.getStartTime();
        const eventEnd = event.getEndTime();
        return (
          (eventStart <= start && eventEnd > start) ||
          (eventStart < end && eventEnd >= end) ||
          (eventStart >= start && eventEnd <= end)
        );
      });

      if (isAvailable) {
        const dayOfWeek = weekdays[start.getDay()];
        availability.push(`${start.toISOString().slice(0, 16).replace("T", " ")} (${dayOfWeek}) - ${end.toISOString().slice(0, 16).replace("T", " ")} (${dayOfWeek})`);
      }
    }
  }

  Logger.log("Availability found: " + JSON.stringify(availability)); // 空き時間をログに出力
  return availability;
}

// 予約を作成
function createReservation(data) {
  Logger.log("Received data for reservation: " + JSON.stringify(data)); // デバッグ用

  if (!data) {
    throw new Error("データが渡されていません");
  }

  const { userName, userEmail, selectedTime } = data;
  const calendarId = "vjnkmjmwom0603@gmail.com"; // カレンダーIDを直接ハードコード

  if (!userName || !userEmail || !selectedTime || !calendarId) {
    throw new Error("必要な情報が不足しています");
  }

  try {
    const [start, end] = selectedTime.split(" - ");
    const startDateTime = new Date(start).toISOString();
    const endDateTime = new Date(end).toISOString();

    // Google Calendar Advanced Service を使ってイベントを作成
    const event = Calendar.Events.insert(
      {
        summary: "カウンセリング予約",
        description: `予約者: ${userName}\nメール: ${userEmail}`,
        start: { dateTime: startDateTime },
        end: { dateTime: endDateTime },
        attendees: [{ email: userEmail }],
        conferenceData: {
          createRequest: {
            requestId: `meet-${Date.now()}`, // 一意のリクエストID
            conferenceSolutionKey: { type: "hangoutsMeet" },
          },
        },
      },
      calendarId,
      { conferenceDataVersion: 1 }
    );

    const meetLink = event.conferenceData.entryPoints.find(entry => entry.entryPointType === "video").uri;

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Reservations");
    if (!sheet) {
      throw new Error("Reservations シートが見つかりません");
    }

    sheet.appendRow([
      userName,
      userEmail,
      selectedTime,
      calendarId,
      meetLink,
      new Date(),
    ]);

    Logger.log("Reservation saved successfully with Meet link: " + meetLink); // 成功ログ
    return { meetLink };

  } catch (error) {
    Logger.log("Error creating reservation: " + error.message); // エラーログ
    return { error: error.message };
  }
}



// デバッグ用関数 (フロントエンドのデータを直接送信する代わりにテスト可能)
function testCreateReservation() {
  const testData = {
    userName: "山田 太郎",
    userEmail: "taro@example.com",
    selectedTime: "2025-01-30 10:00 - 2025-01-30 11:00"
  };
  const result = createReservation(testData);
  Logger.log(result); // テスト結果をログに出力
}
