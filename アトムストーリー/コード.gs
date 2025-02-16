/**********************************
 *  1. グローバル変数/プロパティ
 **********************************/
const properties = PropertiesService.getScriptProperties();

/**********************************
 *  2. 認可URLを生成してユーザーに提示
 *     → ユーザーがブラウザで開き、Chatworkログイン＆許可
 **********************************/
function getAuthorizationCode() {
  // (1) スクリプトプロパティから各種情報を取得
  const redirectUri = properties.getProperty('REDIRECT_URI');  // 例: https://script.google.com/macros/s/XXXXX/exec
  const clientId    = properties.getProperty('CLIENT_ID');     // Chatwork管理画面で発行したクライアントID
  const scope       = properties.getProperty('SCOPE');         // 例: rooms.messages:read

  // (2) stateを生成し、CSRF対策のためプロパティに保存
  const state = generateStateParameter();
  properties.setProperty('STATE', state);

  // (3) code_verifierを生成し、プロパティに保存 (PKCEのため)
  const codeVerifier = generateCodeVerifier(43);
  properties.setProperty('CODE_VERIFIER', codeVerifier);

  // (4) code_challengeを生成
  const codeChallenge = generateCodeChallenge(codeVerifier);

  // (5) 認可リクエストURLを構築
  const authorizationUrl =
    'https://www.chatwork.com/packages/oauth2/login.php' +
    '?response_type=code' +
    '&redirect_uri=' + encodeURIComponent(redirectUri) +
    '&client_id='     + encodeURIComponent(clientId) +
    '&state='         + encodeURIComponent(state) +
    '&scope='         + encodeURIComponent(scope) +
    '&code_challenge='         + encodeURIComponent(codeChallenge) +
    '&code_challenge_method=S256';

  // (6) GAS上でダイアログ表示 → ユーザーがこのURLを開く
  SpreadsheetApp.getUi().alert(
    "以下のURLをブラウザで開き、Chatworkアカウントで認可を許可してください:\n\n" + authorizationUrl
  );
}

/**********************************
 *  3.  ブラウザで認可を許可後、
 *      Chatwork側からリダイレクトされる先の処理
 *      → ここで認可コードを受け取り、トークン発行
 **********************************/
function doGet(e) {
  // Chatworkからリダイレクトされた時のパラメータ
  const authorizationCode = e.parameter.code;
  const state = e.parameter.state;

  // 1) CSRF対策: リクエスト時のstateと一致するかチェック
  if (state !== PropertiesService.getScriptProperties().getProperty('STATE')) {
    return HtmlService.createHtmlOutput('state が一致しません。');
  }

  // 2) 認可コードが正しく取得できていれば、トークンを発行してスクリプトプロパティに保存
  if (authorizationCode) {
    const codeVerifier = PropertiesService.getScriptProperties().getProperty('CODE_VERIFIER');
    const response = getTokens(authorizationCode, codeVerifier);

    PropertiesService.getScriptProperties().setProperty('ACCESS_TOKEN', response.access_token);
    PropertiesService.getScriptProperties().setProperty('REFRESH_TOKEN', response.refresh_token);

    return HtmlService.createHtmlOutput('認証に成功しました。このページは閉じてOKです。');

  } else {
    return HtmlService.createHtmlOutput('認可コードが取得できませんでした。');
  }
}


/**********************************
 *  4. 初回(認可コード)でトークン取得
 **********************************/
function getTokens(authorizationCode, codeVerifier) {
  const clientId    = PropertiesService.getScriptProperties().getProperty('CLIENT_ID');
  const redirectUri = PropertiesService.getScriptProperties().getProperty('REDIRECT_URI');
  const tokenEndpoint = 'https://oauth.chatwork.com/token';

  const params = {
    'grant_type': 'authorization_code',
    'code': authorizationCode,
    'client_id': clientId,
    'redirect_uri': redirectUri,
    'code_verifier': codeVerifier
  };

  const options = {
    method: 'post',
    payload: params,
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
  };

  try {
    const response = UrlFetchApp.fetch(tokenEndpoint, options);
    return JSON.parse(response.getContentText());  // {access_token, refresh_token, ...} が返る
  } catch (error) {
    throw new Error('アクセストークン取得エラー:' + error);
  }
}


/**********************************
 *  5. リフレッシュトークンで
 *     新たにアクセストークンを再発行する関数
 **********************************/
function updateTokens() {
  const clientId     = properties.getProperty('CLIENT_ID');
  const refreshToken = properties.getProperty('REFRESH_TOKEN');
  const scope        = properties.getProperty('SCOPE'); // 初回と同じスコープ
  const tokenEndpoint = 'https://oauth.chatwork.com/token';

  const params = {
    'grant_type': 'refresh_token',
    'refresh_token': refreshToken,
    'client_id': clientId,
    'scope': scope
  };

  const options = {
    method: 'post',
    payload: params,
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
  };

  try {
    const response = UrlFetchApp.fetch(tokenEndpoint, options);
    const data = JSON.parse(response.getContentText());

    // 新しいトークンを上書き保存
    properties.setProperty('ACCESS_TOKEN', data.access_token);
    properties.setProperty('REFRESH_TOKEN', data.refresh_token);
  } catch (error) {
    Logger.log('トークン更新エラー:' + error);
    throw new Error('トークン更新に失敗');
  }
}

/**********************************
 *  6. (例) Chatwork API で
 *     特定のルームのメッセージを取得し、
 *     スプレッドシートに転記する処理
 **********************************/
function getChatworkMessages() {
  // 1) まずトークンをリフレッシュして有効にする
  updateTokens();

  // 2) リクエスト準備
  const accessToken = properties.getProperty('ACCESS_TOKEN');
  const roomId = properties.getProperty('ROOM_ID'); // ex: 123456789
  const requestUrl = `https://api.chatwork.com/v2/rooms/${roomId}/messages?force=1`;

  const options = {
    method: 'get',
    headers: {
      accept: 'application/json',
      authorization: 'Bearer ' + accessToken
    }
  };

  // 3) API呼び出し
  try {
    const response = UrlFetchApp.fetch(requestUrl, options);
    handleResponse(response);
  } catch (err) {
    Logger.log('Chatwork API 呼び出し失敗: ' + err);
  }
}

/**********************************
 *  7. Chatwork API のレスポンスをハンドリング
 **********************************/
function handleResponse(response) {
  const code = response.getResponseCode();
  const data = JSON.parse(response.getContentText());

  switch (code) {
    case 200:
      parseAndOutputMessages(data);
      break;
    case 204:
      Logger.log('メッセージがありません。');
      break;
    case 401:
      Logger.log('認証エラー: トークンが無効かもしれません。');
      break;
    case 403:
      Logger.log('権限エラー: スコープやルームIDの問題など。');
      break;
    case 429:
      Logger.log('レート制限超過: しばらく待って再リクエスト。');
      break;
    default:
      Logger.log('想定外のエラー: ' + code + ' / ' + response.getContentText());
      break;
  }
}

/**********************************
 *  8. 取得したメッセージを
 *     スプレッドシートに転記する例
 **********************************/
function parseAndOutputMessages(messages) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('メッセージ履歴');
  messages.forEach(msg => {
    const updateTime = msg.update_time 
      ? new Date(msg.update_time * 1000)
      : '-';
    const rowData = [
      msg.message_id,
      new Date(msg.send_time * 1000),
      updateTime,
      msg.account.name,
      msg.body
    ];
    sheet.appendRow(rowData);
  });
}

/**********************************
 *  補助関数: ランダム文字列
 **********************************/
function generateCodeVerifier(length) {
  const charset = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-._~';
  let result = '';
  for (let i = 0; i < length; i++) {
    result += charset.charAt(Math.floor(Math.random() * charset.length));
  }
  return result;
}

/**********************************
 * 補助関数: SHA-256 ハッシュ → Base64URL
 **********************************/
function generateCodeChallenge(codeVerifier) {
  const hashBuffer = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, codeVerifier);
  // パディングを除去 (=を削除)
  return Utilities.base64EncodeWebSafe(hashBuffer).replace(/=+$/, '');
}

/**********************************
 * 補助関数: stateパラメータ生成
 **********************************/
function generateStateParameter() {
  return Math.random().toString(36).substring(2, 18) +
         Math.random().toString(36).substring(2, 18);
}





/**
 * Chatworkのトークンをリフレッシュしてから、メッセージを取得し、シートに書き込む例
 */
function getChatworkMessages() {
  // アクセストークンをリフレッシュ
  updateTokens();

  // プロパティから取得
  const accessToken = PropertiesService.getScriptProperties().getProperty('ACCESS_TOKEN');
  const roomId      = PropertiesService.getScriptProperties().getProperty('ROOM_ID'); // 取得したいルームID
  // Chatwork APIエンドポイント（v2版）
  const url = `https://api.chatwork.com/v2/rooms/${roomId}/messages?force=1`;

  const options = {
    method: 'GET',
    headers: {
      "Authorization": "Bearer " + accessToken,
      "Accept": "application/json"
    }
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const statusCode = response.getResponseCode();
    const data = JSON.parse(response.getContentText());

    if (statusCode === 200) {
      parseAndOutputMessages(data);
    } else if (statusCode === 204) {
      Logger.log("メッセージがありません。");
    } else {
      Logger.log("エラー: " + statusCode + " / " + response.getContentText());
    }
  } catch (err) {
    Logger.log("リクエスト失敗: " + err);
  }
}

/**
 * 取得したメッセージ配列をスプレッドシートに書き込むサンプル
 */
function parseAndOutputMessages(messages) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('メッセージ履歴');
  messages.forEach(msg => {
    const updateTime = msg.update_time === 0 ? '-' : new Date(msg.update_time * 1000);
    const rowData = [
      msg.message_id,
      new Date(msg.send_time * 1000),
      updateTime,
      msg.account.name,
      msg.body
    ];
    sheet.appendRow(rowData);
  });
}

