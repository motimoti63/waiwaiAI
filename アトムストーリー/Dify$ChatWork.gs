// 定数定義
const DIFY_API_URL = 'https://api.dify.ai/v1/chat-messages';
const CHATWORK_API_URL = 'https://api.chatwork.com/v2';

// テスト用関数
function testConnection() {
  Logger.log('Testing connections...');
  const scriptProperties = PropertiesService.getScriptProperties();
  
  // APIキー取得
  const DIFY_API_KEY = scriptProperties.getProperty('DIFY_API_KEY');
  const CHATWORK_API_TOKEN = scriptProperties.getProperty('CHATWORK_API_TOKEN');
  const BOT_ACCOUNT_ID = scriptProperties.getProperty('BOT_ACCOUNT_ID');
  
  Logger.log('Checking properties:');
  Logger.log('BOT_ACCOUNT_ID exists: ' + !!BOT_ACCOUNT_ID);
  Logger.log('CHATWORK_API_TOKEN exists: ' + !!CHATWORK_API_TOKEN);
  Logger.log('DIFY_API_KEY exists: ' + !!DIFY_API_KEY);
  
  // Difyテスト
  try {
    const testResponse = callDifyAPI('テストメッセージ', DIFY_API_KEY);
    Logger.log('Dify test response: ' + testResponse);
  } catch (error) {
    Logger.log('Dify test failed: ' + error);
  }
  
  // Chatworkテスト
  try {
    const response = UrlFetchApp.fetch(`${CHATWORK_API_URL}/me`, {
      headers: {
        'X-ChatWorkToken': CHATWORK_API_TOKEN
      }
    });
    Logger.log('Chatwork test successful: ' + response.getContentText());
  } catch (error) {
    Logger.log('Chatwork test failed: ' + error);
  }
}

function doPost(e) {
  Logger.log('Webhook received');
  Logger.log('Payload: ' + JSON.stringify(e.postData.contents));
  
  const scriptProperties = PropertiesService.getScriptProperties();
  const DIFY_API_KEY = scriptProperties.getProperty('DIFY_API_KEY');
  const CHATWORK_API_TOKEN = scriptProperties.getProperty('CHATWORK_API_TOKEN');
  const BOT_ACCOUNT_ID = scriptProperties.getProperty('BOT_ACCOUNT_ID');
  
  try {
    const payload = JSON.parse(e.postData.contents);
    
    // メッセージ情報を取得
    const messageBody = payload.webhook_event.body;
    const roomId = payload.webhook_event.room_id;
    const senderId = payload.webhook_event.account_id;
    
    Logger.log('Message received: ' + messageBody);
    Logger.log('Room ID: ' + roomId);
    Logger.log('Sender ID: ' + senderId);
    Logger.log('BOT_ACCOUNT_ID: ' + BOT_ACCOUNT_ID);
    
    // ボットからのメッセージは無視
    if (senderId.toString() === BOT_ACCOUNT_ID.toString()) {
      Logger.log('Ignoring bot message');
      return ContentService.createTextOutput('OK');
    }
    
    // メンションパターンの確認
    const mentionPattern = `[To:${BOT_ACCOUNT_ID}]`;
    Logger.log('Checking mention pattern: ' + mentionPattern);
    Logger.log('Message contains mention: ' + messageBody.includes(mentionPattern));
    
    if (messageBody.includes(mentionPattern)) {
      Logger.log('Bot mention found');
      
      // メッセージからメンション部分を削除
      const cleanMessage = messageBody
        .replace(new RegExp(`\\[To:${BOT_ACCOUNT_ID}\\].*?(?=\\s|$)`), '')
        .trim();
      
      Logger.log('Clean message: ' + cleanMessage);
      
      if (!cleanMessage) {
        Logger.log('Empty message after cleaning');
        postToChatwork(roomId, 'メッセージ内容を入力してください。', CHATWORK_API_TOKEN);
        return ContentService.createTextOutput('OK');
      }
      
      // Dify APIに送信
      const difyResponse = callDifyAPI(cleanMessage, DIFY_API_KEY);
      Logger.log('Dify response: ' + difyResponse);
      
      if (difyResponse) {
        // Chatworkに返信
        postToChatwork(roomId, difyResponse, CHATWORK_API_TOKEN);
      }
    } else {
      Logger.log('No bot mention found');
    }
    
    return ContentService.createTextOutput('OK');
  } catch (error) {
    Logger.log('Error in doPost: ' + error.toString());
    Logger.log('Error stack: ' + error.stack);
    return ContentService.createTextOutput('Error occurred');
  }
}

// Dify APIにメッセージを送信
function callDifyAPI(message, apiKey) {
  const headers = {
    'Authorization': 'Bearer ' + apiKey,
    'Content-Type': 'application/json'
  };
  
  const payload = {
    'inputs': {},
    'query': message,
    'response_mode': 'blocking',
    'conversation_id': '',
    'user': 'chatwork-user-' + Math.random().toString(36).substr(2, 9)
  };
  
  const options = {
    'method': 'post',
    'headers': headers,
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true
  };
  
  try {
    Logger.log('Calling Dify API...');
    const response = UrlFetchApp.fetch(DIFY_API_URL, options);
    const responseText = response.getContentText();
    Logger.log('Dify response: ' + responseText);
    
    if (response.getResponseCode() === 200) {
      const responseBody = JSON.parse(responseText);
      return responseBody.answer || 'No response from Dify';
    }
    return 'Error calling Dify API';
  } catch (error) {
    Logger.log('Dify API error: ' + error);
    return 'API Error occurred';
  }
}

// Chatworkにメッセージを送信
function postToChatwork(roomId, message, apiToken) {
  const url = `${CHATWORK_API_URL}/rooms/${roomId}/messages`;
  
  const options = {
    'method': 'post',
    'headers': {
      'X-ChatWorkToken': apiToken
    },
    'payload': {
      'body': message
    }
  };
  
  try {
    Logger.log('Sending message to Chatwork...');
    const response = UrlFetchApp.fetch(url, options);
    Logger.log('Chatwork response: ' + response.getContentText());
    return response;
  } catch (error) {
    Logger.log('Chatwork API error: ' + error);
    throw error;
  }
}