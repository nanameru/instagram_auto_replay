// Instagram API関連の定数
const INSTAGRAM_API_BASE_URL = 'https://graph.instagram.com';
const SCRIPT_PROPERTIES = PropertiesService.getScriptProperties();

/**
 * アクセスキーを保存
 */
function saveAccessKey(accessKey) {
  try {
    SCRIPT_PROPERTIES.setProperty('INSTAGRAM_ACCESS_TOKEN', accessKey);
    return true;
  } catch (error) {
    console.error('アクセスキー保存エラー:', error);
    return false;
  }
}

/**
 * アクセスキー情報を取得
 */
function getAccessKeyInfo() {
  try {
    const accessKey = SCRIPT_PROPERTIES.getProperty('INSTAGRAM_ACCESS_TOKEN');
    if (!accessKey) return "未設定";
    
    // アクセストークンの有効性を確認
    const response = validateAccessToken(accessKey);
    if (!response.valid) {
      return "無効なトークン";
    }
    
    return `設定済み (末尾4桁: ${accessKey.slice(-4)})`;
  } catch (error) {
    console.error('アクセスキー情報取得エラー:', error);
    return "取得エラー";
  }
}

/**
 * アクセストークンの有効性を確認
 */
function validateAccessToken(accessToken) {
  try {
    const response = UrlFetchApp.fetch(`${INSTAGRAM_API_BASE_URL}/me?access_token=${accessToken}`);
    const data = JSON.parse(response.getContentText());
    return {
      valid: !data.error,
      userId: data.id
    };
  } catch (error) {
    console.error('トークン検証エラー:', error);
    return { valid: false };
  }
}

/**
 * DMを同期
 */
function syncDmsFromApi(customerId) {
  try {
    const accessToken = SCRIPT_PROPERTIES.getProperty('INSTAGRAM_ACCESS_TOKEN');
    if (!accessToken) {
      throw new Error('アクセストークンが設定されていません');
    }

    // アクセストークンの有効性を確認
    const validation = validateAccessToken(accessToken);
    if (!validation.valid) {
      throw new Error('無効なアクセストークン');
    }

    // 会話を取得
    const conversations = fetchConversations(accessToken);
    if (!conversations || !conversations.data) {
      throw new Error('会話の取得に失敗しました');
    }

    // スプレッドシートに保存されている顧客情報を取得
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Customers');
    if (!sheet) {
      throw new Error('Customersシートが見つかりません');
    }

    // 顧客IDに対応する会話を検索
    const targetConversation = conversations.data.find(conv => {
      // ここで顧客IDとInstagramユーザーIDの紐付けを行う
      // 実際の実装では、スプレッドシートなどで管理している顧客情報と
      // InstagramのユーザーIDを紐付ける必要があります
      return true; // 仮の実装
    });

    if (!targetConversation) {
      throw new Error('指定された顧客との会話が見つかりません');
    }

    // 会話の詳細（メッセージ）を取得
    const messages = fetchMessages(accessToken, targetConversation.id);
    if (!messages || !messages.data) {
      throw new Error('メッセージの取得に失敗しました');
    }

    // メッセージをスプレッドシートに保存
    saveMessagesToSheet(customerId, messages.data);

    return true;
  } catch (error) {
    console.error('DM同期エラー:', error);
    return false;
  }
}

/**
 * 会話一覧を取得
 */
function fetchConversations(accessToken) {
  try {
    const response = UrlFetchApp.fetch(`${INSTAGRAM_API_BASE_URL}/me/conversations?access_token=${accessToken}`);
    return JSON.parse(response.getContentText());
  } catch (error) {
    console.error('会話一覧取得エラー:', error);
    return null;
  }
}

/**
 * 特定の会話のメッセージを取得
 */
function fetchMessages(accessToken, conversationId) {
  try {
    const response = UrlFetchApp.fetch(
      `${INSTAGRAM_API_BASE_URL}/${conversationId}/messages?access_token=${accessToken}`
    );
    return JSON.parse(response.getContentText());
  } catch (error) {
    console.error('メッセージ取得エラー:', error);
    return null;
  }
}

/**
 * メッセージをスプレッドシートに保存
 */
function saveMessagesToSheet(customerId, messages) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Messages');
  if (!sheet) {
    throw new Error('Messagesシートが見つかりません');
  }

  // 既存のメッセージを取得（重複を避けるため）
  const existingData = sheet.getDataRange().getValues();
  const existingMessageIds = new Set(existingData.map(row => row[0]));

  // 新しいメッセージのみを追加
  const newMessages = messages.filter(msg => !existingMessageIds.has(msg.id));
  if (newMessages.length === 0) return;

  // メッセージを整形してシートに追加
  const newRows = newMessages.map(msg => [
    msg.id,
    customerId,
    msg.from.id === validation.userId ? 'user' : 'customer',
    msg.text,
    new Date(msg.timestamp)
  ]);

  sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, newRows[0].length)
    .setValues(newRows);
} 