// Google Apps Script エントリーポイント
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('インスタグラム管理ダッシュボード')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// スプレッドシートからデータを取得する関数
function getSpreadsheetData() {
  // アクティブなスプレッドシートを取得
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  
  // データ範囲を取得
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  
  // ヘッダー行をスキップ
  const headers = values[0];
  const data = values.slice(1);
  
  return {
    headers: headers,
    data: data
  };
}

// 顧客データを構築する関数
function buildCustomersData() {
  const sheetData = getSpreadsheetData();
  const data = sheetData.data;
  
  // 顧客IDごとにデータをグループ化
  const customers = {};
  
  // 顧客名、メッセージ、顧客IDのインデックスを取得
  const nameIndex = sheetData.headers.indexOf('顧客名');
  const messageIndex = sheetData.headers.indexOf('メッセージ');
  const idIndex = sheetData.headers.indexOf('顧客ID');
  
  // データが存在する場合のみ処理
  if (nameIndex >= 0 && messageIndex >= 0 && idIndex >= 0) {
    // 各行を処理
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const customerId = row[idIndex].toString();
      const customerName = row[nameIndex];
      const message = row[messageIndex];
      
      // 空行はスキップ
      if (!customerId || !message) continue;
      
      // 顧客IDが存在しない場合は新しく作成
      if (!customers[customerId]) {
        customers[customerId] = {
          id: customerId,
          name: customerName || 'Unknown',
          username: '@' + (customerName || 'user') + '_' + customerId,
          avatar: 'https://randomuser.me/api/portraits/' + 
                 (Math.random() > 0.5 ? 'men' : 'women') + '/' + 
                 Math.floor(Math.random() * 100) + '.jpg',
          messages: []
        };
      }
      
      // メッセージを追加（空でない場合のみ）
      if (message) {
        const timestamp = new Date().toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'});
        customers[customerId].messages.push({
          sender: 'customer',
          text: message,
          timestamp: timestamp
        });
      }
    }
  }
  
  return customers;
}

// コンテキストに基づいた返信候補を生成する関数
function generateReplySuggestions(customerId, lastMessage) {
    // 実際のアプリケーションではAIや条件分岐などでより適切な候補を生成する
    // ここでは簡易的に特定のキーワードに基づいて候補を返す
    const message = lastMessage.toLowerCase();
    
    if (message.includes('質問') || message.includes('教えて')) {
        return [
            'ご質問ありがとうございます。どのようなことでしょうか？',
            '詳細を教えていただけますと幸いです。',
            'お力になれるよう努めます。もう少し詳しく教えていただけますか？'
        ];
    } else if (message.includes('注文') || message.includes('配送') || message.includes('届け')) {
        return [
            'ご注文の詳細を確認させていただきます。注文番号をお知らせいただけますか？',
            '配送状況を確認いたします。少々お待ちください。',
            'ご注文の商品は現在発送準備中です。もうしばらくお待ちください。'
        ];
    } else if (message.includes('返品') || message.includes('交換')) {
        return [
            '返品・交換についてのお問い合わせですね。当店では購入後14日以内であれば対応可能です。',
            '商品の状態を確認させていただきたいです。お手数ですが写真を送っていただけますか？',
            '返品・交換フォームをお送りいたします。必要事項をご記入ください。'
        ];
    } else if (message.includes('ありがとう')) {
        return [
            'こちらこそありがとうございます。他にご質問があればいつでもどうぞ。',
            'お役に立てて嬉しいです。今後ともよろしくお願いいたします。',
            'またのお問い合わせをお待ちしております。'
        ];
    } else if (message.includes('住宅') || message.includes('不動産') || message.includes('物件')) {
        return [
            '物件についてのお問い合わせですね。具体的にどのような物件をお探しですか？',
            '当社では様々な物件を取り扱っております。ご予算や希望の地域はございますか？',
            '物件の詳細資料をお送りすることも可能です。ご希望の条件をお知らせください。'
        ];
    } else if (message.includes('天気')) {
        return [
            'お天気についてのお話ですね。今日も良い一日をお過ごしください。',
            '天気の良い日は気分も上がりますね。何かお手伝いできることはありますか？',
            '素敵な一日をお過ごしください。他にご質問などありましたらお気軽にどうぞ。'
        ];
    } else {
        return [
            'ご連絡ありがとうございます。どのようなことでしょうか？',
            'いつもご利用ありがとうございます。お手伝いできることがあればお知らせください。',
            'ご不明点があれば、お気軽にお問い合わせください。'
        ];
    }
}

// サーバーサイド関数：顧客データを取得
function getCustomers() {
  return buildCustomersData();
}

// サーバーサイド関数：特定の顧客データを取得
function getCustomer(customerId) {
  const customers = buildCustomersData();
  return customers[customerId];
}

// サーバーサイド関数：メッセージを送信（データ更新）
function sendMessageToCustomer(customerId, text) {
  try {
    // アクティブなスプレッドシートを取得
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    
    // 最終行を取得
    const lastRow = sheet.getLastRow();
    
    // ヘッダー行を取得
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // 各列のインデックスを取得
    const nameIndex = headers.indexOf('顧客名') + 1;
    const messageIndex = headers.indexOf('メッセージ') + 1;
    const idIndex = headers.indexOf('顧客ID') + 1;
    
    // 顧客情報を取得
    const customers = buildCustomersData();
    const customer = customers[customerId];
    
    if (customer && nameIndex > 0 && messageIndex > 0 && idIndex > 0) {
      // 新しい行にデータを追加
      sheet.getRange(lastRow + 1, nameIndex).setValue(customer.name);
      sheet.getRange(lastRow + 1, messageIndex).setValue(text);
      sheet.getRange(lastRow + 1, idIndex).setValue(customerId);
      
      return true;
    }
    
    return false;
  } catch (error) {
    Logger.log('Error in sendMessageToCustomer: ' + error.toString());
    return false;
  }
}

// サーバーサイド関数：返信候補を生成
function getReplySuggestions(customerId) {
  const customers = buildCustomersData();
  const customer = customers[customerId];
  
  if (customer && customer.messages.length > 0) {
    const lastMessage = customer.messages[customer.messages.length - 1];
    return generateReplySuggestions(customerId, lastMessage.text);
  }
  
  return [];
}

// サーバーサイド関数：メッセージを編集
function editMessageInSheet(customerId, oldText, newText) {
  try {
    // アクティブなスプレッドシートを取得
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    
    // データ範囲を取得
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    // ヘッダー行を取得
    const headers = values[0];
    
    // 各列のインデックスを取得
    const messageIndex = headers.indexOf('メッセージ');
    const idIndex = headers.indexOf('顧客ID');
    
    // データが存在する場合のみ処理
    if (messageIndex >= 0 && idIndex >= 0) {
      // 一致する行を検索
      for (let i = 1; i < values.length; i++) {
        const row = values[i];
        const rowCustomerId = row[idIndex].toString();
        const rowMessage = row[messageIndex];
        
        // 顧客IDとメッセージが一致する行を見つけた場合
        if (rowCustomerId === customerId && rowMessage === oldText) {
          // メッセージを更新
          sheet.getRange(i + 1, messageIndex + 1).setValue(newText);
          return true;
        }
      }
    }
    
    return false;
  } catch (error) {
    Logger.log('Error in editMessageInSheet: ' + error.toString());
    return false;
  }
}