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

// スプレッドシートから取り込んだDMデータを顧客リストとして表示する関数
function getImportedDMCustomers() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    
    // データ範囲を取得
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    // ヘッダー行を取得
    const headers = values[0];
    
    // 各列のインデックスを取得
    const idIndex = headers.indexOf('顧客ID');
    const nameIndex = headers.indexOf('顧客名');
    const messageIndex = headers.indexOf('メッセージ');
    const senderIndex = headers.indexOf('送信者');
    const timestampIndex = headers.indexOf('タイムスタンプ');
    
    // DMデータが存在するか確認
    if (idIndex < 0 || nameIndex < 0 || messageIndex < 0) {
      return { error: "必要なデータカラム（顧客ID、顧客名、メッセージ）が見つかりません。" };
    }
    
    // インポートされた顧客データを格納するオブジェクト
    const customers = {};
    
    // すべての行を処理
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      
      // 必要なデータを取得
      const customerId = row[idIndex]?.toString();
      const customerName = row[nameIndex];
      const message = row[messageIndex];
      const sender = senderIndex >= 0 ? row[senderIndex] : '';
      const timestamp = timestampIndex >= 0 ? row[timestampIndex] : new Date().toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'});
      
      // 空の行はスキップ
      if (!customerId || !customerName || !message) continue;
      
      // 顧客IDが存在しない場合は新しく作成
      if (!customers[customerId]) {
        customers[customerId] = {
          id: customerId,
          name: customerName,
          username: '@' + (customerName?.replace(/\s+/g, '_').toLowerCase() || 'user') + '_' + customerId.substring(0, 5),
          avatar: 'https://randomuser.me/api/portraits/' + 
                 (Math.random() > 0.5 ? 'men' : 'women') + '/' + 
                 Math.floor(Math.random() * 100) + '.jpg',
          messages: []
        };
      }
      
      // メッセージを追加
      if (message) {
        customers[customerId].messages.push({
          sender: sender === '自社' ? 'user' : 'customer',
          text: message,
          timestamp: timestamp
        });
      }
    }
    
    return customers;
  } catch (error) {
    Logger.log('Error in getImportedDMCustomers: ' + error.toString());
    return { error: error.toString() };
  }
}

// インポートしたファイルからDMデータを追加する（既存のgetCustomers関数を拡張）
function getCustomers() {
  // 基本の顧客データをビルド
  const baseCustomers = buildCustomersData();
  
  // インポートされたDMデータを取得
  const importedCustomers = getImportedDMCustomers();
  
  // エラーが発生した場合は基本データのみを返す
  if (importedCustomers.error) {
    Logger.log('Warning: ' + importedCustomers.error);
    return baseCustomers;
  }
  
  // 両方のデータをマージ
  const mergedCustomers = { ...baseCustomers };
  
  // インポートされた顧客データを追加
  for (const id in importedCustomers) {
    if (mergedCustomers[id]) {
      // 既存の顧客の場合、メッセージを追加
      mergedCustomers[id].messages = mergedCustomers[id].messages.concat(importedCustomers[id].messages);
    } else {
      // 新しい顧客の場合、そのまま追加
      mergedCustomers[id] = importedCustomers[id];
    }
  }
  
  return mergedCustomers;
}

// サーバーサイド関数：特定の顧客データを取得
function getCustomer(customerId) {
  const customers = getCustomers();
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
    const customers = getCustomers();
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
  const customers = getCustomers();
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

// OCRでスクリーンショットからテキストを抽出する関数
function processOCRImage(base64Image) {
  try {
    // Base64文字列をバイナリデータに変換
    const imageBlob = Utilities.newBlob(Utilities.base64Decode(base64Image), 'image/png', 'instagram-dm.png');
    
    // Google Drive APIを使ってOCR処理
    const resource = {
      title: 'instagram-dm-' + new Date().getTime() + '.png',
      mimeType: 'image/png'
    };
    
    // 画像をDriveにアップロード
    const file = DriveApp.createFile(imageBlob);
    
    // OCR処理のためにDocs APIを使用
    const docContent = DocumentApp.create('OCR-' + new Date().getTime());
    const docId = docContent.getId();
    const doc = DocumentApp.openById(docId);
    
    // Vision APIを使用するための疑似コード
    const ocrText = extractTextFromImage(file.getId());
    
    // 一時ファイルを削除
    DriveApp.getFileById(docId).setTrashed(true);
    file.setTrashed(true);
    
    // 抽出したテキストを処理
    return processDMContent(ocrText, 'OCR');
  } catch (error) {
    Logger.log('Error in processOCRImage: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

// Vision APIを使用してテキスト抽出
function extractTextFromImage(fileId) {
  try {
    // NOTE: このコードを動作させるには、Google Cloud Vision APIを有効化し、
    // Google Apps ScriptのAdvanced Servicesでも有効化する必要があります
    
    // 実際のコードの代わりに、テキスト抽出の結果をシミュレート
    const file = DriveApp.getFileById(fileId);
    const imageBlob = file.getBlob();
    
    // ここで実際にはVision APIを呼び出します
    // Vision.Images.annotate({requests: [{image: {content: imageBlob.getBytes()}, features: [{type: 'TEXT_DETECTION'}]}]});
    
    // シミュレーションのためのダミーテキスト
    const currentDate = new Date().toLocaleString();
    return "ユーザー: こんにちは、物件について質問があります\n" +
           "時間: " + currentDate + "\n" +
           "4tenno.ace: いつもありがとうございます。どのような物件をお探しですか？\n" +
           "時間: " + currentDate + "\n" +
           "ユーザー: 神戸市の3LDKで予算は3000万円以内を探しています\n" +
           "時間: " + currentDate;
  } catch (error) {
    Logger.log('Error in extractTextFromImage: ' + error.toString());
    return "OCR処理中にエラーが発生しました: " + error.toString();
  }
}

// HTMLからDMコンテンツを抽出する関数
function processHTMLContent(htmlContent) {
  try {
    // HTMLコンテンツをパース
    // 実際の実装では、正規表現やHTMLパーサーを使用します
    
    // シンプルな例：メッセージブロックを正規表現で抽出
    const messagePattern = /<div class="message[^>]*>(.*?)<\/div>/g;
    const timePattern = /<span class="timestamp[^>]*>(.*?)<\/span>/g;
    const userPattern = /<span class="username[^>]*>(.*?)<\/span>/g;
    
    // マッチするすべてのメッセージを抽出
    let messages = [];
    let match;
    let index = 0;
    
    // ユーザー名を抽出
    const usernames = [];
    while ((match = userPattern.exec(htmlContent)) !== null) {
      usernames.push(match[1]);
    }
    
    // タイムスタンプを抽出
    const timestamps = [];
    while ((match = timePattern.exec(htmlContent)) !== null) {
      timestamps.push(match[1]);
    }
    
    // メッセージを抽出
    while ((match = messagePattern.exec(htmlContent)) !== null) {
      const messageText = match[1].replace(/<[^>]*>/g, '').trim();
      const sender = index % 2 === 0 ? (usernames[0] || 'ユーザー') : '4tenno.ace';
      const timestamp = timestamps[index] || new Date().toLocaleString();
      
      messages.push({
        sender: sender,
        text: messageText,
        timestamp: timestamp
      });
      
      index++;
    }
    
    return processDMContent(messages, 'HTML');
  } catch (error) {
    Logger.log('Error in processHTMLContent: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

// DMコンテンツを処理してスプレッドシートに保存する関数
function processDMContent(content, source) {
  try {
    // スプレッドシートを取得
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    
    // ヘッダー行を取得
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // 各列のインデックスを取得
    const nameIndex = headers.indexOf('顧客名') + 1;
    const messageIndex = headers.indexOf('メッセージ') + 1;
    const idIndex = headers.indexOf('顧客ID') + 1;
    const timestampIndex = headers.indexOf('タイムスタンプ') + 1;
    const senderIndex = headers.indexOf('送信者') + 1;
    
    // タイムスタンプ列が存在しない場合は追加
    if (timestampIndex <= 0) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue('タイムスタンプ');
    }
    
    // 送信者列が存在しない場合は追加
    if (senderIndex <= 0) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue('送信者');
    }
    
    // 最終行を取得
    const lastRow = sheet.getLastRow();
    
    // 顧客IDを生成（現在の時間を使用）
    const customerId = "instagram_" + new Date().getTime();
    
    let addedRows = 0;
    
    if (source === 'OCR') {
      // OCRテキストを行ごとに分割
      const lines = content.split('\n');
      let currentSender = '';
      let currentMessage = '';
      let currentTimestamp = '';
      
      for (let i = 0; i < lines.length; i++) {
        const line = lines[i].trim();
        
        if (line.startsWith('ユーザー:') || line.startsWith('4tenno.ace:')) {
          // 前のメッセージがあれば保存
          if (currentMessage) {
            sheet.getRange(lastRow + addedRows + 1, nameIndex).setValue(currentSender === 'ユーザー' ? '顧客' : '自社');
            sheet.getRange(lastRow + addedRows + 1, messageIndex).setValue(currentMessage);
            sheet.getRange(lastRow + addedRows + 1, idIndex).setValue(customerId);
            
            // タイムスタンプと送信者を設定
            if (timestampIndex > 0) {
              sheet.getRange(lastRow + addedRows + 1, timestampIndex).setValue(currentTimestamp);
            }
            if (senderIndex > 0) {
              sheet.getRange(lastRow + addedRows + 1, senderIndex).setValue(currentSender);
            }
            
            addedRows++;
          }
          
          // 新しいメッセージを開始
          const parts = line.split(':');
          currentSender = parts[0].trim();
          currentMessage = parts.slice(1).join(':').trim();
        } else if (line.startsWith('時間:')) {
          // タイムスタンプを抽出
          currentTimestamp = line.replace('時間:', '').trim();
        } else {
          // 現在のメッセージに追加
          currentMessage += ' ' + line;
        }
      }
      
      // 最後のメッセージを保存
      if (currentMessage) {
        sheet.getRange(lastRow + addedRows + 1, nameIndex).setValue(currentSender === 'ユーザー' ? '顧客' : '自社');
        sheet.getRange(lastRow + addedRows + 1, messageIndex).setValue(currentMessage);
        sheet.getRange(lastRow + addedRows + 1, idIndex).setValue(customerId);
        
        // タイムスタンプと送信者を設定
        if (timestampIndex > 0) {
          sheet.getRange(lastRow + addedRows + 1, timestampIndex).setValue(currentTimestamp);
        }
        if (senderIndex > 0) {
          sheet.getRange(lastRow + addedRows + 1, senderIndex).setValue(currentSender);
        }
        
        addedRows++;
      }
    } else if (source === 'HTML' && Array.isArray(content)) {
      // HTMLから抽出したメッセージ配列を処理
      content.forEach(msg => {
        sheet.getRange(lastRow + addedRows + 1, nameIndex).setValue(msg.sender === '4tenno.ace' ? '自社' : '顧客');
        sheet.getRange(lastRow + addedRows + 1, messageIndex).setValue(msg.text);
        sheet.getRange(lastRow + addedRows + 1, idIndex).setValue(customerId);
        
        // タイムスタンプと送信者を設定
        if (timestampIndex > 0) {
          sheet.getRange(lastRow + addedRows + 1, timestampIndex).setValue(msg.timestamp);
        }
        if (senderIndex > 0) {
          sheet.getRange(lastRow + addedRows + 1, senderIndex).setValue(msg.sender);
        }
        
        addedRows++;
      });
    }
    
    return {
      success: true,
      customerId: customerId,
      messageCount: addedRows
    };
  } catch (error) {
    Logger.log('Error in processDMContent: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

// テキストからの返信候補を生成する高度な関数
function generateAdvancedReplySuggestions(customerId) {
  try {
    // スプレッドシートからこの顧客のやり取りをすべて取得
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    
    // ヘッダー行を取得
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // 各列のインデックスを取得
    const nameIndex = headers.indexOf('顧客名');
    const messageIndex = headers.indexOf('メッセージ');
    const idIndex = headers.indexOf('顧客ID');
    const senderIndex = headers.indexOf('送信者');
    
    // データ範囲を取得
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    // このIDに関連するすべてのメッセージを取得
    const conversation = [];
    
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const rowId = row[idIndex];
      
      if (rowId === customerId) {
        conversation.push({
          sender: senderIndex >= 0 ? row[senderIndex] : row[nameIndex],
          message: row[messageIndex]
        });
      }
    }
    
    // 過去のやり取りに基づいて返信を生成
    // 最後のメッセージを取得
    if (conversation.length === 0) {
      return generateReplySuggestions(customerId, '');
    }
    
    const lastMessage = conversation[conversation.length - 1].message;
    
    // 通常の返信候補生成関数を呼び出す
    return generateReplySuggestions(customerId, lastMessage);
  } catch (error) {
    Logger.log('Error in generateAdvancedReplySuggestions: ' + error.toString());
    return [];
  }
}
}