// Webアプリのエントリーポイント
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .addMetaTag('viewport', 'initial-scale=0.4, user-scalable=no')
    .setTitle('カロリー記録')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 年に対応したスプレッドシートを検索して取得
function findSpreadsheetByYear(year) {
  const fileName = `カロリー管理_${year}`;

  // Google Driveで該当ファイルを検索
  const files = DriveApp.getFilesByName(fileName);

  if (files.hasNext()) {
    const file = files.next();
    return SpreadsheetApp.openById(file.getId());
  }

  // ファイルが見つからない場合はエラー
  throw new Error(`${fileName} というスプレッドシートが見つかりません。先に作成してください。`);
}

// お気に入りシートを取得または作成
function getOrCreateFavoritesSheet() {
  const now = new Date();
  const year = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy');
  const ss = findSpreadsheetByYear(year);

  let sheet = ss.getSheetByName('お気に入り');

  if (!sheet) {
    sheet = ss.insertSheet('お気に入り');
    // ヘッダー行を追加
    sheet.appendRow(['食べ物', '重さ(g)', '摂取カロリー(kcal)']);

    // ヘッダー行のスタイル設定
    const headerRange = sheet.getRange(1, 1, 1, 3);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#f3f3f3');
  }

  return sheet;
}

// 年と月に対応したシートを取得または作成
function getOrCreateSheet() {
  const now = new Date();
  const year = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy');
  const month = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyyMM');

  // 年のスプレッドシートを検索
  const ss = findSpreadsheetByYear(year);

  Logger.log('Looking for sheet: ' + month);
  Logger.log('Available sheets: ' + ss.getSheets().map(s => s.getName()).join(', '));

  let sheet = ss.getSheetByName(month);

  // シートが存在しなければ作成
  if (!sheet) {
    Logger.log('Sheet not found, creating: ' + month);
    sheet = ss.insertSheet(month);
    // ヘッダー行を追加（画像に合わせた構造）
    sheet.appendRow(['日付', '時間', '食べ物', '', '', '重さ(g)', '摂取カロリー(kcal:キロカロリー)', '', '', '日付', '合計カロリー', '', '月の総合摂取カロリー(kcal)', '', '日平均摂取カロリー(kcal)']);

    // ヘッダー行のスタイル設定
    const headerRange = sheet.getRange(1, 1, 1, 15);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#f3f3f3');
  } else {
    Logger.log('Found sheet: ' + sheet.getName());
  }

  return sheet;
}

// カロリーを記録する関数
function addCalorieRecord(food, weight, calories, addToFavorite) {
  try {
    const sheet = getOrCreateSheet();

    // 現在時刻を取得
    const now = new Date();
    const dateStr = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd');
    const timeStr = Utilities.formatDate(now, 'Asia/Tokyo', 'HH:mm');

    // ヘッダー行を探す
    const allData = sheet.getDataRange().getValues();
    let headerRow = 0;
    for (let i = 0; i < allData.length; i++) {
      if (allData[i][0] === '日付' || allData[i][0] === 'Date') {
        headerRow = i + 1; // スプレッドシートは1始まり
        break;
      }
    }

    // データの最終行を探す（A列に日付があるかで判定）
    let lastDataRow = headerRow;
    for (let i = headerRow; i < allData.length; i++) {
      if (allData[i][0] && allData[i][0] !== '') {
        lastDataRow = i + 1; // スプレッドシートは1始まり
      }
    }

    // 次の行に書き込む
    const nextRow = lastDataRow + 1;

    // 特定の行に書き込み
    sheet.getRange(nextRow, 1).setValue(dateStr);  // A列: 日付
    sheet.getRange(nextRow, 2).setValue(timeStr);  // B列: 時間
    sheet.getRange(nextRow, 3).setValue(food);     // C列: 食べ物
    sheet.getRange(nextRow, 7).setValue(weight);   // G列: 重さ
    sheet.getRange(nextRow, 8).setValue(calories); // H列: 摂取カロリー

    // お気に入りに追加
    if (addToFavorite) {
      addToFavorites(food, weight, calories);
    }

    return {
      success: true,
      message: '記録しました！',
      data: {
        date: dateStr,
        time: timeStr,
        food: food,
        weight: weight,
        calories: calories
      }
    };
  } catch (error) {
    return {
      success: false,
      message: 'エラー: ' + error.message
    };
  }
}

// お気に入りを取得
function getFavorites() {
  try {
    const sheet = getOrCreateFavoritesSheet();
    const lastRow = sheet.getLastRow();

    if (lastRow <= 1) {
      return {
        success: true,
        data: []
      };
    }

    // ヘッダーを除いてデータを取得
    const values = sheet.getRange(2, 1, lastRow - 1, 3).getValues();

    const favorites = values
      .filter(row => row[0] && row[0] !== '') // 空行を除外
      .map(row => ({
        food: row[0],
        weight: row[1],
        calories: row[2]
      }));

    return {
      success: true,
      data: favorites
    };
  } catch (error) {
    Logger.log('ERROR in getFavorites: ' + error.message);
    return {
      success: false,
      message: 'エラー: ' + error.message,
      data: []
    };
  }
}

// お気に入りに追加
function addToFavorites(food, weight, calories) {
  try {
    const sheet = getOrCreateFavoritesSheet();

    // 既に存在するかチェック
    const favorites = getFavorites();
    if (favorites.success && favorites.data) {
      const exists = favorites.data.some(fav => 
        fav.food === food && 
        fav.weight === weight && 
        fav.calories === calories
      );

      if (exists) {
        return {
          success: true,
          message: 'このお気に入りは既に登録されています'
        };
      }
    }

    // 新しい行に追加
    sheet.appendRow([food, weight, calories]);

    return {
      success: true,
      message: 'お気に入りに追加しました'
    };
  } catch (error) {
    return {
      success: false,
      message: 'エラー: ' + error.message
    };
  }
}

// お気に入りを削除
function removeFavorite(index) {
  try {
    const sheet = getOrCreateFavoritesSheet();
    const lastRow = sheet.getLastRow();

    if (lastRow <= 1) {
      return {
        success: false,
        message: 'お気に入りがありません'
      };
    }

    // インデックスは0始まりなので、実際の行番号は index + 2 (ヘッダー行 + 1)
    const rowToDelete = index + 2;

    if (rowToDelete > lastRow) {
      return {
        success: false,
        message: '削除対象が見つかりません'
      };
    }

    sheet.deleteRow(rowToDelete);
    
    return {
      success: true,
      message: 'お気に入りを削除しました'
    };
  } catch (error) {
    return {
      success: false,
      message: 'エラー: ' + error.message
    };
  }
}

function getMonthlyCalories(targetYearMonth) {
  try {
    const now = new Date();
    const yearMonth = targetYearMonth || Utilities.formatDate(now, 'Asia/Tokyo', 'yyyyMM');
    
    const year = yearMonth.substring(0, 4);
    const ss = findSpreadsheetByYear(year);
    
    let sheet = ss.getSheetByName(yearMonth);
    
    if (!sheet) {
      return { success: true, data: [] };
    }
    
    const lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) {
      return { success: true, data: [] };
    }
    
    // 1回で全データ取得（元のまま）
    const allData = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
    
    const dailyCalories = {};
    
    for (let i = 0; i < allData.length; i++) {
      const row = allData[i];
      const dateValue = row[0];  // A列
      const caloriesValue = row[7];  // H列（インデックス7）
      
      if (!dateValue || dateValue === '' || !caloriesValue || caloriesValue === '') {
        continue;
      }
      
      let dateStr;
      if (dateValue instanceof Date) {
        dateStr = Utilities.formatDate(dateValue, 'Asia/Tokyo', 'yyyy-MM-dd');
      } else if (typeof dateValue === 'string') {
        dateStr = dateValue.substring(0, 10);
      } else {
        continue;
      }
      
      const calories = Number(caloriesValue);
      if (isNaN(calories)) {
        continue;
      }
      
      if (!dailyCalories[dateStr]) {
        dailyCalories[dateStr] = 0;
      }
      dailyCalories[dateStr] += calories;
    }
    
    const result = {
      success: true,
      data: Object.keys(dailyCalories)
        .sort()
        .map(date => ({
          date: date,
          calories: Math.round(dailyCalories[date])
        }))
    };
    
    return result;
    
  } catch (error) {
    Logger.log('ERROR in getMonthlyCalories: ' + error.message);
    return {
      success: false,
      message: 'エラー: ' + error.message,
      data: []
    };
  }
}


// 指定された年月の日ごとの総カロリーを取得
function getAvailableMonths() {
  try {
    const now = new Date();
    const currentYear = parseInt(Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy'));
    
    const months = [];
    
    // 過去2年分をチェック
    for (let year = currentYear; year >= currentYear - 2; year--) {
      try {
        const ss = findSpreadsheetByYear(year.toString());
        const sheets = ss.getSheets();
        
        sheets.forEach(sheet => {
          const sheetName = sheet.getName();
          // yyyyMM形式のシート名のみを抽出
          if (/^\d{6}$/.test(sheetName)) {
            months.push({
              value: sheetName,
              label: sheetName.substring(0, 4) + '年' + sheetName.substring(4, 6) + '月'
            });
          }
        });
      } catch (e) {
        // その年のスプレッドシートが存在しない場合はスキップ
        Logger.log('Spreadsheet not found for year: ' + year);
      }
    }
    
    // 降順でソート（新しい月が先）
    months.sort((a, b) => b.value.localeCompare(a.value));
    
    return {
      success: true,
      data: months
    };
  } catch (error) {
    Logger.log('ERROR in getAvailableMonths: ' + error.message);
    return {
      success: false,
      message: 'エラー: ' + error.message,
      data: []
    };
  }
}

// デバッグ用
function debugGetMonthlyCalories() {
  const result = getMonthlyCalories();
  Logger.log('=== DEBUG RESULT ===');
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

// スプレッドシートのURLを取得
function getSpreadsheetUrl() {
  try {
    const now = new Date();
    const year = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy');
    const ss = findSpreadsheetByYear(year);
    
    return {
      success: true,
      url: ss.getUrl()
    };
  } catch (error) {
    return {
      success: false,
      message: 'エラー: ' + error.message
    };
  }
}

// 現在のスプレッドシート情報を取得（デバッグ用）
function getCurrentSheetInfo() {
  try {
    const now = new Date();
    const year = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy');
    const month = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyyMM');
    const ss = findSpreadsheetByYear(year);
    
    return {
      year: year,
      month: month,
      spreadsheetName: ss.getName(),
      spreadsheetId: ss.getId(),
      url: ss.getUrl()
    };
  } catch (error) {
    return {
      error: error.message
    };
  }
}

// 新年のスプレッドシートを自動作成する関数（オプション）
function createNewYearSpreadsheet() {
  const now = new Date();
  const year = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy');
  const month = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyyMM');
  const fileName = `カロリー管理_${year}`;
  
  // 既に存在するかチェック
  const files = DriveApp.getFilesByName(fileName);
  if (files.hasNext()) {
    return {
      success: false,
      message: `${fileName} は既に存在します`
    };
  }
  
  // 新しいスプレッドシートを作成
  const ss = SpreadsheetApp.create(fileName);
  
  // デフォルトシートの名前を変更
  const defaultSheet = ss.getSheets()[0];
  defaultSheet.setName(month);
  
  // ヘッダー行を追加
  defaultSheet.appendRow(['日付', '時間', '食べ物', '', '', '重さ(g)', '摂取カロリー(kcal:キロカロリー)', '', '', '日付', '合計カロリー', '', '月の総合摂取カロリー(kcal)', '', '日平均摂取カロリー(kcal)']);
  const headerRange = defaultSheet.getRange(1, 1, 1, 15);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#f3f3f3');
  
  return {
    success: true,
    message: `${fileName} を作成しました`,
    url: ss.getUrl()
  };
}

function debugSpreadsheetStructure() {
  const sheet = getOrCreateSheet();
  const lastRow = sheet.getLastRow();

  Logger.log('=== SPREADSHEET STRUCTURE ===');
  Logger.log('Sheet name: ' + sheet.getName());
  Logger.log('Last row: ' + lastRow);

  // 全データを取得（A列からH列まで8列）
  const allData = sheet.getRange(1, 1, lastRow, 8).getValues();

  for (let i = 0; i < allData.length; i++) {
    const row = allData[i];
    Logger.log(`Row ${i + 1}: A="${row[0]}", B="${row[1]}", C="${row[2]}", F="${row[5]}", G="${row[6]}", H="${row[7]}"`);
  }

  Logger.log('=== END ===');
}

// ==================== AI機能 ====================

/**
 * Gemini APIを使ってカロリーを推定する
 * @param {string} userQuery - ユーザーの質問（テキスト）
 * @param {string} imageBase64 - Base64エンコードされた画像（オプション）
 * @return {Object} - 推定結果
 */
function estimateCaloriesWithAI(userQuery, imageBase64) {
  try {
    // スクリプトプロパティからAPIキーを取得
    const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    
    if (!apiKey) {
      return {
        success: false,
        message: 'Gemini APIキーが設定されていません。スクリプトエディタの「プロジェクトの設定」→「スクリプト プロパティ」から「GEMINI_API_KEY」を設定してください。'
      };
    }

    const url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' + apiKey;
    
    // プロンプトを作成
    const systemPrompt = `あなたは栄養士です。ユーザーが食べた食事の内容から、以下の情報を推定してください:

1. 食べ物名（具体的に）
2. 重さ（グラム、gで）
3. カロリー（kcal）

回答は必ず以下のJSON形式のみで返してください（他の文章は一切含めないでください）:
{
  "food": "食べ物名",
  "weight": 重さの数値,
  "calories": カロリーの数値
}

例:
入力: "唐揚げ5個食べました"
出力: {"food": "唐揚げ", "weight": 150, "calories": 450}`;

    // リクエストボディを構築
    const parts = [
      { text: systemPrompt },
      { text: "\n\nユーザーの入力: " + userQuery }
    ];
    
    // 画像がある場合は追加
    if (imageBase64) {
      // data:image/jpeg;base64, などのプレフィックスを削除
      const base64Data = imageBase64.replace(/^data:image\/\w+;base64,/, '');
      
      parts.push({
        inline_data: {
          mime_type: "image/jpeg",
          data: base64Data
        }
      });
    }

    const payload = {
      contents: [{
        parts: parts
      }],
      generationConfig: {
        temperature: 0.4,
        maxOutputTokens: 500
      }
    };

    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    Logger.log('Calling Gemini API...');
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    Logger.log('Response code: ' + responseCode);
    Logger.log('Response: ' + responseText);

    if (responseCode !== 200) {
      return {
        success: false,
        message: 'APIエラー: ' + responseText
      };
    }

    const result = JSON.parse(responseText);
    
    // レスポンスからテキストを抽出
    if (!result.candidates || !result.candidates[0] || !result.candidates[0].content) {
      return {
        success: false,
        message: 'AIからの応答が不正です'
      };
    }

    const aiText = result.candidates[0].content.parts[0].text;
    Logger.log('AI response text: ' + aiText);

    // JSONを抽出（```json ... ```で囲まれている場合に対応）
    let jsonText = aiText.trim();
    const jsonMatch = jsonText.match(/```json\s*([\s\S]*?)\s*```/);
    if (jsonMatch) {
      jsonText = jsonMatch[1].trim();
    } else {
      // {}で囲まれた部分を探す
      const bracketMatch = jsonText.match(/\{[\s\S]*\}/);
      if (bracketMatch) {
        jsonText = bracketMatch[0];
      }
    }

    Logger.log('Extracted JSON: ' + jsonText);

    // JSONをパース
    const calorieData = JSON.parse(jsonText);

    // バリデーション
    if (!calorieData.food || !calorieData.weight || !calorieData.calories) {
      return {
        success: false,
        message: 'AIの応答が不完全です: ' + aiText
      };
    }

    return {
      success: true,
      data: {
        food: String(calorieData.food),
        weight: Number(calorieData.weight),
        calories: Number(calorieData.calories)
      },
      aiResponse: aiText
    };

  } catch (error) {
    Logger.log('ERROR in estimateCaloriesWithAI: ' + error.message);
    Logger.log('Stack trace: ' + error.stack);
    return {
      success: false,
      message: 'エラー: ' + error.message
    };
  }
}

// テスト用関数
function testAIEstimate() {
  const result = estimateCaloriesWithAI('唐揚げ5個食べました', null);
  Logger.log('=== TEST RESULT ===');
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

function forceAuthorization() {
  // 1. スプレッドシート権限
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log('Spreadsheet OK');
  
  // 2. 外部リクエスト権限（これが重要！）
  try {
    const response = UrlFetchApp.fetch('https://www.google.com');
    Logger.log('UrlFetch OK: ' + response.getResponseCode());
  } catch (e) {
    Logger.log('UrlFetch ERROR: ' + e.message);
  }
}
