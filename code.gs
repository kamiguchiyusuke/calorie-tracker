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

// 今月の日ごとの総カロリーを取得
function getMonthlyCalories() {
  try {
    const sheet = getOrCreateSheet();
    const sheetName = sheet.getName();
    const lastRow = sheet.getLastRow();
    
    Logger.log('Sheet name: ' + sheetName);
    Logger.log('Last row: ' + lastRow);
    
    if (lastRow <= 1) {
      return {
        success: true,
        data: []
      };
    }
    
    // 全データを取得（A列からH列まで8列取得）
    const allData = sheet.getRange(1, 1, lastRow, 8).getValues();
    
    // ヘッダー行を探す
    let headerRow = 0;
    for (let i = 0; i < allData.length; i++) {
      if (allData[i][0] === '日付' || allData[i][0] === 'Date') {
        headerRow = i;
        Logger.log('Found header at row: ' + (i + 1));
        break;
      }
    }
    
    // ヘッダー行の次の行からデータ開始
    const dataStartRow = headerRow + 1;
    
    if (dataStartRow >= allData.length) {
      Logger.log('No data rows found');
      return {
        success: true,
        data: []
      };
    }
    
    const values = allData.slice(dataStartRow);
    Logger.log('Total data rows: ' + values.length);
    
    // 日付ごとにカロリーを集計
    const dailyCalories = {};
    
    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      const dateValue = row[0]; // A列: 日付
      const caloriesValue = row[7]; // H列: 摂取カロリー（インデックス7）
      
      Logger.log(`Data row ${i + 1}: Date raw=${dateValue}, Calories raw=${caloriesValue}`);
      
      // 空行をスキップ
      if (!dateValue || dateValue === '') {
        Logger.log(`Data row ${i + 1}: Skipping - empty date`);
        continue;
      }
      
      // カロリーが空をスキップ
      if (caloriesValue === '' || caloriesValue === null || caloriesValue === undefined) {
        Logger.log(`Data row ${i + 1}: Skipping - empty calories`);
        continue;
      }
      
      // 日付を文字列に変換
      let dateStr;
      if (dateValue instanceof Date) {
        dateStr = Utilities.formatDate(dateValue, 'Asia/Tokyo', 'yyyy-MM-dd');
      } else if (typeof dateValue === 'string') {
        dateStr = dateValue.substring(0, 10);
      } else {
        Logger.log(`Data row ${i + 1}: Unknown date type: ${typeof dateValue}`);
        continue;
      }
      
      // カロリーを数値に変換
      const calories = Number(caloriesValue);
      if (isNaN(calories)) {
        Logger.log(`Data row ${i + 1}: Invalid calories: ${caloriesValue}`);
        continue;
      }
      
      Logger.log(`Data row ${i + 1}: Valid - Date=${dateStr}, Calories=${calories}`);
      
      // 日付ごとに合計
      if (!dailyCalories[dateStr]) {
        dailyCalories[dateStr] = 0;
      }
      dailyCalories[dateStr] += calories;
      
      Logger.log(`Data row ${i + 1}: Running total for ${dateStr} = ${dailyCalories[dateStr]}`);
    }
    
    Logger.log('Final dailyCalories object: ' + JSON.stringify(dailyCalories));
    
    // 配列に変換してソート
    const result = Object.keys(dailyCalories)
      .sort()
      .map(date => ({
        date: date,
        calories: Math.round(dailyCalories[date])
      }));
    
    Logger.log('Final result array: ' + JSON.stringify(result));
    
    return {
      success: true,
      data: result
    };
  } catch (error) {
    Logger.log('ERROR in getMonthlyCalories: ' + error.message);
    Logger.log('Stack trace: ' + error.stack);
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
