// ボタン関連の関数
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('テニススコア')
    .addItem('クリア', 'clearData')
    .addItem('登録', 'registerData')
    .addItem('取得', 'getAllGame')
    .addToUi();
}

/**
 * DBから指定した日付の最大ゲーム番号を取得する関数
 * @param {Date|string} gameDate - 検索する日付
 * @return {number} - 指定日付の最大ゲーム番号（該当ゲームがない場合は0）
 */
function getDBMaxGameNumber(gameDate) {
 var ss = SpreadsheetApp.getActiveSpreadsheet();
 var outputSheet = ss.getSheetByName('スコア出力');
 var lastRow = outputSheet.getLastRow();
 
 // 日付をYYYY-MM-DD形式に変換
 var formattedDate = Utilities.formatDate(new Date(gameDate), Session.getScriptTimeZone(), "yyyy-MM-dd");
 
 // ヘッダー行しかない場合または空の場合は0を返す
 if (lastRow <= 1) {
   return 0;
 }
 
 // date, gameNoの列を取得
 var outputData = outputSheet.getRange(2, 1, lastRow - 1, 2).getValues();
 
 // 同じ日付のゲーム番号の最大値を検索
 var maxGameNo = 0;
 for (var i = 0; i < outputData.length; i++) {
   var rowDate = outputData[i][0];
   
   // 日付が有効かチェック
   if (rowDate) {
     // 日付を文字列に変換して比較
     var rowDateStr = Utilities.formatDate(new Date(rowDate), Session.getScriptTimeZone(), "yyyy-MM-dd");
     
     if (rowDateStr === formattedDate && outputData[i][1] > maxGameNo) {
       maxGameNo = outputData[i][1];
     }
   }
 }
 
 return maxGameNo;
}

// データの入力チェックと登録
function registerData() {
  registerOneGame("B2")
  SpreadsheetApp.getUi().alert("データが登録されました。");
}

/**
 * 1ページ分（12ゲーム）のデータを登録する関数
 * @param {string} topLeftCell - ページの左上セル（例: "B18"）
 * @return {Object} - 処理結果の情報（成功したゲーム数、失敗したゲーム数）
 */
function registerOnePage(topLeftCell) {
  // "スコア入力"シート取得
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('スコア入力');
  
  // セル参照をA1形式からrow,columnに変換
  const cellA1Notation = topLeftCell.toUpperCase();
  const cell = sheet.getRange(cellA1Notation);
  const startRow = cell.getRow();
  const startCol = cell.getColumn();
  
  // 処理結果を格納するオブジェクト
  const result = {
    totalGames: 12,
    successCount: 0,
    failedCount: 0,
    failedGames: []
  };
  
  // SheetInfoクラスから試合位置テーブルを取得
  const gamePositions = SheetInfo.positions;
  
  // 12試合分処理
  for (let i = 0; i < gamePositions.length; i++) {
    // 相対位置を取得
    const rowOffset = gamePositions[i][0];
    const colOffset = gamePositions[i][1];
    
    // 実際のセル位置を計算
    const gameRow = startRow + rowOffset;
    const gameCol = startCol + colOffset;
    
    // 現在の試合の開始セルをA1形式で取得
    const gameTopLeftCell = sheet.getRange(gameRow, gameCol).getA1Notation();
    
    // registerOneGame関数を呼び出して1試合分の処理を実行
    const success = registerOneGame(gameTopLeftCell);
    
    // 結果を記録
    if (success) {
      result.successCount++;
      Logger.log(`試合 ${i+1}: ${gameTopLeftCell} の登録成功`);
    } else {
      result.failedCount++;
      result.failedGames.push({
        gameNumber: i + 1,
        cellReference: gameTopLeftCell
      });
      Logger.log(`試合 ${i+1}: ${gameTopLeftCell} の登録失敗`);
    }
  }
  
  // 処理結果のサマリーをログに出力
  Logger.log(`登録処理完了: 成功=${result.successCount}, 失敗=${result.failedCount}`);
  
  // 失敗したゲームがある場合は詳細を表示
  if (result.failedCount > 0) {
    Logger.log("登録に失敗したゲーム:");
    result.failedGames.forEach(game => {
      Logger.log(`試合 ${game.gameNumber}: ${game.cellReference}`);
    });
  }
  
  return result;
}

/**
 * 1ゲームのデータを登録する関数
 * @param {string} topLeftCell - 登録するゲームの左上セル（例: "B2"、"H22"）
 * @return {boolean} - 登録成功時はtrue、失敗時はfalse
 */
function registerOneGame(topLeftCell) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var scoreSheet = ss.getSheetByName('スコア入力');
  var outputSheet = ss.getSheetByName('スコア出力');
  
  // 左上セルの位置を取得
  var cellRange = scoreSheet.getRange(topLeftCell);
  var row = cellRange.getRow();
  var col = cellRange.getColumn();
  
  // 左上セルから日付を取得
  var gameDate = cellRange.getValue();
  
  // B2の日付チェック（必須）
  var mainDateCell = scoreSheet.getRange("B2");
  if (!mainDateCell.getValue()) {
    SpreadsheetApp.getUi().alert("B2セルに日付を入力してください。");
    return false;
  }
  
  // 日付が空の場合、前のゲームの日付を使用
  if (!gameDate) {
    // 出力シートの最後の日付を取得
    var lastRow = outputSheet.getLastRow();
    if (lastRow > 1) {
      gameDate = outputSheet.getRange(lastRow, 1).getValue();
    } else {
      // 出力シートにデータがない場合はB2の日付を使用
      gameDate = mainDateCell.getValue();
    }
  }
  
  // 日付をYYYY-MM-DD形式に変換
  var formattedDate = Utilities.formatDate(new Date(gameDate), Session.getScriptTimeZone(), "yyyy-MM-dd");
  
  // 左上セルの列に応じて、名前列とID列、スコア列を設定
  var nameCol = col + SheetInfo.OFFSET_COL_NAME; // 左上セルから1つ右の列が名前列
  var idCol = col + SheetInfo.OFFSET_COL_ID;     // 左上セルから2つ右の列がID列
  var scoreCol = col + SheetInfo.OFFSET_COL_POINT; // 左上セルから3つ右の列がスコア列
  
  // チームAの選手データを取得
  var memberA1Name = scoreSheet.getRange(row, nameCol).getValue();
  var memberA1Id = scoreSheet.getRange(row, idCol).getValue();
  var memberA2Name = scoreSheet.getRange(row + 1, nameCol).getValue();
  var memberA2Id = scoreSheet.getRange(row + 1, idCol).getValue();
  var scoreA = scoreSheet.getRange(row, scoreCol).getValue();
  
  // チームBの選手データを取得
  var memberB1Name = scoreSheet.getRange(row + 2, nameCol).getValue();
  var memberB1Id = scoreSheet.getRange(row + 2, idCol).getValue();
  var memberB2Name = scoreSheet.getRange(row + 3, nameCol).getValue();
  var memberB2Id = scoreSheet.getRange(row + 3, idCol).getValue();
  var scoreB = scoreSheet.getRange(row + 2, scoreCol).getValue();
  
  // データチェック - 最低限の必要データが揃っているか確認
  if ((!memberA1Id && !memberA2Id) || (!memberB1Id && !memberB2Id) || 
      (scoreA === "" && scoreB === "")) {
    // 必要なデータが不足している場合は登録しない
    return false;
  }
  
  // 最大ゲーム番号を取得
  var maxGameNo = getMaxGameNumber(gameDate);
  var gameCounter = maxGameNo + 1;
  
  // ヘッダー行がない場合は追加
  if (outputSheet.getLastRow() == 0) {
    outputSheet.appendRow(['date', 'gameNo', 'ID', 'pairID', 'serve1st', 'serve2nd', 'gamePt', 'serveTurn', 'row']);
  }
  
  // データを出力シートに追加
  // チームAの選手データを出力
  if (memberA1Id) {
    outputSheet.appendRow([
      formattedDate,
      gameCounter,
      memberA1Id,
      memberA2Id || "", // パートナーIDがなければ空欄
      '', // serve1st
      '', // serve2nd
      scoreA,
      '', // serveTurn
      1 // row
    ]);
  }
  
  if (memberA2Id && memberA2Id !== memberA1Id) {
    outputSheet.appendRow([
      formattedDate,
      gameCounter,
      memberA2Id,
      memberA1Id || "", // パートナーIDがなければ空欄
      '', // serve1st
      '', // serve2nd
      scoreA,
      '', // serveTurn
      2 // row
    ]);
  }
  
  // チームBの選手データを出力
  if (memberB1Id) {
    outputSheet.appendRow([
      formattedDate,
      gameCounter,
      memberB1Id,
      memberB2Id || "", // パートナーIDがなければ空欄
      '', // serve1st
      '', // serve2nd
      scoreB,
      '', // serveTurn
      3 // row
    ]);
  }
  
  if (memberB2Id && memberB2Id !== memberB1Id) {
    outputSheet.appendRow([
      formattedDate,
      gameCounter,
      memberB2Id,
      memberB1Id || "", // パートナーIDがなければ空欄
      '', // serve1st
      '', // serve2nd
      scoreB,
      '', // serveTurn
      4 // row
    ]);
  }
  
  return true;
}
