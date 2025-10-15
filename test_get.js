/**
 * テストを開始する前の準備
 */
function prepareForTests() {
  // デバッグモードを有効化
  UIHelper.setDebugMode(true);
  Logger.log("テスト環境を準備しました - UIアラートは表示されません");
}

/**
 * テスト後の後片付け
 */
function cleanupAfterTests() {
  // デバッグモードを無効化
  UIHelper.setDebugMode(false);
  Logger.log("テスト環境をクリーンアップしました - 通常モードに戻しました");
}

/**
 * getOneGame関数の単体テスト - 正常系
 */
function test_getOneGame_B3_OK() {
  // テスト環境の準備
  prepareForTests();
  
  let result = "エラー"; // デフォルト値
  
  try {
    Logger.log("実行関数: test_getOneGame_B3_OK");
    // テスト前にバッファをクリア
    gameCollectorInstance.clearBuffer();
    
    // テスト前の準備：テスト用データを設定
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const scoreSheet = ss.getSheetByName('スコア入力');
    
    // 元の値を保存
    const originalValues = saveOriginalValues("B3");
    const originalDate = scoreSheet.getRange(SheetInfo.DATE_CELL).getValue();
    
    // テストデータを設定
    scoreSheet.getRange(SheetInfo.DATE_CELL).setValue(new Date("2025/04/15")); // B1に日付
    scoreSheet.getRange("C3").setValue(407);                   // チームA ID1
    scoreSheet.getRange("C4").setValue(39);                    // チームA ID2
    scoreSheet.getRange("D3").setValue(2);                     // チームA スコア
    scoreSheet.getRange("C5").setValue(491);                   // チームB ID1
    scoreSheet.getRange("C6").setValue(66);                    // チームB ID2
    scoreSheet.getRange("D5").setValue(5);                     // チームB スコア
    
    // テスト実行
    const success = gameCollectorInstance.getOneGame("B3");
    
    // 元の値に戻す
    restoreOriginalValues("B3", originalValues);
    scoreSheet.getRange(SheetInfo.DATE_CELL).setValue(originalDate);
    
    // 結果を検証
    if (success && gameCollectorInstance.success) {
      Logger.log("テスト成功: getOneGame関数が正常に動作しました");
      
      // 期待される結果と実際の結果を比較
      const expectedGameCounter = getDBMaxGameNumber(new Date("2025/04/15")) + 1;
      const buffer = gameCollectorInstance.getBuffer();
      
      // バッファに4つのデータがあることを確認
      if (buffer.length !== 4) {
        Logger.log(`テスト失敗: バッファのサイズが不正です。期待値: 4, 実際: ${buffer.length}`);
        result = "失敗";
        return result;
      }
      
      // データの内容を検証
      const expectedData = [
        { ID: 407, pairID: 39, gamePt: 2, row: 1 },
        { ID: 39, pairID: 407, gamePt: 2, row: 2 },
        { ID: 491, pairID: 66, gamePt: 5, row: 3 },
        { ID: 66, pairID: 491, gamePt: 5, row: 4 }
      ];
      
      let isDataValid = true;
      
      for (let i = 0; i < 4; i++) {
        if (buffer[i].ID != expectedData[i].ID || 
            buffer[i].pairID != expectedData[i].pairID || 
            buffer[i].gamePt != expectedData[i].gamePt || 
            buffer[i].row != expectedData[i].row) {
          isDataValid = false;
          Logger.log(`データ不一致 [${i}]: 
            実際: ID=${buffer[i].ID}, pairID=${buffer[i].pairID}, gamePt=${buffer[i].gamePt}, row=${buffer[i].row}
            期待: ID=${expectedData[i].ID}, pairID=${expectedData[i].pairID}, gamePt=${expectedData[i].gamePt}, row=${expectedData[i].row}`);
        }
      }
      
      if (!isDataValid) {
        result = "失敗";
        return result;
      }
      
      // ゲーム番号の検証
      if (buffer[0].gameNo !== expectedGameCounter) {
        Logger.log(`ゲーム番号が不正です: ${buffer[0].gameNo} (期待値: ${expectedGameCounter})`);
        result = "失敗";
        return result;
      }
      
      result = "成功";
    } else {
      Logger.log("テスト失敗: getOneGame関数がfalseを返しました");
      result = "失敗";
    }
  } catch (error) {
    Logger.log(`テスト実行中にエラーが発生しました: ${error.message}`);
    result = "エラー";
  } finally {
    // テスト環境のクリーンアップ（常に実行される）
    cleanupAfterTests();
  }
  
  return result;
}

/**
 * getOneGame関数の単体テスト - 日付なしケース（失敗ケース）
 */
function test_getOneGame_B8_NG() {
  // テスト環境の準備
  prepareForTests();
  
  let result = "エラー"; // デフォルト値
  
  try {
    Logger.log("実行関数: test_getOneGame_B8_NG");
    // テスト前にバッファをクリア
    gameCollectorInstance.clearBuffer();
    
    // テスト前の準備：テスト用データを設定
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const scoreSheet = ss.getSheetByName('スコア入力');
    
    // 元の値を保存
    const originalValues = saveOriginalValues("B8");
    const originalDate = scoreSheet.getRange(SheetInfo.DATE_CELL).getValue();
    
    // テストデータを設定（B1に日付なし）
    scoreSheet.getRange(SheetInfo.DATE_CELL).setValue("");   // B1に日付なし
    scoreSheet.getRange("C8").setValue(407);                 // チームA ID1
    scoreSheet.getRange("C9").setValue(39);                  // チームA ID2
    scoreSheet.getRange("D8").setValue(5);                   // チームA スコア
    scoreSheet.getRange("C10").setValue(491);                // チームB ID1
    scoreSheet.getRange("C11").setValue(66);                 // チームB ID2
    scoreSheet.getRange("D10").setValue(0);                  // チームB スコア
    
    // テスト実行
    const success = gameCollectorInstance.getOneGame("B8");
    
    // 元の値に戻す
    restoreOriginalValues("B8", originalValues);
    scoreSheet.getRange(SheetInfo.DATE_CELL).setValue(originalDate);
    
    // 結果を検証
    if (!success && !gameCollectorInstance.success) {
      Logger.log("テスト成功: 日付が空白の場合、getOneGame関数がfalseを返しました");
      result = "成功";
    } else {
      Logger.log("テスト失敗: 日付が空白なのにgetOneGame関数がtrueを返しました");
      result = "失敗";
    }
  } catch (error) {
    Logger.log(`テスト実行中にエラーが発生しました: ${error.message}`);
    result = "エラー";
  } finally {
    // テスト環境のクリーンアップ（常に実行される）
    cleanupAfterTests();
  }
  
  return result;
}

/**
 * getOneGame関数の単体テスト - 日付引数指定のケース（成功ケース）
 */
function test_getOneGame_B8_OK() {
  // テスト環境の準備
  prepareForTests();
  
  let result = "エラー"; // デフォルト値
  
  try {
    Logger.log("実行関数: test_getOneGame_B8_OK");
    // テスト前にバッファをクリア
    gameCollectorInstance.clearBuffer();
    
    // テスト前の準備：テスト用データを設定
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const scoreSheet = ss.getSheetByName('スコア入力');
    
    // 元の値を保存
    const originalValues = saveOriginalValues("B8");
    
    // テストデータを設定（B1の日付は関係なく引数で渡す）
    scoreSheet.getRange("C8").setValue(407);                 // チームA ID1
    scoreSheet.getRange("C9").setValue(39);                  // チームA ID2
    scoreSheet.getRange("D8").setValue(5);                   // チームA スコア
    scoreSheet.getRange("C10").setValue(491);                // チームB ID1
    scoreSheet.getRange("C11").setValue(66);                 // チームB ID2
    scoreSheet.getRange("D10").setValue(0);                  // チームB スコア
    
    // テスト実行 - 日付引数を指定
    const success = gameCollectorInstance.getOneGame("B8", "2025/04/15");
    
    // 元の値に戻す
    restoreOriginalValues("B8", originalValues);
    
    // 結果を検証
    if (success && gameCollectorInstance.success) {
      Logger.log("テスト成功: 日付引数を指定した場合、getOneGame関数がtrueを返しました");
      
      // バッファを取得
      const buffer = gameCollectorInstance.getBuffer();
      
      // 日付の確認
      const dateStr = Utilities.formatDate(new Date(buffer[0].date), Session.getScriptTimeZone(), "yyyy/MM/dd");
      if (dateStr !== "2025/04/15") {
        Logger.log(`日付が不正です: ${dateStr} (期待値: 2025/04/15)`);
        result = "失敗";
        return result;
      }
      
      result = "成功";
    } else {
      Logger.log("テスト失敗: 日付引数を指定したのにgetOneGame関数がfalseを返しました");
      result = "失敗";
    }
  } catch (error) {
    Logger.log(`テスト実行中にエラーが発生しました: ${error.message}`);
    result = "エラー";
  } finally {
    // テスト環境のクリーンアップ（常に実行される）
    cleanupAfterTests();
  }
  
  return result;
}

/**
 * getOneGame関数の単体テスト - ID重複のケース（失敗ケース）
 */
function test_getOneGame_B13_NG() {
  // テスト環境の準備
  prepareForTests();
  
  let result = "エラー"; // デフォルト値
  
  try {
    Logger.log("実行関数: test_getOneGame_B13_NG");
    // テスト前にバッファをクリア
    gameCollectorInstance.clearBuffer();
    
    // テスト前の準備：テスト用データを設定
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const scoreSheet = ss.getSheetByName('スコア入力');
    
    // 元の値を保存
    const originalValues = saveOriginalValues("B13");
    
    // テストデータを設定（ID重複）
    scoreSheet.getRange("C13").setValue(66);                 // チームA ID1
    scoreSheet.getRange("C14").setValue(39);                 // チームA ID2
    scoreSheet.getRange("D13").setValue(3);                  // チームA スコア
    scoreSheet.getRange("C15").setValue(73);                 // チームB ID1
    scoreSheet.getRange("C16").setValue(39);                 // チームB ID2（重複）
    scoreSheet.getRange("D15").setValue(5);                  // チームB スコア
    
    // テスト実行 - 日付引数を指定
    const success = gameCollectorInstance.getOneGame("B13", "2025/04/15");
    
    // 元の値に戻す
    restoreOriginalValues("B13", originalValues);
    
    // 結果を検証
    if (!success && !gameCollectorInstance.success) {
      Logger.log("テスト成功: ID重複の場合、getOneGame関数がfalseを返しました");
      result = "成功";
    } else {
      Logger.log("テスト失敗: ID重複なのにgetOneGame関数がtrueを返しました");
      result = "失敗";
    }
  } catch (error) {
    Logger.log(`テスト実行中にエラーが発生しました: ${error.message}`);
    result = "エラー";
  } finally {
    // テスト環境のクリーンアップ（常に実行される）
    cleanupAfterTests();
  }
  
  return result;
}

/**
 * getOneGame関数の単体テスト - スコア不正ケース（失敗ケース）
 */
function test_getOneGame_B18_NG() {
  // テスト環境の準備
  prepareForTests();
  
  let result = "エラー"; // デフォルト値
  
  try {
    Logger.log("実行関数: test_getOneGame_B18_NG");
    // テスト前にバッファをクリア
    gameCollectorInstance.clearBuffer();
    
    // テスト前の準備：テスト用データを設定
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const scoreSheet = ss.getSheetByName('スコア入力');
    
    // 元の値を保存
    const originalValues = saveOriginalValues("B18");
    const originalDate = scoreSheet.getRange(SheetInfo.DATE_CELL).getValue();
    
    // テストデータを設定（スコア不正）
    scoreSheet.getRange(SheetInfo.DATE_CELL).setValue(new Date("2025/04/17")); // B1に日付
    scoreSheet.getRange("C18").setValue(66);                 // チームA ID1
    scoreSheet.getRange("C19").setValue(39);                 // チームA ID2
    scoreSheet.getRange("D18").setValue(3);                  // チームA スコア
    scoreSheet.getRange("C20").setValue(73);                 // チームB ID1
    scoreSheet.getRange("C21").setValue(491);                // チームB ID2
    scoreSheet.getRange("D20").setValue(3);                  // チームB スコア（不正）
    
    // テスト実行
    const success = gameCollectorInstance.getOneGame("B18");
    
    // 元の値に戻す
    restoreOriginalValues("B18", originalValues);
    scoreSheet.getRange(SheetInfo.DATE_CELL).setValue(originalDate);
    
    // 結果を検証
    if (!success && !gameCollectorInstance.success) {
      Logger.log("テスト成功: スコアが不正な場合（3対3）、getOneGame関数がfalseを返しました");
      result = "成功";
    } else {
      Logger.log("テスト失敗: スコアが不正（3対3）なのにgetOneGame関数がtrueを返しました");
      result = "失敗";
    }
  } catch (error) {
    Logger.log(`テスト実行中にエラーが発生しました: ${error.message}`);
    result = "エラー";
  } finally {
    // テスト環境のクリーンアップ（常に実行される）
    cleanupAfterTests();
  }
  
  return result;
}

/**
 * getOneSheet関数の単体テスト
 */
function test_getOneSheet() {
  // テスト環境の準備
  prepareForTests();
  
  let result = "エラー"; // デフォルト値
  
  try {
    Logger.log("実行関数: test_getOneSheet");
    // テスト前にバッファをクリア
    gameCollectorInstance.clearBuffer();
    
    // テスト実行
    const testResult = gameCollectorInstance.getOneSheet("B3");
    
    // 結果の基本情報を検証
    Logger.log(`収集結果: 成功=${testResult.successCount}, 失敗=${testResult.failedCount}, 総試合数=${testResult.totalGames}`);
    
    const buffer = gameCollectorInstance.getBuffer();
    Logger.log(`バッファサイズ: ${buffer.length} レコード`);
    
    // シート上に有効なデータがない場合もあるので、結果の情報だけを検証
    if (testResult && typeof testResult.successCount !== 'undefined') {
      Logger.log("テスト成功: getOneSheet関数が正しく実行されました");
      
      // 収集されたデータがある場合はサンプル表示
      if (buffer.length > 0) {
        Logger.log("収集されたデータ（先頭サンプル）:");
        Logger.log(JSON.stringify(buffer[0]));
      }
      
      result = "成功";
    } else {
      Logger.log("テスト失敗: getOneSheet関数が正しい結果を返しませんでした");
      result = "失敗";
    }
  } catch (error) {
    Logger.log(`テスト実行中にエラーが発生しました: ${error.message}`);
    result = "エラー";
  } finally {
    // テスト環境のクリーンアップ（常に実行される）
    cleanupAfterTests();
  }
  
  return result;
}

/**
 * getBufferMaxGameNumber関数のテスト
 */
function test_getBufferMaxGameNumber() {
  // テスト環境の準備
  prepareForTests();
  
  let result = "エラー"; // デフォルト値
  
  try {
    Logger.log("実行関数: test_getBufferMaxGameNumber");
    // テスト前にバッファをクリア
    gameCollectorInstance.clearBuffer();
    
    // テスト用にバッファにデータを直接追加
    gameCollectorInstance.buffer = [
      // 2025/04/20のゲーム
      { date: new Date("2025/04/20"), gameNo: 1, ID: "001", row: 1 },
      { date: new Date("2025/04/20"), gameNo: 2, ID: "002", row: 2 },
      // 2025/04/21のゲーム
      { date: new Date("2025/04/21"), gameNo: 1, ID: "003", row: 1 },
      { date: new Date("2025/04/21"), gameNo: 3, ID: "004", row: 2 }
    ];
    
    // テスト実行
    const maxNo1 = gameCollectorInstance.getBufferMaxGameNumber("2025/04/20");
    const maxNo2 = gameCollectorInstance.getBufferMaxGameNumber("2025/04/21");
    const maxNo3 = gameCollectorInstance.getBufferMaxGameNumber("2025/04/22"); // 存在しない日付
    
    // 結果を検証
    let testPassed = true;
    
    if (maxNo1 !== 2) {
      Logger.log(`テスト失敗: 2025/04/20の最大ゲーム番号 - 期待値:2, 実際:${maxNo1}`);
      testPassed = false;
    }
    
    if (maxNo2 !== 3) {
      Logger.log(`テスト失敗: 2025/04/21の最大ゲーム番号 - 期待値:3, 実際:${maxNo2}`);
      testPassed = false;
    }
    
    if (maxNo3 !== 0) {
      Logger.log(`テスト失敗: 2025/04/22の最大ゲーム番号 - 期待値:0, 実際:${maxNo3}`);
      testPassed = false;
    }
    
    if (testPassed) {
      Logger.log("テスト成功: getBufferMaxGameNumber関数が正しく動作しました");
      result = "成功";
    } else {
      result = "失敗";
    }
  } catch (error) {
    Logger.log(`テスト実行中にエラーが発生しました: ${error.message}`);
    result = "エラー";
  } finally {
    // テスト環境のクリーンアップ（常に実行される）
    cleanupAfterTests();
  }
  
  return result;
}

/**
 * 全テストケースを実行
 */
function runAllTests() {
  // テスト環境の準備
  prepareForTests();
  
  Logger.log("===== GameCollector テストスイート 実行開始 =====");
  
  const testFunctions = [
    { name: "test_getOneGame_B3_OK", func: test_getOneGame_B3_OK },
    { name: "test_getOneGame_B8_NG", func: test_getOneGame_B8_NG },
    { name: "test_getOneGame_B8_OK", func: test_getOneGame_B8_OK },
    { name: "test_getOneGame_B13_NG", func: test_getOneGame_B13_NG },
    { name: "test_getOneGame_B18_NG", func: test_getOneGame_B18_NG },
    { name: "test_getOneSheet", func: test_getOneSheet },
    { name: "test_getBufferMaxGameNumber", func: test_getBufferMaxGameNumber }
  ];
  
  const results = {
    total: testFunctions.length,
    success: 0,
    failure: 0,
    error: 0,
    details: {}
  };
  
  for (const test of testFunctions) {
    Logger.log(`----- ${test.name} 実行開始 -----`);
    
    try {
      const result = test.func();
      results.details[test.name] = result;
      
      if (result === "成功") {
        results.success++;
      } else if (result === "失敗") {
        results.failure++;
      } else {
        results.error++;
      }
      
      Logger.log(`----- ${test.name} 実行結果: ${result} -----\n`);
    } catch (error) {
      Logger.log(`----- ${test.name} 実行エラー: ${error.message} -----\n`);
      results.details[test.name] = "エラー";
      results.error++;
    }
  }
  
  // 結果サマリーを出力
  Logger.log("===== テスト結果サマリー =====");
  Logger.log(`テスト総数: ${results.total}`);
  Logger.log(`成功: ${results.success}`);
  Logger.log(`失敗: ${results.failure}`);
  Logger.log(`エラー: ${results.error}`);
  
  for (const [testName, result] of Object.entries(results.details)) {
    Logger.log(`${testName}: ${result}`);
  }
  
  // テスト環境のクリーンアップ
  cleanupAfterTests();
  
  return results;
}

/**
 * 元の値を保存する関数
 */
function saveOriginalValues(topLeftCell) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const scoreSheet = ss.getSheetByName('スコア入力');
  
  const cellRange = scoreSheet.getRange(topLeftCell);
  const row = cellRange.getRow();
  const col = cellRange.getColumn();
  
  // 保存するセル範囲（4人分のチーム表示、ID、名前、スコア）
  const values = scoreSheet.getRange(row, col, 4, 4).getValues();
  
  return {
    values: values
  };
}

/**
 * 元の値を復元する関数
 */
function restoreOriginalValues(topLeftCell, originalValues) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const scoreSheet = ss.getSheetByName('スコア入力');
  
  const cellRange = scoreSheet.getRange(topLeftCell);
  const row = cellRange.getRow();
  const col = cellRange.getColumn();
  
  // 値を復元
  scoreSheet.getRange(row, col, 4, 4).setValues(originalValues.values);
}
