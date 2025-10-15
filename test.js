/**
* clearOneGame関数の単体テストを行う関数
*/
function test_clearOneGame() {
 // テスト対象のセル位置（新レイアウトでは列G、行8から開始の2試合目）
 const testCell = "G8";
 
 // テストを実行
 try {
  // "スコア入力"シート取得
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('スコア入力');
   
   // 元のセルの値を保存
   const originalFormulas = {
     nameFormulas: [],
     idFormulas: [],
     upperScoreFormula: "",
     lowerScoreFormula: ""
   };
   
   for (let i = 0; i < 4; i++) {
     const row = 8 + i;
     originalFormulas.nameFormulas.push(sheet.getRange(`I${row}`).getFormula());
     originalFormulas.idFormulas.push(sheet.getRange(`H${row}`).getFormula());
   }
   
   originalFormulas.upperScoreFormula = sheet.getRange("J8").getFormula();
   originalFormulas.lowerScoreFormula = sheet.getRange("J10").getFormula();
   
   // clearOneGame実行
   clearOneGame(testCell);
   
   // 期待される数式
   const expectedFormulas = {
     nameFormulas: [
       '=IFERROR(VLOOKUP(H8,\'会員マスター\'!B:C,2,FALSE),"")',
       '=IFERROR(VLOOKUP(H9,\'会員マスター\'!B:C,2,FALSE),"")',
       '=IFERROR(VLOOKUP(H10,\'会員マスター\'!B:C,2,FALSE),"")',
       '=IFERROR(VLOOKUP(H11,\'会員マスター\'!B:C,2,FALSE),"")'
     ],
     idFormulas: [
       '=IFERROR(IFERROR(INDEX(\'会員マスター\'!B:B, MATCH(I8, \'会員マスター\'!C:C, 0)), INDEX(\'会員マスター\'!B:B, MATCH(I8, \'会員マスター\'!D:D, 0))),"")',
       '=IFERROR(IFERROR(INDEX(\'会員マスター\'!B:B, MATCH(I9, \'会員マスター\'!C:C, 0)), INDEX(\'会員マスター\'!B:B, MATCH(I9, \'会員マスター\'!D:D, 0))),"")',
       '=IFERROR(IFERROR(INDEX(\'会員マスター\'!B:B, MATCH(I10, \'会員マスター\'!C:C, 0)), INDEX(\'会員マスター\'!B:B, MATCH(I10, \'会員マスター\'!D:D, 0))),"")',
       '=IFERROR(IFERROR(INDEX(\'会員マスター\'!B:B, MATCH(I11, \'会員マスター\'!C:C, 0)), INDEX(\'会員マスター\'!B:B, MATCH(I11, \'会員マスター\'!D:D, 0))),"")'
     ],
     upperScoreFormula: '=IF(OR(J10=0, J10=1, J10=2, J10=3), 5, "")',
     lowerScoreFormula: '=IF(OR(J8=0, J8=1, J8=2, J8=3), 5, "")'
   };
   
   // 結果を検証
   let allPassed = true;
   const results = [];
   
   // 会員名とID式の検証
   for (let i = 0; i < 4; i++) {
     const row = 8 + i;
     const actualNameFormula = sheet.getRange(`I${row}`).getFormula();
     const actualIdFormula = sheet.getRange(`H${row}`).getFormula();
     
     const nameMatch = actualNameFormula === expectedFormulas.nameFormulas[i];
     const idMatch = actualIdFormula === expectedFormulas.idFormulas[i];
     
     if (!nameMatch || !idMatch) {
       allPassed = false;
       results.push(`行 ${row}: ${nameMatch ? '✓' : '✗'} 会員名式, ${idMatch ? '✓' : '✗'} ID式`);
       
       if (!nameMatch) {
         results.push(`  期待: ${expectedFormulas.nameFormulas[i]}`);
         results.push(`  実際: ${actualNameFormula}`);
       }
       
       if (!idMatch) {
         results.push(`  期待: ${expectedFormulas.idFormulas[i]}`);
         results.push(`  実際: ${actualIdFormula}`);
       }
     }
   }
   
   // 上段ポイント式の検証
   const actualUpperScoreFormula = sheet.getRange("J8").getFormula();
   const upperScoreMatch = actualUpperScoreFormula === expectedFormulas.upperScoreFormula;
   
   if (!upperScoreMatch) {
     allPassed = false;
     results.push("上段ポイント式: ✗");
     results.push(`  期待: ${expectedFormulas.upperScoreFormula}`);
     results.push(`  実際: ${actualUpperScoreFormula}`);
   }
   
   // 下段ポイント式の検証
   const actualLowerScoreFormula = sheet.getRange("J10").getFormula();
   const lowerScoreMatch = actualLowerScoreFormula === expectedFormulas.lowerScoreFormula;
   
   if (!lowerScoreMatch) {
     allPassed = false;
     results.push("下段ポイント式: ✗");
     results.push(`  期待: ${expectedFormulas.lowerScoreFormula}`);
     results.push(`  実際: ${actualLowerScoreFormula}`);
   }
   
   // 結果を出力
   if (allPassed) {
     Logger.log("成功: すべてのセルに正しい数式が設定されました。");
     return "成功";
   } else {
     Logger.log("失敗: 以下のセルで期待値と異なる数式が設定されました:");
     results.forEach(result => Logger.log(result));
     return "失敗";
   }
 } catch (error) {
   Logger.log(`テスト実行時にエラーが発生しました: ${error.message}`);
   return "失敗";
 }
}

/**
* clearOnePage関数の単体テストを行う関数
*/
function test_clearOnePage() {
 // テスト対象のセル位置（1ページ目の開始位置）
 const testCell = "B3";
 
 // テストを実行
 try {
  // "スコア入力"シート取得
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('スコア入力');
   
   // 試合位置テーブルを取得
   const gamePositions = SheetInfo.positions;
   
   // 元のセルの値を保持する配列
   const originalFormulas = [];
   
   // 開始位置を取得
   const startCell = sheet.getRange(testCell);
   const startRow = startCell.getRow();
   const startCol = startCell.getColumn();
   
   // 各試合の元の値を保存
   for (let gameIndex = 0; gameIndex < gamePositions.length; gameIndex++) {
     const rowOffset = gamePositions[gameIndex][0];
     const colOffset = gamePositions[gameIndex][1];
     
     const gameStartRow = startRow + rowOffset;
     const gameStartCol = startCol + colOffset;
     
     const gameFormulas = {
       gameIndex: gameIndex + 1,
       nameFormulas: [],
       idFormulas: [],
       upperScoreFormula: '',
       lowerScoreFormula: ''
     };
     
     // 4行分の会員名とID式を保存
     for (let i = 0; i < 4; i++) {
       const currentRow = gameStartRow + i;
       gameFormulas.nameFormulas.push(sheet.getRange(currentRow, gameStartCol + SheetInfo.OFFSET_COL_NAME).getFormula());
       gameFormulas.idFormulas.push(sheet.getRange(currentRow, gameStartCol + SheetInfo.OFFSET_COL_ID).getFormula());
     }
     
     // 上段と下段のスコア式を保存
     gameFormulas.upperScoreFormula = sheet.getRange(gameStartRow, gameStartCol + SheetInfo.OFFSET_COL_POINT).getFormula();
     gameFormulas.lowerScoreFormula = sheet.getRange(gameStartRow + 2, gameStartCol + SheetInfo.OFFSET_COL_POINT).getFormula();
     
     originalFormulas.push(gameFormulas);
   }
   
   // clearOnePage実行
   clearOnePage(testCell);
   
   // 結果を検証
   let allPassed = true;
   const results = [];
   
   // 各試合の数式を検証
   for (let gameIndex = 0; gameIndex < gamePositions.length; gameIndex++) {
     const rowOffset = gamePositions[gameIndex][0];
     const colOffset = gamePositions[gameIndex][1];
     
     const gameStartRow = startRow + rowOffset;
     const gameStartCol = startCol + colOffset;
     
     const gameNumber = gameIndex + 1;
     let gameTestPassed = true;
     const gameResults = [];
     
     // 4行分の会員名とID式を検証
     for (let i = 0; i < 4; i++) {
       const currentRow = gameStartRow + i;
       
       // 会員名セルの検証
       const nameCell = sheet.getRange(currentRow, gameStartCol + SheetInfo.OFFSET_COL_NAME);
       const actualNameFormula = nameCell.getFormula();
       const idCellRef = sheet.getRange(currentRow, gameStartCol + SheetInfo.OFFSET_COL_ID).getA1Notation();
       const expectedNameFormula = `=IFERROR(VLOOKUP(${idCellRef},'会員マスター'!B:C,2,FALSE),"")`;
       
       // ID列セルの検証
       const idCell = sheet.getRange(currentRow, gameStartCol + SheetInfo.OFFSET_COL_ID);
       const actualIdFormula = idCell.getFormula();
       const nameCellRef = sheet.getRange(currentRow, gameStartCol + SheetInfo.OFFSET_COL_NAME).getA1Notation();
       const expectedIdFormula = `=IFERROR(IFERROR(INDEX('会員マスター'!B:B, MATCH(${nameCellRef}, '会員マスター'!C:C, 0)), INDEX('会員マスター'!B:B, MATCH(${nameCellRef}, '会員マスター'!D:D, 0))),"")`;
       
       const nameMatch = (actualNameFormula === expectedNameFormula);
       const idMatch = (actualIdFormula === expectedIdFormula);
       
       if (!nameMatch || !idMatch) {
         gameTestPassed = false;
         gameResults.push(`  行 ${currentRow}: ${nameMatch ? '✓' : '✗'} 会員名式, ${idMatch ? '✓' : '✗'} ID式`);
         
         if (!nameMatch) {
           gameResults.push(`    期待: ${expectedNameFormula}`);
           gameResults.push(`    実際: ${actualNameFormula}`);
         }
         
         if (!idMatch) {
           gameResults.push(`    期待: ${expectedIdFormula}`);
           gameResults.push(`    実際: ${actualIdFormula}`);
         }
       }
     }
     
     // 上段スコアセルの検証
     const upperScoreCell = sheet.getRange(gameStartRow, gameStartCol + SheetInfo.OFFSET_COL_POINT);
     const actualUpperScoreFormula = upperScoreCell.getFormula();
     const lowerCellRef = sheet.getRange(gameStartRow + 2, gameStartCol + SheetInfo.OFFSET_COL_POINT).getA1Notation();
     const expectedUpperScoreFormula = `=IF(OR(${lowerCellRef}=0, ${lowerCellRef}=1, ${lowerCellRef}=2, ${lowerCellRef}=3), 5, "")`;
     
     // 下段スコアセルの検証
     const lowerScoreCell = sheet.getRange(gameStartRow + 2, gameStartCol + SheetInfo.OFFSET_COL_POINT);
     const actualLowerScoreFormula = lowerScoreCell.getFormula();
     const upperCellRef = sheet.getRange(gameStartRow, gameStartCol + SheetInfo.OFFSET_COL_POINT).getA1Notation();
     const expectedLowerScoreFormula = `=IF(OR(${upperCellRef}=0, ${upperCellRef}=1, ${upperCellRef}=2, ${upperCellRef}=3), 5, "")`;
     
     const upperScoreMatch = (actualUpperScoreFormula === expectedUpperScoreFormula);
     const lowerScoreMatch = (actualLowerScoreFormula === expectedLowerScoreFormula);
     
     if (!upperScoreMatch || !lowerScoreMatch) {
       gameTestPassed = false;
       
       if (!upperScoreMatch) {
         gameResults.push(`  上段スコア式: ✗`);
         gameResults.push(`    期待: ${expectedUpperScoreFormula}`);
         gameResults.push(`    実際: ${actualUpperScoreFormula}`);
       }
       
       if (!lowerScoreMatch) {
         gameResults.push(`  下段スコア式: ✗`);
         gameResults.push(`    期待: ${expectedLowerScoreFormula}`);
         gameResults.push(`    実際: ${actualLowerScoreFormula}`);
       }
     }
     
     if (!gameTestPassed) {
       allPassed = false;
       results.push(`試合 ${gameNumber} (セル ${sheet.getRange(gameStartRow, gameStartCol).getA1Notation()}): 失敗`);
       gameResults.forEach(result => results.push(result));
     }
   }
   
   // 結果を出力
   if (allPassed) {
     Logger.log("成功: 28試合すべてのセルに正しい数式が設定されました。");
     return "成功";
   } else {
     Logger.log("失敗: 以下の試合で期待値と異なる数式が設定されました:");
     results.forEach(result => Logger.log(result));
     return "失敗";
   }
   
 } catch (error) {
   Logger.log(`テスト実行時にエラーが発生しました: ${error.message}`);
   return "失敗";
 }
}

/**
* getMaxGameNumber関数の単体テスト
*/
function test_getMaxGameNumber() {
 // テスト対象の日付
 const testDate = "2025-03-31";
 
 // 期待値
 const expectedMaxGameNo = 1;
 
 try {
   // getMaxGameNumber関数を実行
   const actualMaxGameNo = getMaxGameNumber(testDate);
   
   // 結果を検証
   if (actualMaxGameNo === expectedMaxGameNo) {
     Logger.log(`成功: 日付 ${testDate} の最大ゲーム番号は ${actualMaxGameNo} です。`);
     return "成功";
   } else {
     Logger.log(`失敗: 日付 ${testDate} の最大ゲーム番号は ${expectedMaxGameNo} であるべきですが、${actualMaxGameNo} が返されました。`);
     return "失敗";
   }
 } catch (error) {
   Logger.log(`テスト実行時にエラーが発生しました: ${error.message}`);
   return "失敗";
 }
}

/**
 * PropertiesServiceを使ってデータを保存するテスト関数
 */
function test_saveDataToProperties() {
  // テスト用のデータオブジェクト
  const testData = {
    name: "テストデータ",
    date: new Date().toISOString(),
    values: [1, 2, 3, 4, 5],
    nestedObject: {
      key1: "value1",
      key2: "value2"
    }
  };
  
  try {
    // データをスクリプトプロパティに保存
    PropertiesService.getScriptProperties().setProperty("myTestData", JSON.stringify(testData));
    Logger.log("データを保存しました: " + JSON.stringify(testData));
    return "保存成功";
  } catch (error) {
    Logger.log("データの保存に失敗しました: " + error.message);
    return "保存失敗: " + error.message;
  }
}

/**
 * PropertiesServiceから保存したデータを取得するテスト関数
 */
function test_retrieveDataFromProperties() {
  try {
    // スクリプトプロパティからデータを取得
    const savedDataString = PropertiesService.getScriptProperties().getProperty("myTestData");
    
    if (!savedDataString) {
      Logger.log("保存されたデータがありません。まずtest_saveDataToPropertiesを実行してください。");
      return "データなし";
    }
    
    // JSON文字列をオブジェクトに変換
    const savedData = JSON.parse(savedDataString);
    
    // 取得したデータの内容を確認
    Logger.log("取得したデータ:");
    Logger.log("名前: " + savedData.name);
    Logger.log("日付: " + savedData.date);
    Logger.log("値の配列: " + savedData.values);
    Logger.log("ネストしたオブジェクト: " + JSON.stringify(savedData.nestedObject));
    
    return "取得成功: " + savedData.name;
  } catch (error) {
    Logger.log("データの取得に失敗しました: " + error.message);
    return "取得失敗: " + error.message;
  }
}

/**
 * スクリプトプロパティのデータを完全に確認するテスト関数
 */
function test_validatePropertiesData() {
  try {
    // まず新しいテストデータを保存
    const originalData = {
      testNumber: 42,
      testString: "テスト文字列",
      testArray: [10, 20, 30],
      testObject: { a: 1, b: 2 },
      testDate: new Date().toISOString()
    };
    
    // データを保存
    PropertiesService.getScriptProperties().setProperty("validationTestData", JSON.stringify(originalData));
    Logger.log("元のデータを保存しました");
    
    // 保存したデータを取得
    const retrievedDataString = PropertiesService.getScriptProperties().getProperty("validationTestData");
    const retrievedData = JSON.parse(retrievedDataString);
    
    // データの内容を比較
    let isValid = true;
    const validationResults = [];
    
    // 各プロパティを検証
    for (const key in originalData) {
      if (typeof originalData[key] === 'object' && !Array.isArray(originalData[key])) {
        // オブジェクトの場合は文字列化して比較
        const original = JSON.stringify(originalData[key]);
        const retrieved = JSON.stringify(retrievedData[key]);
        
        if (original !== retrieved) {
          isValid = false;
          validationResults.push(`${key}: 不一致 - 元: ${original}, 取得: ${retrieved}`);
        } else {
          validationResults.push(`${key}: 一致`);
        }
      } else {
        // プリミティブ値または配列の場合
        const original = Array.isArray(originalData[key]) ? 
                        JSON.stringify(originalData[key]) : originalData[key];
        const retrieved = Array.isArray(retrievedData[key]) ? 
                         JSON.stringify(retrievedData[key]) : retrievedData[key];
                         
        if (original !== retrieved) {
          isValid = false;
          validationResults.push(`${key}: 不一致 - 元: ${original}, 取得: ${retrieved}`);
        } else {
          validationResults.push(`${key}: 一致`);
        }
      }
    }
    
    // 結果をログに出力
    Logger.log("検証結果:");
    validationResults.forEach(result => Logger.log(result));
    
    if (isValid) {
      Logger.log("検証成功: すべてのデータが正しく保存・取得されました");
      return "検証成功";
    } else {
      Logger.log("検証失敗: データの不一致があります");
      return "検証失敗";
    }
    
  } catch (error) {
    Logger.log("検証中にエラーが発生しました: " + error.message);
    return "検証エラー: " + error.message;
  }
}
