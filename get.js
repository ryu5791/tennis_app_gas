/**
* ゲームデータの収集と検証を管理するクラス
*/
class GameCollector {
 /**
  * コンストラクタ
  */
 constructor() {
   // 会員データのベーステンプレート
   this.baseRecord = {
     date: null,
     gameNo: null,
     ID: null,
     pairID: null,
     serve1st: "",
     serve2nd: "",
     gamePt: null,
     serveTurn: "",
     row: null
   };
   
   // 登録用バッファ（成功したゲームデータを蓄積する）
   this.buffer = [];
   
   // 処理結果フラグ
   this.success = false;
 }
 
 /**
  * バッファ内の指定日付の最大ゲーム番号を取得する
  * @param {Date|string} gameDate - 検索する日付
  * @return {number} - 最大ゲーム番号（該当ゲームがない場合は0）
  */
 getBufferMaxGameNumber(gameDate) {
   if (!this.buffer || this.buffer.length === 0) {
     return 0;
   }
   
   // 日付をYYYY-MM-DD形式に変換して比較
   const formattedDate = Utilities.formatDate(new Date(gameDate), Session.getScriptTimeZone(), "yyyy-MM-dd");
   
   // バッファ内の指定日付のゲームをフィルタリングし、最大ゲーム番号を取得
   let maxGameNo = 0;
   
   for (const record of this.buffer) {
     if (record.date) {
       const recordDateStr = Utilities.formatDate(new Date(record.date), Session.getScriptTimeZone(), "yyyy-MM-dd");
       if (recordDateStr === formattedDate && record.gameNo > maxGameNo) {
         maxGameNo = record.gameNo;
       }
     }
   }
   
   return maxGameNo;
 }
 
 /**
  * 1ゲームのデータが正しく入力されているかをチェックするメソッド
  * @param {string} topLeftCell - チェックするゲームの左上セル (例: "B2")
  * @param {Date|string|null} date - デフォルトの日付 (指定がなければnull)
  * @return {boolean} - 検証成功時はtrue、失敗時はfalse
  */
 getOneGame(topLeftCell, date = null) {
   // スプレッドシート取得
   const ss = SpreadsheetApp.getActiveSpreadsheet();
   const scoreSheet = ss.getSheetByName('スコア入力');
   const masterSheet = ss.getSheetByName('会員マスター');
   
   // 左上セルの位置を取得
   const cellRange = scoreSheet.getRange(topLeftCell);
   const row = cellRange.getRow();
   const col = cellRange.getColumn();
   
   // 行のオフセット定数
   const TEAM_A_MEMBER1_OFFSET = 0;
   const TEAM_A_MEMBER2_OFFSET = 1;
   const TEAM_B_MEMBER1_OFFSET = 2;
   const TEAM_B_MEMBER2_OFFSET = 3;
   
   // 処理結果を初期化
   this.success = false;
   
   // ============ 判定基準1: 日付チェック ============
   // 左上セルから日付取得（OFFSET_DATE_POSITIONを適用）
   const dateOffsetRow = SheetInfo.OFFSET_DATE_POSITION[0];
   const dateOffsetCol = SheetInfo.OFFSET_DATE_POSITION[1];
   const gameDate = scoreSheet.getRange(row + dateOffsetRow, col + dateOffsetCol).getValue();
   
   // 日付が入力されているか確認
   if (!gameDate && !date) {
//     UIHelper.showAlert("日付が入力されていません。");
     return false;
   }
   
   // 日付を設定
   const gameFormattedDate = gameDate || date;
   
   // ============ IDと得点データの一括取得 ============
   // 名前列、ID列、スコア列の位置を計算
   const nameCol = col + SheetInfo.OFFSET_COL_NAME;
   const idCol = col + SheetInfo.OFFSET_COL_ID;
   const scoreCol = col + SheetInfo.OFFSET_COL_POINT;
   
   // 4人分のデータを一度に取得（4行 × 3列: 名前、ID、スコア）
   const values = scoreSheet.getRange(row, nameCol, 4, scoreCol - nameCol + 1).getValues();
   
   // 選手データを取得
   const memberA1Name = values[TEAM_A_MEMBER1_OFFSET][0];
   const memberA1Id = values[TEAM_A_MEMBER1_OFFSET][1];
   const memberA2Name = values[TEAM_A_MEMBER2_OFFSET][0];
   const memberA2Id = values[TEAM_A_MEMBER2_OFFSET][1];
   const scoreA = values[TEAM_A_MEMBER1_OFFSET][2];
   
   const memberB1Name = values[TEAM_B_MEMBER1_OFFSET][0];
   const memberB1Id = values[TEAM_B_MEMBER1_OFFSET][1];
   const memberB2Name = values[TEAM_B_MEMBER2_OFFSET][0];
   const memberB2Id = values[TEAM_B_MEMBER2_OFFSET][1];
   const scoreB = values[TEAM_B_MEMBER1_OFFSET][2];
   
   // 全IDの配列を作成（空のIDは除外）
   const allIds = [memberA1Id, memberA2Id, memberB1Id, memberB2Id].filter(id => id);
   
   // 各チームに2人ずついるか確認（厳密なダブルス）
   if (
     (!memberA1Id || !memberA2Id) || 
     (!memberB1Id || !memberB2Id)
   ) {
     Logger.log(`各チームに2人の選手が必要です: ${topLeftCell}`);
     return false;
   }
   
   // ============ 判定基準2: IDが会員マスターに登録されているかチェック ============
   // 会員マスターからIDを取得
   const masterData = masterSheet.getDataRange().getValues();
   const validIds = masterData.slice(1).map(row => row[1].toString().trim()).filter(id => id);
   
   // IDと名前のマッピングを作成
   const playerIdToName = {
     [memberA1Id]: memberA1Name,
     [memberA2Id]: memberA2Name,
     [memberB1Id]: memberB1Name,
     [memberB2Id]: memberB2Id
   };
   
   // すべてのIDが会員マスターに登録されているか確認
   for (const id of allIds) {
     if (!validIds.includes(id.toString())) {
       Logger.log(`無効なID: ${id} (${topLeftCell}) 選手名: ${playerIdToName[id] || '不明'}`);
       return false;
     }
   }
   
   // ============ 判定基準3: IDの重複チェック ============
   // 重複を除いたIDの配列の長さと元の配列の長さを比較
   const uniqueIds = [...new Set(allIds)];
   if (uniqueIds.length !== allIds.length) {
     Logger.log(`IDが重複しています: ${topLeftCell}`);
     return false;
   }
   
   // ============ 判定基準4: スコアのチェック ============
   // スコアを数値に変換
   const numScoreA = Number(scoreA);
   const numScoreB = Number(scoreB);
   
   // 少なくとも一方のスコアがある必要がある
   if (isNaN(numScoreA) && isNaN(numScoreB)) {
     Logger.log(`スコアが入力されていません: ${topLeftCell}`);
     return false;
   }
   
   // スコアチェック: 片方が5でもう片方が0～3であることを確認
   if (
     (numScoreA === 5 && (numScoreB < 0 || numScoreB > 3)) ||
     (numScoreB === 5 && (numScoreA < 0 || numScoreA > 3)) ||
     (numScoreA !== 5 && numScoreB !== 5)
   ) {
     Logger.log(`スコアが正しくありません: A=${numScoreA}, B=${numScoreB} (${topLeftCell})`);
     return false;
   }
   
   // 最大ゲーム番号を取得
   const gameCounter = getMaxGameNumber(gameFormattedDate) + 1;
   
   // すべてのチェックを通過したらデータを設定
   const gameData = [
     {
       date: gameFormattedDate,
       gameNo: gameCounter,
       ID: memberA1Id,
       pairID: memberA2Id || "",
       gamePt: scoreA,
       row: 1
     },
     {
       date: gameFormattedDate,
       gameNo: gameCounter,
       ID: memberA2Id,
       pairID: memberA1Id || "",
       gamePt: scoreA,
       row: 2
     },
     {
       date: gameFormattedDate,
       gameNo: gameCounter,
       ID: memberB1Id,
       pairID: memberB2Id || "",
       gamePt: scoreB,
       row: 3
     },
     {
       date: gameFormattedDate,
       gameNo: gameCounter,
       ID: memberB2Id,
       pairID: memberB1Id || "",
       gamePt: scoreB,
       row: 4
     }
   ];
   
   // 成功フラグを設定
   this.success = true;
   
   // 登録用バッファにデータを追加
   this.buffer = this.buffer.concat(gameData);
   
   return true;
 }
 
 /**
  * 1シート分のゲームデータを収集する
  * @param {string} topLeftCell - シートの左上セル位置 (例: "B2")
  * @param {Date|string|null} previousSheetDate - 前のシートの有効な日付（nullの場合は1枚目）
  * @return {Object} - 処理結果情報
  */
 getOneSheet(topLeftCell, previousSheetDate = null) {
   // アクティブなスプレッドシート取得
   const ss = SpreadsheetApp.getActiveSpreadsheet();
   const sheet = ss.getActiveSheet();
   const masterSheet = ss.getSheetByName('会員マスター');
   
   // セル参照をA1形式からrow,columnに変換
   const cellA1Notation = topLeftCell.toUpperCase();
   const cell = sheet.getRange(cellA1Notation);
   const startRow = cell.getRow();
   const startCol = cell.getColumn();
   
   // バッファをクリアしない（すでに親関数でクリアされているため）
   // this.buffer = [];
   
   // 処理結果を格納するオブジェクト
   const result = {
     totalGames: SheetInfo.positions.length,
     successCount: 0,
     failedCount: 0,
     failedGames: [],
     successGames: [],  
     gameDetails: [],  // ゲームの詳細情報を格納
     message: "",      // メッセージを追加
     validDate: null   // このシートで有効となった日付を返す
   };
   
  // 会員マスターからIDと名前のマッピングを作成
  const masterData = masterSheet.getDataRange().getValues();
  const idToName = {};
  for (let i = 1; i < masterData.length; i++) {
    if (masterData[i][1]) { // IDが存在する場合
      // dispName(列2)ではなく、name(列0)を使用
      idToName[masterData[i][1].toString()] = masterData[i][0] || masterData[i][1]; // nameがなければIDを使用
    }
  }   
   // SheetInfoクラスから試合位置テーブルを取得
   const gamePositions = SheetInfo.positions;
   
   // 現在のページ情報を取得
   const currentPageInfo = SheetInfo.pageInfo.find(page => page.position === topLeftCell);
   const startGameNumber = currentPageInfo ? currentPageInfo.startGameNo : 1;
   
   // 日付の初期化 - 前のシートの日付を使用
   let currentDate = previousSheetDate;
   
   // 1枚目のシートで日付がない場合はエラー
   if (previousSheetDate === null) {
     // 1枚目のシートの日付セルをチェック
     const firstGameDateCell = sheet.getRange(startRow + SheetInfo.OFFSET_DATE_POSITION[0], 
                                             startCol + SheetInfo.OFFSET_DATE_POSITION[1]);
     const firstGameDate = firstGameDateCell.getValue();
     
     if (!firstGameDate) {
       result.message = "1枚目のシートに日付が入力されていません。";
       return result;
     }
     currentDate = firstGameDate;
   }
   
   // シート内の全ゲームを処理
   for (let i = 0; i < gamePositions.length; i++) {
     // 実際のゲーム番号を計算
     const actualGameNumber = startGameNumber + i;
     
     // 相対位置を取得
     const rowOffset = gamePositions[i][0];
     const colOffset = gamePositions[i][1];
     
     // 実際のセル位置を計算
     const gameRow = startRow + rowOffset;
     const gameCol = startCol + colOffset;
     
     // 現在の試合の開始セルをA1形式で取得
     const gameTopLeftCell = sheet.getRange(gameRow, gameCol).getA1Notation();
     
     // ゲームデータが空でないかチェック
     const checkRange = sheet.getRange(gameRow, gameCol, 4, 4);
     const checkValues = checkRange.getValues();
     let hasData = false;
     
     // 4行4列のデータのいずれかに値があるかをチェック
     for (let r = 0; r < 4; r++) {
       for (let c = 0; c < 4; c++) {
         if (checkValues[r][c]) {
           hasData = true;
           break;
         }
       }
       if (hasData) break;
     }
     
     // データがない場合はスキップ
     if (!hasData) {
       continue;
     }
     
     // 個別ゲームの日付をチェック（このゲームに固有の日付があるか）
     const gameDateCell = sheet.getRange(gameRow + SheetInfo.OFFSET_DATE_POSITION[0], 
                                        gameCol + SheetInfo.OFFSET_DATE_POSITION[1]);
     const gameSpecificDate = gameDateCell.getValue();
     
     // ゲーム固有の日付がある場合は更新
     if (gameSpecificDate) {
       currentDate = gameSpecificDate;
     }
     
     // バッファの現在のサイズを記録
     const bufferSizeBefore = this.buffer.length;
     
     // getOneGame関数を呼び出して1試合分の処理を実行（現在の有効な日付を渡す）
     const success = this.getOneGame(gameTopLeftCell, currentDate);
     
     // 結果を記録
     if (success) {
       result.successCount++;
       result.successGames.push(actualGameNumber);
       
       // バッファから追加されたデータを取得
       const addedData = this.buffer.slice(bufferSizeBefore);
       
       if (addedData.length > 0) {
         const gameData = addedData[0]; // 最初のレコードから情報を取得
         
         // 日付をフォーマット
         const dateStr = Utilities.formatDate(new Date(gameData.date), Session.getScriptTimeZone(), "yyyy/MM/dd");
         
         // ゲーム番号をフォーマット
         const gameNoStr = ("0" + gameData.gameNo).slice(-2);
         
         // チームAのプレイヤー情報を取得
         const teamA = addedData.filter(d => d.row <= 2);
         const teamB = addedData.filter(d => d.row > 2);
         
         // プレイヤー名とポイントを取得
         const teamAInfo = teamA.map(player => {
           const name = idToName[player.ID] || player.ID;
           return `${name}さん(${player.ID})`;
         }).join('、');
         
         const teamBInfo = teamB.map(player => {
           const name = idToName[player.ID] || player.ID;
           return `${name}さん(${player.ID})`;
         }).join('、');
         
         const gamePtA = teamA[0].gamePt;
         const gamePtB = teamB[0].gamePt;
         
         result.gameDetails.push({
           gameNumber: actualGameNumber,
           dateStr: dateStr,
           gameNoStr: gameNoStr,
           teamAInfo: teamAInfo,
           teamAPoints: gamePtA,
           teamBInfo: teamBInfo,
           teamBPoints: gamePtB,
           success: true
         });
       }
     } else {
       result.failedCount++;
       result.failedGames.push({
         gameNumber: actualGameNumber,
         cellReference: gameTopLeftCell
       });
       
       // エラー理由を特定するためにログを確認
       let errorReason = "エラー";
       
       // 日付チェック
       if (!currentDate) {
         errorReason = "日付が入力されていません";
       } else {
         // その他のエラーの場合、ID重複やスコア不正などをチェック
         // ここでは簡略化して、主なエラーメッセージを設定
         if (Logger.getLog().includes("IDが重複しています")) {
           errorReason = "IDが重複しています";
         } else if (Logger.getLog().includes("スコアが正しくありません")) {
           errorReason = "スコアが不正です";
         } else if (Logger.getLog().includes("各チームに2人の選手が必要です")) {
           errorReason = "各チームに2人の選手が必要です";
         }
       }
       
       result.gameDetails.push({
         gameNumber: actualGameNumber,
         success: false,
         errorReason: errorReason
       });
     }
   }
   
   // 有効な日付を結果に設定
   result.validDate = currentDate;
   
   // 処理結果のサマリーをログに出力
   Logger.log(`データ収集完了: 成功=${result.successCount}, 失敗=${result.failedCount}, バッファサイズ=${this.buffer.length}`);
   
   // 結果メッセージを構築
   if (result.gameDetails.length > 0) {
     let message = "";
     
     result.gameDetails.forEach(game => {
       if (game.success) {
         message += `第${game.gameNumber}ゲーム：${game.dateStr} - No.${game.gameNoStr}\n`;
         message += `　${game.teamAPoints}pt：${game.teamAInfo}\n`;
         message += `　${game.teamBPoints}pt：${game.teamBInfo}\n`;
       } else {
         message += `第${game.gameNumber}ゲーム：NG（${game.errorReason}）\n`;
       }
     });
     
     result.message = message;
   } else {
     result.message = "処理対象のゲームがありませんでした。";
   }
   
   return result;
 }
 
 /**
  * バッファをクリアする
  */
 clearBuffer() {
   this.buffer = [];
   this.success = false;
 }
 
 /**
  * バッファの内容を取得する
  * @return {Array} - バッファに格納されているデータ
  */
 getBuffer() {
   return this.buffer;
 }
}

function getAllGame()
{
 try {
   // GameCollectorインスタンスが存在しない場合は作成
   if (!gameCollectorInstance) {
     gameCollectorInstance = new GameCollector();
   }
   
   // インスタンスをクリアする代わりに、バッファのみクリア
   gameCollectorInstance.clearBuffer();
   
   let fullMessage = "";
   let hasData = false;
   let totalSuccessCount = 0;
   let totalFailedCount = 0;
   let currentValidDate = null; // 現在有効な日付を追跡
   
   // 各ページを処理
   SheetInfo.pageInfo.forEach((pageInfo, index) => {
     // 1枚目はnullを渡し、2枚目以降は前のシートの有効日付を渡す
     const previousDate = index === 0 ? null : currentValidDate;
     
     const result = gameCollectorInstance.getOneSheet(pageInfo.position, previousDate);
     
     // 1枚目で日付がない場合のエラーハンドリング
     if (index === 0 && result.message === "1枚目のシートに日付が入力されていません。") {
       UIHelper.showAlert(result.message);
       return;
     }
     
     // 有効な日付を更新（次のシートで使用）
     if (result.validDate) {
       currentValidDate = result.validDate;
     }
     
     if (result.message && result.message !== "処理対象のゲームがありませんでした。") {
       fullMessage += `=== ${pageInfo.pageName} ===\n`;
       fullMessage += result.message;
       fullMessage += "\n";
       hasData = true;
       totalSuccessCount += result.successCount;
       totalFailedCount += result.failedCount;
     }
   });
   
   // メッセージがない場合
   if (!fullMessage) {
     fullMessage = "処理対象のゲームがありませんでした。";
     UIHelper.showAlert(fullMessage);
     return;
   }
   
   // データがある場合、保存するか確認
   if (hasData) {
     // サマリーを追加
     fullMessage += `\n=== サマリー ===\n`;
     fullMessage += `成功: ${totalSuccessCount}件\n`;
     fullMessage += `失敗: ${totalFailedCount}件\n`;
     fullMessage += `\n保存しますか？`;
     
     // UIを使用してYES/NOダイアログを表示
     const ui = SpreadsheetApp.getUi();
     const response = ui.alert(
       'データ保存の確認',
       fullMessage,
       ui.ButtonSet.YES_NO
     );
     
     // YESが選択された場合、データを保存
     if (response === ui.Button.YES) {
       const saveResult = saveBufferData();
       if (saveResult.success) {
         UIHelper.showAlert(`データを保存しました。\n保存件数: ${saveResult.savedCount}件`);
       } else {
         UIHelper.showAlert(`保存中にエラーが発生しました: ${saveResult.error}`);
       }
     } else {
       UIHelper.showAlert("保存をキャンセルしました。");
     }
   }
 } catch (error) {
   UIHelper.showAlert(`エラーが発生しました: ${error.message}`);
   Logger.log(`Error in getAllGame: ${error.message}`);
   Logger.log(`Stack trace: ${error.stack}`);
 }
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

/**
* 指定した日付の最大ゲーム番号を取得する関数
* まずバッファ内を検索し、なければDBを検索する
* @param {Date|string} gameDate - 検索する日付
* @return {number} - 指定日付の最大ゲーム番号（該当ゲームがない場合は0）
*/
function getMaxGameNumber(gameDate) {
 // GameCollectorのインスタンスが存在するかチェック
 if (typeof gameCollectorInstance === 'undefined' || gameCollectorInstance === null) {
   // インスタンスがなければDBから直接取得
   return getDBMaxGameNumber(gameDate);
 }
 
 // バッファから最大ゲーム番号を取得
 return gameCollectorInstance.getBufferMaxGameNumber(gameDate);
 
/* const bufferMaxNo = gameCollectorInstance.getBufferMaxGameNumber(gameDate);
 
 // DBから最大ゲーム番号を取得
 const dbMaxNo = getDBMaxGameNumber(gameDate);
 
 // バッファとDBの最大値を比較して大きい方を返す
 return Math.max(bufferMaxNo, dbMaxNo);*/
}

/**
* バッファのデータをスコア出力シートに保存する関数
* @return {Object} - 保存結果 {success: boolean, savedCount: number, error: string}
*/
function saveBufferData() {
 try {
   // GameCollectorインスタンスの存在確認
   if (!gameCollectorInstance || !gameCollectorInstance.buffer || gameCollectorInstance.buffer.length === 0) {
     return {
       success: false,
       savedCount: 0,
       error: "保存するデータがありません"
     };
   }
   
   const ss = SpreadsheetApp.getActiveSpreadsheet();
   const outputSheet = ss.getSheetByName('スコア出力');
   
   // ヘッダー行がない場合は追加
   if (outputSheet.getLastRow() == 0) {
     outputSheet.appendRow(['date', 'gameNo', 'ID', 'pairID', 'serve1st', 'serve2nd', 'gamePt', 'serveTurn', 'row']);
   }
   
   // バッファのデータを取得
   const buffer = gameCollectorInstance.getBuffer();
   let savedCount = 0;
   
   // バッファの各レコードを出力シートに追加
   buffer.forEach(record => {
     // 日付をYYYY-MM-DD形式に変換
     const formattedDate = Utilities.formatDate(new Date(record.date), Session.getScriptTimeZone(), "yyyy-MM-dd");
     
     outputSheet.appendRow([
       formattedDate,
       record.gameNo,
       record.ID,
       record.pairID || "",
       record.serve1st || "",
       record.serve2nd || "",
       record.gamePt,
       record.serveTurn || "",
       record.row
     ]);
     
     savedCount++;
   });
   
   // バッファをクリア
   gameCollectorInstance.clearBuffer();
   
   return {
     success: true,
     savedCount: savedCount,
     error: ""
   };
   
 } catch (error) {
   Logger.log(`保存中にエラーが発生しました: ${error.message}`);
   return {
     success: false,
     savedCount: 0,
     error: error.message
   };
 }
}

// グローバルなGameCollectorインスタンス
if (typeof gameCollectorInstance === 'undefined' || gameCollectorInstance === null) {
 gameCollectorInstance = new GameCollector();
}