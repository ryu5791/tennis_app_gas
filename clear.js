/**
 * ゲームエリアに罫線を設定する関数
 * @param {Sheet} sheet - シートオブジェクト
 * @param {number} startRow - 開始行
 * @param {number} startCol - 開始列
 */
function setBordersForGame(sheet, startRow, startCol) {
  
  // 全体の範囲（4行×4列）を取得
  const fullRange = sheet.getRange(startRow, startCol, 4, 4);
  
  // 各行の細線を設定（内部の横線）
  for (let i = 1; i < 4; i++) {
    const rowRange = sheet.getRange(startRow + i, startCol, 1, 4);
    rowRange.setBorder(true, null, null, null, null, null, "black", SpreadsheetApp.BorderStyle.SOLID);
  }
  
  // 各列の細線を設定（内部の縦線）
  for (let j = 1; j < 4; j++) {
    const colRange = sheet.getRange(startRow, startCol + j, 4, 1);
    colRange.setBorder(null, true, null, null, null, null, "black", SpreadsheetApp.BorderStyle.SOLID);
  }
  
  // 右下のセル（スコア）の罫線が消えていないことを確認
  // 必要に応じて個別に設定することも可能
  const bottomRightCell = sheet.getRange(startRow + 3, startCol + 3);
  bottomRightCell.setBorder(true, true, true, true, null, null, "black", SpreadsheetApp.BorderStyle.SOLID_THICK);

  // 日付セルの罫線をクリア
  const dateCellRow = startRow + SheetInfo.OFFSET_DATE_POSITION[0];
  const dateCellCol = startCol + SheetInfo.OFFSET_DATE_POSITION[1];
  const dateCell = sheet.getRange(dateCellRow, dateCellCol, SheetInfo.ROWS_PER_GAME);
  dateCell.setBorder(null, null, null, null, false, false);

  // 全体の外枠を太線で設定
  fullRange.setBorder(true, true, true, true, null, null, "black", SpreadsheetApp.BorderStyle.SOLID_THICK);
}

/**
* 指定したセルを起点に会員名とIDセルに数式を挿入する関数
* @param {string} topLeftCell - 開始セルの位置 (例: "A1")
*/
function clearOneGame(topLeftCell) {
  // "スコア入力"シート取得
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('スコア入力');
 
  // セル参照をA1形式からrow,columnに変換
  const cellA1Notation = topLeftCell.toUpperCase();
  const cell = sheet.getRange(cellA1Notation);
  const startRow = cell.getRow();
  const col = cell.getColumn();
  
  // 日付セルをクリア
  const dateCellRow = startRow + SheetInfo.OFFSET_DATE_POSITION[0];
  const dateCellCol = col + SheetInfo.OFFSET_DATE_POSITION[1];
  const dateCell = sheet.getRange(dateCellRow, dateCellCol);
  dateCell.clearContent();
  
  // 指定されたセルとその下3行（計4行）に処理を行う
  for (let i = 0; i < 4; i++) {
    const currentRow = startRow + i;
    
    // 会員名セルを取得
    const memberNameCell = sheet.getRange(currentRow, col + SheetInfo.OFFSET_COL_NAME);
    
    // ID列の参照を取得
    const idCellRef = sheet.getRange(currentRow, col + SheetInfo.OFFSET_COL_ID).getA1Notation();
    
    // VLOOKUP式を構築
    const formula = `=IFERROR(VLOOKUP(${idCellRef},'会員マスター'!B:C,2,FALSE),"")`;
    
    // 会員名セルに数式を設定
    memberNameCell.setFormula(formula);
    
    // ID列のセルを取得
    const memberIdCell = sheet.getRange(currentRow, col + SheetInfo.OFFSET_COL_ID);
    
    // 会員名セルの参照を取得
    const memberNameCellRef = sheet.getRange(currentRow, col + SheetInfo.OFFSET_COL_NAME).getA1Notation();
    
    // INDEX/MATCH式を構築
    const indexMatchFormula = `=IFERROR(IFERROR(INDEX('会員マスター'!B:B, MATCH(${memberNameCellRef}, '会員マスター'!C:C, 0)), INDEX('会員マスター'!B:B, MATCH(${memberNameCellRef}, '会員マスター'!D:D, 0))),"")`;
    
    // ID列のセルに数式を設定
    memberIdCell.setFormula(indexMatchFormula);
  }

  // ゲーム上段のクリア
  // 上段セルを取得
  const scoreUpperCell = sheet.getRange(startRow + SheetInfo.OFFSET_ROW_UPPOINT, col + SheetInfo.OFFSET_COL_POINT);

  // 下段セルの参照を取得
  const lowerCellRef = sheet.getRange(startRow + SheetInfo.OFFSET_ROW_LOPOINT, col + SheetInfo.OFFSET_COL_POINT).getA1Notation();

  // 上段セルに数式を設定
  scoreUpperCell.setFormula(`=IF(OR(${lowerCellRef}=0, ${lowerCellRef}=1, ${lowerCellRef}=2, ${lowerCellRef}=3), 5, "")`);

  // ゲーム下段のクリア
  // 下段セルを取得
  const scoreLowerCell = sheet.getRange(startRow + SheetInfo.OFFSET_ROW_LOPOINT, col + SheetInfo.OFFSET_COL_POINT);

  // 上段セルの参照を取得
  const upperCellRef = sheet.getRange(startRow + SheetInfo.OFFSET_ROW_UPPOINT, col + SheetInfo.OFFSET_COL_POINT).getA1Notation();

  // 下段セルに数式を設定
  scoreLowerCell.setFormula(`=IF(OR(${upperCellRef}=0, ${upperCellRef}=1, ${upperCellRef}=2, ${upperCellRef}=3), 5, "")`);
  
  // 罫線の設定
  setBordersForGame(sheet, startRow, col);
}
/**
 * ページ内の12試合分のデータに対して数式を挿入する関数
 * @param {string} topLeftCell - ページの開始セル位置 (例: "B18")
 */
function clearOnePage(topLeftCell) {
  // "スコア入力"シート取得
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('スコア入力');
  
  // セル参照をA1形式からrow,columnに変換
  const cellA1Notation = topLeftCell.toUpperCase();
  const cell = sheet.getRange(cellA1Notation);
  const startRow = cell.getRow();
  const startCol = cell.getColumn();
  
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
    
    // clearOneGame関数を呼び出して1試合分の処理を実行
    clearOneGame(gameTopLeftCell);
    
    Logger.log(`試合 ${i+1}: ${gameTopLeftCell} の処理完了`);
  }
  
  Logger.log(`12試合分の処理が完了しました。開始セル: ${topLeftCell}`);
}

/**
 * すべてのページのデータをクリアする関数（確認付き）
 */
function clearAllPage() {
  // 確認ダイアログを表示
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '確認',
    'すべてのページのデータをクリアしますか？\nこの操作は取り消せません。',
    ui.ButtonSet.YES_NO
  );
  
  // NOが選択された場合は中止
  if (response !== ui.Button.YES) {
    UIHelper.showAlert("クリアをキャンセルしました。");
    return;
  }
  
  try {
    Logger.log("clearAllPage開始");
    
    // 各ページの情報を取得
    const pageInfo = SheetInfo.pageInfo;
    
    // 処理結果を格納
    let successCount = 0;
    let failedPages = [];
    
    // 各ページを順番に処理
    pageInfo.forEach(page => {
      try {
        Logger.log(`${page.pageName}のクリアを開始: ${page.position}`);
        clearOnePage(page.position);
        Logger.log(`${page.pageName}のクリアが完了`);
        successCount++;
      } catch (pageError) {
        Logger.log(`${page.pageName}のクリアに失敗: ${pageError.message}`);
        failedPages.push(page.pageName);
      }
    });
    
    // 結果メッセージを作成
    let message = `クリア完了: ${successCount}/${pageInfo.length}ページ`;
    if (failedPages.length > 0) {
      message += `\n失敗したページ: ${failedPages.join(', ')}`;
    }
    
    UIHelper.showAlert(message);
    
  } catch (error) {
    Logger.log(`Error in clearAllPage: ${error.message}`);
    UIHelper.showAlert(`クリア中にエラーが発生しました: ${error.message}`);
  }
}