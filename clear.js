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
  scoreLowerCell.setFormula(`=IF(OR(${upperCellRef}=0, ${lowerCellRef}=1, ${upperCellRef}=2, ${upperCellRef}=3), 5, "")`);
}