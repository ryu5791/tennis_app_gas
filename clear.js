/**
 * 1試合分のデータにシートの数式を挿入して初期状態に戻す関数
 * @param {string} topLeftCell - ゲームの左上セル (例: "B3")
 */
function clearOneGame(topLeftCell) {
  // "スコア入力"シート取得
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('スコア入力');
  
  // セル参照をA1形式からrow,columnに変換
  const cellA1Notation = topLeftCell.toUpperCase();
  const cell = sheet.getRange(cellA1Notation);
  const startRow = cell.getRow();
  const col = cell.getColumn();

  // 4名分の会員名、ID部分に数式を設定
  for (let i = 0; i < 4; i++) {
    const currentRow = startRow + i;
    
    // ID列のセル位置を取得 (ID列が先)
    const idCol = col + SheetInfo.OFFSET_COL_ID;
    const idCellRef = sheet.getRange(currentRow, idCol).getA1Notation();
    
    // 会員名列のセル (名前列は後)
    const nameCol = col + SheetInfo.OFFSET_COL_NAME;
    const nameCell = sheet.getRange(currentRow, nameCol);
    const nameCellRef = nameCell.getA1Notation();
    
    // 会員名のルックアップ式を設定（ID→会員名）
    const nameFormula = `=IFERROR(VLOOKUP(${idCellRef},'会員マスター'!B:C,2,FALSE),"")`;
    nameCell.setFormula(nameFormula);
    
    // ID列のセル
    const idCell = sheet.getRange(currentRow, idCol);
    
    // IDのルックアップ式を設定（会員名→ID）
    const idFormula = `=IFERROR(IFERROR(INDEX('会員マスター'!B:B, MATCH(${nameCellRef}, '会員マスター'!C:C, 0)), INDEX('会員マスター'!B:B, MATCH(${nameCellRef}, '会員マスター'!D:D, 0))),"")`;
    idCell.setFormula(idFormula);
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
 * ページ内のデータをクリアする関数
 * シートを削除して「スコア入力フォーマット」から再作成
 * @param {string} topLeftCell - ページの開始セル位置 (例: "B3")
 */
function clearOnePage(topLeftCell) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 「スコア入力」シートを削除
    const inputSheet = ss.getSheetByName('スコア入力');
    if (inputSheet) {
      ss.deleteSheet(inputSheet);
      Logger.log("「スコア入力」シートを削除しました");
    }
    
    // 「スコア入力フォーマット」シートをコピー
    const formatSheet = ss.getSheetByName('スコア入力フォーマット');
    if (!formatSheet) {
      throw new Error("「スコア入力フォーマット」シートが見つかりません");
    }
    
    // シートをコピーして名前を変更
    const newSheet = formatSheet.copyTo(ss);
    newSheet.setName('スコア入力');
    
    // シートを最初の位置に移動（オプション）
    ss.setActiveSheet(newSheet);
    ss.moveActiveSheet(1);
    
    Logger.log("「スコア入力」シートを再作成しました");
    
  } catch (error) {
    Logger.log(`clearOnePage実行エラー: ${error.message}`);
    throw error;
  }
}

/**
 * すべてのページのデータをクリアする関数（確認付き）
 * ※1ページ構成のため、実質clearOnePageと同じ
 */
function clearAllPage() {
  // 確認ダイアログを表示
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '確認',
    'すべてのデータをクリアしますか？\nこの操作は取り消せません。',
    ui.ButtonSet.YES_NO
  );
  
  // NOが選択された場合は中止
  if (response !== ui.Button.YES) {
    UIHelper.showAlert("クリアをキャンセルしました。");
    return;
  }
  
  try {
    Logger.log("clearAllPage開始");
    
    // clearOnePage（シート削除&再作成）を実行
    clearOnePage("B3");
    
    UIHelper.showAlert("クリアが完了しました。");
    
  } catch (error) {
    Logger.log(`Error in clearAllPage: ${error.message}`);
    UIHelper.showAlert(`クリア中にエラーが発生しました: ${error.message}`);
  }
}

/**
 * クリアボタンから呼び出される関数
 */
function clearData() {
  clearAllPage();
}

/**
 * 1試合分の罫線を設定する関数
 * @param {Sheet} sheet - 対象シート
 * @param {number} startRow - ゲームの開始行
 * @param {number} col - ゲームの開始列
 */
function setBordersForGame(sheet, startRow, col) {
  // 罫線のスタイルを設定
  const solidStyle = SpreadsheetApp.BorderStyle.SOLID;
  const mediumStyle = SpreadsheetApp.BorderStyle.SOLID_MEDIUM;
  
  // 4名分のセル範囲を取得
  const gameRange = sheet.getRange(startRow, col + SheetInfo.OFFSET_COL_ID, 4, 3);
  
  // 外枠を太線で設定
  gameRange.setBorder(true, true, true, true, false, false, null, mediumStyle);
  
  // 内部の横線を細線で設定
  for (let i = 0; i < 3; i++) {
    const rowRange = sheet.getRange(startRow + i, col + SheetInfo.OFFSET_COL_ID, 1, 3);
    rowRange.setBorder(false, false, true, false, false, false, null, solidStyle);
  }
  
  // チームA/Bの境界線を太線で設定
  const teamBorderRange = sheet.getRange(startRow + 1, col + SheetInfo.OFFSET_COL_ID, 1, 3);
  teamBorderRange.setBorder(false, false, true, false, false, false, null, mediumStyle);
}
