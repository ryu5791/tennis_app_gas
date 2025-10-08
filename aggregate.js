/**
 * スコア集計機能
 * 期間を指定して会員とゲストのスコアを集計する
 */

/**
 * 集計処理のメイン関数（メニューから呼び出される）
 */
function aggregateScores() {
  try {
    // 日付入力ダイアログを表示
    const dateRange = showDateInputDialog();
    
    if (!dateRange) {
      UIHelper.showAlert("集計をキャンセルしました。");
      return;
    }
    
    const { startDate, endDate } = dateRange;
    
    // 集計処理を実行
    const result = executeAggregation(startDate, endDate);
    
    if (result.success) {
      UIHelper.showAlert(
        `集計が完了しました。\n` +
        `対象期間: ${formatDate(startDate)} ～ ${formatDate(endDate)}\n` +
        `会員: ${result.memberCount}名\n` +
        `ゲスト: ${result.guestCount}名`
      );
    } else {
      UIHelper.showAlert(`集計中にエラーが発生しました: ${result.error}`);
    }
    
  } catch (error) {
    Logger.log(`Error in aggregateScores: ${error.message}`);
    UIHelper.showAlert(`エラーが発生しました: ${error.message}`);
  }
}

/**
 * 日付入力ダイアログを表示
 * @return {Object|null} {startDate: Date, endDate: Date} または null
 */
function showDateInputDialog() {
  const ui = SpreadsheetApp.getUi();
  
  // 開始日の入力
  const startDateResponse = ui.prompt(
    '集計期間の設定',
    '開始日を入力してください（例: 2025/01/01）:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (startDateResponse.getSelectedButton() !== ui.Button.OK) {
    return null;
  }
  
  const startDateStr = startDateResponse.getResponseText().trim();
  if (!startDateStr) {
    UIHelper.showAlert("開始日が入力されていません。");
    return null;
  }
  
  // 終了日の入力
  const endDateResponse = ui.prompt(
    '集計期間の設定',
    '終了日を入力してください（例: 2025/12/31）:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (endDateResponse.getSelectedButton() !== ui.Button.OK) {
    return null;
  }
  
  const endDateStr = endDateResponse.getResponseText().trim();
  if (!endDateStr) {
    UIHelper.showAlert("終了日が入力されていません。");
    return null;
  }
  
  // 日付の妥当性チェック
  try {
    const startDate = new Date(startDateStr);
    const endDate = new Date(endDateStr);
    
    if (isNaN(startDate.getTime()) || isNaN(endDate.getTime())) {
      UIHelper.showAlert("日付の形式が正しくありません。");
      return null;
    }
    
    if (startDate > endDate) {
      UIHelper.showAlert("開始日は終了日より前の日付を指定してください。");
      return null;
    }
    
    return { startDate, endDate };
    
  } catch (error) {
    UIHelper.showAlert("日付の変換に失敗しました: " + error.message);
    return null;
  }
}

/**
 * 集計処理を実行
 * @param {Date} startDate - 開始日
 * @param {Date} endDate - 終了日
 * @return {Object} 処理結果
 */
function executeAggregation(startDate, endDate) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. スコア出力シートから期間内のデータを取得
    const scoreData = getScoreDataInRange(startDate, endDate);
    Logger.log(`対象データ数: ${scoreData.length}件`);
    
    // 2. IDごとに集計
    const aggregatedData = aggregateByPlayer(scoreData);
    Logger.log(`集計対象プレイヤー数: ${Object.keys(aggregatedData).length}名`);
    
    // 3. 参加日数を計算または取得
    const participationDays = getParticipationDays(scoreData, startDate, endDate);
    
    // 4. HDCP情報を取得
    const hdcpData = getHDCPData();
    
    // 5. 会員とゲストに分類
    const { members, guests } = classifyPlayers(aggregatedData, hdcpData, participationDays);
    Logger.log(`会員: ${members.length}名, ゲスト: ${guests.length}名`);
    
    // 6. 会員の参加日数閾値を計算（20番目）
    const threshold = calculateThreshold(members);
    Logger.log(`参加日数閾値: ${threshold}日`);
    
    // 7. 閾値以上の会員を抽出してNet順にソート
    const qualifiedMembers = members
      .filter(m => m.participationDays >= threshold)
      .sort((a, b) => b.net - a.net);
    Logger.log(`閾値以上の会員: ${qualifiedMembers.length}名`);
    
    // 8. スコア集計シートに出力
    outputToSheet(qualifiedMembers, guests, threshold, startDate, endDate);
    
    return {
      success: true,
      memberCount: qualifiedMembers.length,
      guestCount: guests.length
    };
    
  } catch (error) {
    Logger.log(`Error in executeAggregation: ${error.message}`);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * スコア出力シートから期間内のデータを取得
 * @param {Date} startDate - 開始日
 * @param {Date} endDate - 終了日
 * @return {Array} スコアデータの配列
 */
function getScoreDataInRange(startDate, endDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const scoreSheet = ss.getSheetByName('スコア出力');
  
  if (!scoreSheet) {
    throw new Error("スコア出力シートが見つかりません");
  }
  
  const lastRow = scoreSheet.getLastRow();
  if (lastRow <= 1) {
    return [];
  }
  
  // データを取得（ヘッダー行を除く）
  const data = scoreSheet.getRange(2, 1, lastRow - 1, 10).getValues();
  
  // 期間内のデータをフィルタリング
  const filteredData = data.filter(row => {
    const date = new Date(row[0]);
    return date >= startDate && date <= endDate;
  });
  
  // オブジェクト配列に変換
  return filteredData.map(row => ({
    date: new Date(row[0]),
    gameNo: row[1],
    id: row[2],
    pairId: row[3],
    gamePt: row[6],
    row: row[8]
  }));
}

/**
 * プレイヤーごとにデータを集計
 * @param {Array} scoreData - スコアデータ
 * @return {Object} ID別の集計データ
 */
function aggregateByPlayer(scoreData) {
  const aggregated = {};
  
  scoreData.forEach(record => {
    const id = record.id;
    
    if (!aggregated[id]) {
      aggregated[id] = {
        id: id,
        gameCount: 0,
        totalPoints: 0,
        dates: new Set()  // 参加日の重複を排除するためのSet
      };
    }
    
    aggregated[id].gameCount++;
    aggregated[id].totalPoints += Number(record.gamePt);
    aggregated[id].dates.add(formatDate(record.date));
  });
  
  // Gross（平均）を計算
  Object.values(aggregated).forEach(player => {
    player.gross = player.totalPoints / player.gameCount;
    player.participationDaysDB = player.dates.size;  // DB計算の参加日数
    delete player.dates;  // Setは不要になったので削除
  });
  
  return aggregated;
}

/**
 * 参加日数を取得（シートがあればそれを使用、なければDB計算値）
 * @param {Array} scoreData - スコアデータ
 * @param {Date} startDate - 開始日
 * @param {Date} endDate - 終了日
 * @return {Object} ID別の参加日数
 */
function getParticipationDays(scoreData, startDate, endDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const participationSheet = ss.getSheetByName('参加日数');
  
  // 参加日数シートがある場合
  if (participationSheet) {
    Logger.log("参加日数シートから参加日数を取得します");
    return getParticipationDaysFromSheet(participationSheet, startDate, endDate);
  }
  
  // ない場合はDB計算値を使用
  Logger.log("DB計算の参加日数を使用します");
  const participation = {};
  
  scoreData.forEach(record => {
    const id = record.id;
    if (!participation[id]) {
      participation[id] = new Set();
    }
    participation[id].add(formatDate(record.date));
  });
  
  const result = {};
  Object.keys(participation).forEach(id => {
    result[id] = participation[id].size;
  });
  
  return result;
}

/**
 * 参加日数シートから参加日数を取得
 * @param {Sheet} sheet - 参加日数シート
 * @param {Date} startDate - 開始日
 * @param {Date} endDate - 終了日
 * @return {Object} ID別の参加日数
 */
function getParticipationDaysFromSheet(sheet, startDate, endDate) {
  // TODO: 参加日数シートの構造に応じて実装を調整
  // 仮実装：ID列(A列)と参加日数列(B列)があると仮定
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return {};
  }
  
  const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  const result = {};
  
  data.forEach(row => {
    const id = row[0];
    const days = row[1];
    if (id && days) {
      result[id] = days;
    }
  });
  
  return result;
}

/**
 * HDCP情報を取得
 * @return {Object} ID別のHDCP情報
 */
function getHDCPData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hdcpSheet = ss.getSheetByName('HDCP');
  
  if (!hdcpSheet) {
    Logger.log("HDCPシートが見つかりません。全員をゲストとして扱います。");
    return {};
  }
  
  const lastRow = hdcpSheet.getLastRow();
  if (lastRow <= 1) {
    return {};
  }
  
  // TODO: HDCPシートの構造に応じて調整
  // 仮実装：B列=ID, T列(20列目)=会員フラグ, 適切な列にHDCP値があると仮定
  const data = hdcpSheet.getRange(2, 1, lastRow - 1, 20).getValues();
  const result = {};
  
  data.forEach(row => {
    const id = row[1];  // B列（インデックス1）
    const isMember = row[19] === 1;  // T列（インデックス19）
    const hdcp = row[5] || 0;  // F列を仮のHDCP列とする（要調整）
    
    if (id) {
      result[id] = {
        isMember: isMember,
        hdcp: hdcp
      };
    }
  });
  
  return result;
}

/**
 * プレイヤーを会員とゲストに分類
 * @param {Object} aggregatedData - 集計データ
 * @param {Object} hdcpData - HDCP情報
 * @param {Object} participationDays - 参加日数
 * @return {Object} {members: Array, guests: Array}
 */
function classifyPlayers(aggregatedData, hdcpData, participationDays) {
  const members = [];
  const guests = [];
  
  Object.values(aggregatedData).forEach(player => {
    const id = player.id;
    const hdcpInfo = hdcpData[id] || { isMember: false, hdcp: 0 };
    const days = participationDays[id] || player.participationDaysDB;
    
    // Net = Gross - HDCP
    const net = player.gross - hdcpInfo.hdcp;
    
    const playerData = {
      id: id,
      gameCount: player.gameCount,
      totalPoints: player.totalPoints,
      gross: player.gross,
      hdcp: hdcpInfo.hdcp,
      net: net,
      participationDays: days
    };
    
    if (hdcpInfo.isMember) {
      members.push(playerData);
    } else {
      guests.push(playerData);
    }
  });
  
  return { members, guests };
}

/**
 * 参加日数の閾値を計算（20番目の値）
 * @param {Array} members - 会員データ
 * @return {number} 閾値
 */
function calculateThreshold(members) {
  if (members.length === 0) {
    return 0;
  }
  
  // 参加日数の降順でソート
  const sorted = members
    .map(m => m.participationDays)
    .sort((a, b) => b - a);
  
  // 20番目の値を取得（配列は0始まりなので19番目のインデックス）
  const index = Math.min(19, sorted.length - 1);
  return sorted[index] || 0;
}

/**
 * スコア集計シートに結果を出力
 * @param {Array} qualifiedMembers - 閾値以上の会員
 * @param {Array} guests - ゲスト
 * @param {number} threshold - 参加日数閾値
 * @param {Date} startDate - 開始日
 * @param {Date} endDate - 終了日
 */
function outputToSheet(qualifiedMembers, guests, threshold, startDate, endDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let outputSheet = ss.getSheetByName('スコア集計');
  
  // シートがなければ作成
  if (!outputSheet) {
    outputSheet = ss.insertSheet('スコア集計');
  } else {
    // 既存のシートをクリア
    outputSheet.clear();
  }
  
  // ヘッダー情報を出力
  outputSheet.getRange(1, 1).setValue('集計期間:');
  outputSheet.getRange(1, 2).setValue(`${formatDate(startDate)} ～ ${formatDate(endDate)}`);
  outputSheet.getRange(2, 1).setValue('参加日数閾値:');
  outputSheet.getRange(2, 2).setValue(`${threshold}日以上`);
  
  let currentRow = 4;
  
  // 会員セクション
  outputSheet.getRange(currentRow, 1).setValue('【会員】');
  currentRow++;
  
  // 会員のヘッダー
  const memberHeaders = ['順位', '会員ID', '会員名', '合計', '試合数', 'Gross', 'HDCP', 'Net', 'クロス順位', '参加日数', '備考'];
  outputSheet.getRange(currentRow, 1, 1, memberHeaders.length).setValues([memberHeaders]);
  currentRow++;
  
  // 会員データを出力
  qualifiedMembers.forEach((member, index) => {
    const name = getPlayerName(member.id);
    const row = [
      index + 1,  // 順位
      member.id,
      name,
      member.totalPoints,
      member.gameCount,
      member.gross.toFixed(3),
      member.hdcp.toFixed(3),
      member.net.toFixed(3),
      '',  // クロス順位（要確認）
      member.participationDays,
      ''   // 備考
    ];
    outputSheet.getRange(currentRow, 1, 1, row.length).setValues([row]);
    currentRow++;
  });
  
  currentRow += 2;
  
  // ゲストセクション
  outputSheet.getRange(currentRow, 1).setValue('【ゲスト】');
  currentRow++;
  
  // ゲストのヘッダー
  outputSheet.getRange(currentRow, 1, 1, memberHeaders.length).setValues([memberHeaders]);
  currentRow++;
  
  // ゲストデータを出力（Net順）
  guests.sort((a, b) => b.net - a.net);
  guests.forEach((guest, index) => {
    const name = getPlayerName(guest.id);
    const row = [
      index + 1,
      guest.id,
      name,
      guest.totalPoints,
      guest.gameCount,
      guest.gross.toFixed(3),
      guest.hdcp.toFixed(3),
      guest.net.toFixed(3),
      '',
      guest.participationDays,
      ''
    ];
    outputSheet.getRange(currentRow, 1, 1, row.length).setValues([row]);
    currentRow++;
  });
  
  // 書式設定
  formatAggregationSheet(outputSheet, qualifiedMembers.length, guests.length);
}

/**
 * プレイヤー名を取得
 * @param {string|number} id - プレイヤーID
 * @return {string} プレイヤー名
 */
function getPlayerName(id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName('会員マスター');
  
  if (!masterSheet) {
    return String(id);
  }
  
  const data = masterSheet.getDataRange().getValues();
  
  // ID列（B列=インデックス1）と名前列（A列=インデックス0）を検索
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] == id) {
      return data[i][0] || String(id);
    }
  }
  
  return String(id);
}

/**
 * スコア集計シートの書式を設定
 * @param {Sheet} sheet - 対象シート
 * @param {number} memberCount - 会員数
 * @param {number} guestCount - ゲスト数
 */
function formatAggregationSheet(sheet, memberCount, guestCount) {
  // ヘッダー行を太字に
  sheet.getRange(5, 1, 1, 11).setFontWeight('bold');
  sheet.getRange(5 + memberCount + 3, 1, 1, 11).setFontWeight('bold');
  
  // 列幅を調整
  sheet.setColumnWidth(1, 50);   // 順位
  sheet.setColumnWidth(2, 80);   // 会員ID
  sheet.setColumnWidth(3, 120);  // 会員名
  sheet.setColumnWidth(4, 60);   // 合計
  sheet.setColumnWidth(5, 60);   // 試合数
  sheet.setColumnWidth(6, 70);   // Gross
  sheet.setColumnWidth(7, 70);   // HDCP
  sheet.setColumnWidth(8, 70);   // Net
  sheet.setColumnWidth(9, 80);   // クロス順位
  sheet.setColumnWidth(10, 80);  // 参加日数
  sheet.setColumnWidth(11, 100); // 備考
  
  // 数値列を右揃え
  const numericColumns = [1, 2, 4, 5, 6, 7, 8, 10];
  numericColumns.forEach(col => {
    sheet.getRange(6, col, memberCount + guestCount + 10).setHorizontalAlignment('right');
  });
}

/**
 * 日付を文字列にフォーマット
 * @param {Date} date - 日付オブジェクト
 * @return {string} フォーマットされた日付文字列
 */
function formatDate(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy/MM/dd");
}

/**
 * メニューに集計機能を追加
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('テニススコア')
    .addItem('クリア', 'clearData')
    .addItem('全ページクリア', 'clearAllPage')
    .addItem('登録(1ゲーム)', 'registerData')
    .addItem('登録(全ゲーム)', 'getAllGame')
    .addSeparator()
    .addItem('スコア集計', 'aggregateScores')  // 新規追加
    .addToUi();
}