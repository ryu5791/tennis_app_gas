/**
 * スコア集計機能
 * 期間を指定して会員とゲストのスコアを集計する
 */

// マスターデータのキャッシュ（関数実行ごとにクリア）
let masterDataCache = null;

/**
 * 会員マスターデータを辞書型で取得（キャッシュ利用）
 * @return {Object} {idToName: {}, nameToId: {}} の辞書
 */
function getMasterDataDict() {
  if (masterDataCache !== null) {
    return masterDataCache;
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dict = {
    idToName: {},
    nameToId: {}
  };
  
  // 会員マスターから取得
  const masterSheet = ss.getSheetByName('会員マスター');
  if (masterSheet) {
    const masterData = masterSheet.getDataRange().getValues();
    // ヘッダー行をスキップ（1行目から開始）
    for (let i = 1; i < masterData.length; i++) {
      const name = masterData[i][0];  // A列: 名前
      const id = masterData[i][1];    // B列: ID
      if (id !== '' && id !== null) {
        dict.idToName[id] = name || String(id);
        if (name) {
          dict.nameToId[name] = id;
        }
      }
    }
  }
  
  masterDataCache = dict;
  return dict;
}

/**
 * キャッシュをクリアする
 */
function clearMasterDataCache() {
  masterDataCache = null;
}

/**
 * 集計処理のメイン関数（メニューから呼び出される）
 */
function aggregateScores() {
  try {
    // キャッシュをクリア
    clearMasterDataCache();
    
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
    
    // 5.5. グロス順位を計算（上位10名）
    calculateGrossRank(members);
    
    // 6. 会員の参加日数閾値を計算（20番目）
    const threshold = calculateThreshold(members);
    Logger.log(`参加日数閾値: ${threshold}日`);
    
    // 7. 閾値以上の会員、または備考（前期/前々期の順位）がある会員を抽出
    const qualifiedMembers = members
      .filter(m => {
        // 参加日数が閾値以上
        if (m.participationDays >= threshold) return true;
        // 備考に「前期」または「前々期」が含まれている場合も対象
        if (m.remarks && (m.remarks.includes('前期') || m.remarks.includes('前々期'))) return true;
        return false;
      })
      .sort((a, b) => b.net - a.net);
    Logger.log(`閾値以上または備考ありの会員: ${qualifiedMembers.length}名`);
    
    // 8. スコア集計シートに出力（全会員を出力）
    outputToSheet(members, guests, threshold, startDate, endDate);
    
    // 9. スコア集計（試合数加味）シートを作成
    outputWeightedSheet(qualifiedMembers, members, guests, threshold, startDate, endDate);
    
    // 10. 月別参加日数シートを出力
    const monthlyData = calculateMonthlyParticipation(scoreData, startDate, endDate);
    outputMonthlyParticipationSheet(monthlyData, hdcpData, startDate, endDate);
    
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
 * 月別参加日数を計算
 * @param {Array} scoreData - スコアデータ
 * @param {Date} startDate - 開始日
 * @param {Date} endDate - 終了日
 * @return {Object} ID別の月別参加日数 {id: {1: 3, 2: 5, ...}}
 */
function calculateMonthlyParticipation(scoreData, startDate, endDate) {
  const monthlyData = {};
  
  scoreData.forEach(record => {
    const id = record.id;
    const date = new Date(record.date);
    const month = date.getMonth() + 1; // 1-12
    
    if (!monthlyData[id]) {
      monthlyData[id] = {};
      for (let m = 1; m <= 12; m++) {
        monthlyData[id][m] = new Set(); // 重複を排除するためSet使用
      }
    }
    
    // その月の日付を記録
    const dateStr = formatDate(date);
    monthlyData[id][month].add(dateStr);
  });
  
  // Setをカウント数に変換
  const result = {};
  Object.keys(monthlyData).forEach(id => {
    result[id] = {};
    for (let m = 1; m <= 12; m++) {
      result[id][m] = monthlyData[id][m].size;
    }
    // 年間合計を計算
    result[id].total = Object.values(result[id]).reduce((sum, count) => sum + count, 0);
  });
  
  return result;
}

/**
 * 月別参加日数シートを出力
 * @param {Object} monthlyData - 月別参加日数データ
 * @param {Object} hdcpData - HDCP情報
 * @param {Date} startDate - 開始日
 * @param {Date} endDate - 終了日
 */
function outputMonthlyParticipationSheet(monthlyData, hdcpData, startDate, endDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const year = startDate.getFullYear();
  
  // シートを作成または取得
  let sheet = ss.getSheetByName('参加日数');
  if (!sheet) {
    sheet = ss.insertSheet('参加日数');
  } else {
    sheet.clear();
  }
  
  // タイトル行
  sheet.getRange(1, 2).setValue(`${year}年`);
  sheet.getRange(1, 7).setValue('年間レッツ会計報告 及び コート使用状況');
  
  // ヘッダー行
  const headers = ['ID', '会員名', '備考欄', '年会費'];
  const months = ['1月', '2月', '3月', '4月', '5月', '6月', '7月', '8月', '9月', '10月', '11月', '12月'];
  sheet.getRange(4, 1, 1, 4).setValues([headers]);
  sheet.getRange(5, 4).setValue('年会員');
  sheet.getRange(5, 5, 1, months.length).setValues([months]);  // E列(5列目)から月名を配置
  
  let currentRow = 6;
  
  // 会員のみを抽出してソート（hdcpDataは既に辞書型なのでO(1)で検索可能）
  const members = Object.keys(monthlyData)
    .filter(id => hdcpData[id] && hdcpData[id].isMember)
    .sort((a, b) => Number(a) - Number(b));
  
  // データ行を出力
  members.forEach(id => {
    const name = getPlayerName(id);
    const row = [id, name, '', 3]; // 年会費は3（固定値、必要に応じて変更）
    
    // 各月の参加日数を追加
    for (let m = 1; m <= 12; m++) {
      const count = monthlyData[id][m] || 0;
      row.push(count > 0 ? count : '－');
    }
    
    sheet.getRange(currentRow, 1, 1, row.length).setValues([row]);
    currentRow++;
  });
  
  // 書式設定
  formatMonthlyParticipationSheet(sheet, members.length);
}

/**
 * 月別参加日数シートの書式を設定
 * @param {Sheet} sheet - 対象シート
 * @param {number} memberCount - 会員数
 */
function formatMonthlyParticipationSheet(sheet, memberCount) {
  // 列幅調整
  sheet.setColumnWidth(1, 60);   // ID
  sheet.setColumnWidth(2, 120);  // 会員名
  sheet.setColumnWidth(3, 100);  // 備考欄
  sheet.setColumnWidth(4, 70);   // 年会費
  
  // 月の列幅
  for (let col = 5; col <= 16; col++) {
    sheet.setColumnWidth(col, 50);
  }
  
  // タイトル行の書式
  sheet.getRange(1, 2).setFontSize(14).setFontWeight('bold');
  sheet.getRange(1, 7).setFontSize(12).setFontWeight('bold');
  
  // ヘッダー行の書式（4-5行目）
  sheet.getRange(4, 1, 2, 16)
    .setBackground('#D9EAD3')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  // データ行の書式
  if (memberCount > 0) {
    // ID列を中央揃え
    sheet.getRange(6, 1, memberCount, 1).setHorizontalAlignment('center');
    
    // 年会費と月の列を中央揃え
    sheet.getRange(6, 4, memberCount, 13).setHorizontalAlignment('center');
    
    // 月の列（参加日数）に色付け（交互に）
    for (let col = 5; col <= 16; col += 2) {
      sheet.getRange(6, col, memberCount, 1).setBackground('#E8F4E8');
    }
    
    // 罫線
    sheet.getRange(4, 1, memberCount + 2, 16).setBorder(
      true, true, true, true, true, true,
      'black', SpreadsheetApp.BorderStyle.SOLID
    );
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
  
  // ない場合はDB計算値を使用（辞書型で構築）
  Logger.log("DB計算の参加日数を使用します");
  const participation = {};
  
  // Setを使って日付の重複を排除
  scoreData.forEach(record => {
    const id = record.id;
    if (!participation[id]) {
      participation[id] = new Set();
    }
    participation[id].add(formatDate(record.date));
  });
  
  // Setのサイズ（ユニークな日数）を数値化
  const result = {};
  Object.keys(participation).forEach(id => {
    result[id] = participation[id].size;
  });
  
  return result;
}

/**
 * 参加日数シートから参加日数を取得（辞書型で返す）
 * @param {Sheet} sheet - 参加日数シート
 * @param {Date} startDate - 開始日
 * @param {Date} endDate - 終了日
 * @return {Object} ID別の参加日数
 */
function getParticipationDaysFromSheet(sheet, startDate, endDate) {
  // A列: 会員ID, B列: 参加日数
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return {};
  }
  
  const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  const result = {};
  
  // 配列を一度だけループして辞書を構築
  data.forEach(row => {
    const id = row[0];
    const days = row[1];
    if (id !== '' && id !== null && days !== '' && days !== null) {
      result[id] = Number(days);
    }
  });
  
  return result;
}

/**
 * HDCP情報を取得（辞書型で返す）
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
  
  // A列=ID, N列(14列目)=新ハンディ, O列(15列目)=備考, T列(20列目)=会員フラグ
  const data = hdcpSheet.getRange(2, 1, lastRow - 1, 20).getValues();
  const result = {};
  
  // 配列を一度だけループして辞書を構築
  data.forEach(row => {
    const id = row[0];  // A列（インデックス0）
    if (id === '' || id === null) return;
    
    const memberFlag = row[19];  // T列（インデックス19）
    // T列が厳密に1の場合のみ会員とする
    const isMember = (memberFlag === 1 || memberFlag === "1" || Number(memberFlag) === 1);
    const hdcp = row[13] || 0;  // N列（インデックス13）新ハンディ
    const remarks = row[14] || '';  // O列（インデックス14）備考欄
    
    result[id] = {
      isMember: isMember,
      hdcp: Number(hdcp),
      remarks: String(remarks).trim()
    };
    
    Logger.log(`ID: ${id}, T列の値: ${memberFlag}, 会員判定: ${isMember}`);
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
    const hdcpInfo = hdcpData[id] || { isMember: false, hdcp: 0, remarks: '' };
    const days = participationDays[id] || player.participationDaysDB;
    
    // Net = Gross + HDCP（ハンディキャップを加算）
    const net = player.gross + hdcpInfo.hdcp;
    
    const playerData = {
      id: id,
      gameCount: player.gameCount,
      totalPoints: player.totalPoints,
      gross: player.gross,
      hdcp: hdcpInfo.hdcp,
      net: net,
      participationDays: days,
      remarks: hdcpInfo.remarks || ''
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
 * グロス順位を計算（上位10名）
 * @param {Array} members - 会員データ
 */
function calculateGrossRank(members) {
  // グロス降順でソート（コピーを作成）
  const sortedByGross = [...members].sort((a, b) => b.gross - a.gross);
  
  // 上位10名にグロス順位を設定
  sortedByGross.forEach((member, index) => {
    if (index < 10) {
      member.grossRank = index + 1;
    } else {
      member.grossRank = '';
    }
  });
}

/**
 * スコア集計シートに結果を出力
 * @param {Array} members - 全会員
 * @param {Array} guests - ゲスト
 * @param {number} threshold - 参加日数閾値
 * @param {Date} startDate - 開始日
 * @param {Date} endDate - 終了日
 */
function outputToSheet(members, guests, threshold, startDate, endDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // スコア集計シートを作成または取得
  let outputSheet = ss.getSheetByName('スコア集計');
  if (!outputSheet) {
    outputSheet = ss.insertSheet('スコア集計');
  } else {
    outputSheet.clear();
  }
  
  let currentRow = 1;
  
  // ヘッダー行
  const memberHeaders = ['順位', '会員ＩＤ', '会員名', '合計', '試合数', 'Gross', 'HDCP', 'Net', 'ｸﾞﾛｽ順位', '参加日数', '備考'];
  outputSheet.getRange(currentRow, 1, 1, memberHeaders.length).setValues([memberHeaders]);
  outputSheet.getRange(currentRow, 1, 1, memberHeaders.length)
    .setBackground('#4A86E8')
    .setFontColor('white')
    .setFontWeight('bold');
  currentRow++;
  
  // 会員データを会員ID昇順でソート（非破壊）
  const sortedMembers = [...members].sort((a, b) => Number(a.id) - Number(b.id));
  
  // 会員データを出力
  sortedMembers.forEach((member, index) => {
    const name = getPlayerName(member.id);
    const row = [
      index + 1,
      member.id,
      name,
      member.totalPoints,
      member.gameCount,
      member.gross.toFixed(3),
      member.hdcp.toFixed(3),
      member.net.toFixed(3),
      member.grossRank || '',  // グロス順位
      member.participationDays,
      member.remarks || ''   // 備考
    ];
    outputSheet.getRange(currentRow, 1, 1, row.length).setValues([row]);
    currentRow++;
  });
  
  // ゲストデータを会員ID昇順でソート（非破壊）
  const sortedGuests = [...guests].sort((a, b) => Number(a.id) - Number(b.id));
  
  // ゲストデータを出力（順位欄に「ゲスト」を表示）
  sortedGuests.forEach((guest) => {
    const name = getPlayerName(guest.id);
    const row = [
      'ゲスト',  // 順位欄に「ゲスト」
      guest.id,
      name,
      guest.totalPoints,
      guest.gameCount,
      guest.gross.toFixed(3),
      guest.hdcp.toFixed(3),
      guest.net.toFixed(3),
      '',
      guest.participationDays,
      guest.remarks || ''
    ];
    outputSheet.getRange(currentRow, 1, 1, row.length).setValues([row]);
    currentRow++;
  });
  
  // 書式設定
  formatAggregationSheet(outputSheet, members.length, guests.length);
}

/**
 * スコア集計（試合数加味）シートに結果を出力
 * @param {Array} qualifiedMembers - 閾値以上の会員
 * @param {Array} allMembers - 全会員（閾値未満含む）
 * @param {Array} guests - ゲスト
 * @param {number} threshold - 参加日数閾値
 * @param {Date} startDate - 開始日
 * @param {Date} endDate - 終了日
 */
function outputWeightedSheet(qualifiedMembers, allMembers, guests, threshold, startDate, endDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // スコア集計（試合数加味）シートを作成または取得
  let weightedSheet = ss.getSheetByName('スコア集計（試合数加味）');
  if (!weightedSheet) {
    weightedSheet = ss.insertSheet('スコア集計（試合数加味）');
  } else {
    weightedSheet.clear();
  }
  
  let currentRow = 1;
  
  // ヘッダー行
  const headers = ['順位', '会員ＩＤ', '会員名', '合計', '試合数', 'Gross', 'HDCP', 'Net', 'ｸﾞﾛｽ順位', '参加日数', '備考'];
  weightedSheet.getRange(currentRow, 1, 1, headers.length).setValues([headers]);
  weightedSheet.getRange(currentRow, 1, 1, headers.length)
    .setBackground('#6AA84F')
    .setFontColor('white')
    .setFontWeight('bold');
  currentRow++;
  
  // 閾値以上の会員をNet順降順でソート（非破壊）
  const sortedQualified = [...qualifiedMembers].sort((a, b) => b.net - a.net);
  
  // 閾値以上の会員データを出力（順位付き）
  sortedQualified.forEach((member, index) => {
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
      member.grossRank || '',  // グロス順位
      member.participationDays,
      member.remarks || ''
    ];
    weightedSheet.getRange(currentRow, 1, 1, row.length).setValues([row]);
    currentRow++;
  });
  
  // 閾値未満の会員を抽出してNet順降順でソート
  const qualifiedIds = new Set(qualifiedMembers.map(m => m.id));
  const unqualifiedMembers = allMembers.filter(m => !qualifiedIds.has(m.id));
  const sortedUnqualified = [...unqualifiedMembers].sort((a, b) => b.net - a.net);
  
  // 閾値未満の会員データを出力（順位なし）
  sortedUnqualified.forEach((member) => {
    const name = getPlayerName(member.id);
    const row = [
      '',  // 順位欄は空欄
      member.id,
      name,
      member.totalPoints,
      member.gameCount,
      member.gross.toFixed(3),
      member.hdcp.toFixed(3),
      member.net.toFixed(3),
      member.grossRank || '',
      member.participationDays,
      member.remarks || ''
    ];
    weightedSheet.getRange(currentRow, 1, 1, row.length).setValues([row]);
    currentRow++;
  });
  
  // 規定日数の注釈行
  const noteRow = ['', '', '規定日数', '＝', '日数上位 20名', '', `日数  ${threshold}日以上をランキング対象とする`, '', '', '', ''];
  weightedSheet.getRange(currentRow, 1, 1, noteRow.length).setValues([noteRow]);
  currentRow++;
  
  // ゲストデータをNet順降順でソート（非破壊）
  const sortedGuests = [...guests].sort((a, b) => b.net - a.net);
  
  // ゲストデータを出力（順位欄に「ゲスト」）
  sortedGuests.forEach((guest) => {
    const name = getPlayerName(guest.id);
    const row = [
      'ゲスト',  // 順位欄に「ゲスト」
      guest.id,
      name,
      guest.totalPoints,
      guest.gameCount,
      guest.gross.toFixed(3),
      guest.hdcp.toFixed(3),
      guest.net.toFixed(3),
      '',
      guest.participationDays,
      guest.remarks || ''
    ];
    weightedSheet.getRange(currentRow, 1, 1, row.length).setValues([row]);
    currentRow++;
  });
  
  // 書式設定
  formatWeightedSheet(weightedSheet, qualifiedMembers.length + unqualifiedMembers.length, guests.length);
}

/**
 * スコア集計（試合数加味）シートの書式を設定
 * @param {Sheet} sheet - 対象シート
 * @param {number} memberCount - 会員数
 * @param {number} guestCount - ゲスト数
 */
function formatWeightedSheet(sheet, memberCount, guestCount) {
  // 列幅を調整
  sheet.setColumnWidth(1, 50);   // 順位
  sheet.setColumnWidth(2, 80);   // 会員ID
  sheet.setColumnWidth(3, 120);  // 会員名
  sheet.setColumnWidth(4, 70);   // 合計
  sheet.setColumnWidth(5, 70);   // 試合数
  sheet.setColumnWidth(6, 80);   // Gross
  sheet.setColumnWidth(7, 80);   // HDCP
  sheet.setColumnWidth(8, 80);   // Net
  sheet.setColumnWidth(9, 80);   // クロス順位
  sheet.setColumnWidth(10, 80);  // 参加日数
  sheet.setColumnWidth(11, 150); // 備考
  
  // 会員データの書式
  if (memberCount > 0) {
    sheet.getRange(2, 1, memberCount, 1).setHorizontalAlignment('center');  // 順位
    sheet.getRange(2, 2, memberCount, 1).setHorizontalAlignment('center');  // ID
    sheet.getRange(2, 4, memberCount, 7).setHorizontalAlignment('right');   // 数値列
  }
  
  // ゲストデータの書式（開始行を計算）
  // レイアウト: ヘッダー(1) + 会員データ + 空行(1) + 区切り(1) + 注釈(1) + 空行(1) + ゲストタイトル(1) + ゲストヘッダー(1)
  const guestStartRow = 7 + memberCount;
  if (guestCount > 0) {
    sheet.getRange(guestStartRow + 1, 1, guestCount, 1).setHorizontalAlignment('center');
    sheet.getRange(guestStartRow + 1, 2, guestCount, 1).setHorizontalAlignment('center');
    sheet.getRange(guestStartRow + 1, 4, guestCount, 7).setHorizontalAlignment('right');
  }
  
  // 罫線を設定
  if (memberCount > 0) {
    sheet.getRange(1, 1, memberCount + 1, 11).setBorder(
      true, true, true, true, true, true,
      'black', SpreadsheetApp.BorderStyle.SOLID
    );
  }
  
  if (guestCount > 0) {
    sheet.getRange(guestStartRow, 1, guestCount + 1, 11).setBorder(
      true, true, true, true, true, true,
      'black', SpreadsheetApp.BorderStyle.SOLID
    );
  }
}

/**
 * プレイヤー名を取得（辞書型検索）
 * @param {string|number} id - プレイヤーID
 * @return {string} プレイヤー名
 */
function getPlayerName(id) {
  const masterDict = getMasterDataDict();
  return masterDict.idToName[id] || String(id);
}

/**
 * スコア集計シートの書式を設定
 * @param {Sheet} sheet - 対象シート
 * @param {number} memberCount - 会員数
 * @param {number} guestCount - ゲスト数
 */
function formatAggregationSheet(sheet, memberCount, guestCount) {
  // 列幅を調整
  sheet.setColumnWidth(1, 50);   // 順位
  sheet.setColumnWidth(2, 80);   // 会員ID
  sheet.setColumnWidth(3, 120);  // 会員名
  sheet.setColumnWidth(4, 70);   // 合計
  sheet.setColumnWidth(5, 70);   // 試合数
  sheet.setColumnWidth(6, 80);   // Gross
  sheet.setColumnWidth(7, 80);   // HDCP
  sheet.setColumnWidth(8, 80);   // Net
  sheet.setColumnWidth(9, 80);   // クロス順位
  sheet.setColumnWidth(10, 80);  // 参加日数
  sheet.setColumnWidth(11, 150); // 備考
  
  // 数値列を右揃えまたは中央揃え
  const memberStartRow = 6;
  const guestStartRow = memberStartRow + memberCount + 3;
  
  if (memberCount > 0) {
    sheet.getRange(memberStartRow, 1, memberCount, 1).setHorizontalAlignment('center');  // 順位
    sheet.getRange(memberStartRow, 2, memberCount, 1).setHorizontalAlignment('center');  // ID
    sheet.getRange(memberStartRow, 4, memberCount, 7).setHorizontalAlignment('right');   // 数値列
  }
  
  if (guestCount > 0) {
    sheet.getRange(guestStartRow, 1, guestCount, 1).setHorizontalAlignment('center');
    sheet.getRange(guestStartRow, 2, guestCount, 1).setHorizontalAlignment('center');
    sheet.getRange(guestStartRow, 4, guestCount, 7).setHorizontalAlignment('right');
  }
  
  // 罫線を設定
  if (memberCount > 0) {
    sheet.getRange(5, 1, memberCount + 1, 11).setBorder(
      true, true, true, true, true, true,
      'black', SpreadsheetApp.BorderStyle.SOLID
    );
  }
  
  if (guestCount > 0) {
    sheet.getRange(guestStartRow - 1, 1, guestCount + 1, 11).setBorder(
      true, true, true, true, true, true,
      'black', SpreadsheetApp.BorderStyle.SOLID
    );
  }
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
 * main.jsのonOpen関数に以下を追加してください：
 * .addSeparator()
 * .addItem('スコア集計', 'aggregateScores')
 */