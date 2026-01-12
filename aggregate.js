/**
 * スコア集計機能
 * 期間を指定して会員とゲストのスコアを集計する
 * v1.0.6 - 書式改善、HTMLダイアログ対応
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
    
    // 参加日数閾値を入力するダイアログを表示
    const threshold = showThresholdInputDialog(startDate, endDate);
    
    if (threshold === null) {
      UIHelper.showAlert("集計をキャンセルしました。");
      return;
    }
    
    // 集計処理を実行（閾値を渡す）
    const result = executeAggregation(startDate, endDate, threshold);
    
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
 * 参加日数閾値入力ダイアログを表示
 * @param {Date} startDate - 開始日
 * @param {Date} endDate - 終了日
 * @return {number|null} 閾値（キャンセル時はnull）
 */
function showThresholdInputDialog(startDate, endDate) {
  const ui = SpreadsheetApp.getUi();
  
  // 初期値を計算（期間の月数 × 2）
  const defaultThreshold = calculateDefaultThreshold(startDate, endDate);
  
  // 閾値入力ダイアログ
  const response = ui.prompt(
    '参加日数のボーダーライン設定',
    `参加日数のボーダーラインを入力してください\n(未入力なら${defaultThreshold}日になります)`,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) {
    return null;
  }
  
  const inputText = response.getResponseText().trim();
  
  // 空の場合は初期値を使用
  if (!inputText) {
    return defaultThreshold;
  }
  
  // 数値に変換
  const threshold = parseInt(inputText, 10);
  
  if (isNaN(threshold) || threshold < 0) {
    UIHelper.showAlert("有効な数値を入力してください。");
    return null;
  }
  
  return threshold;
}

/**
 * デフォルトの参加日数閾値を計算（期間の月数 × 2）
 * @param {Date} startDate - 開始日
 * @param {Date} endDate - 終了日
 * @return {number} デフォルト閾値
 */
function calculateDefaultThreshold(startDate, endDate) {
  // 開始月と終了月を取得
  const startYear = startDate.getFullYear();
  const startMonth = startDate.getMonth();
  const endYear = endDate.getFullYear();
  const endMonth = endDate.getMonth();
  
  // 月数を計算（両端を含む）
  const monthCount = (endYear - startYear) * 12 + (endMonth - startMonth) + 1;
  
  // 月数 × 2 を返す
  return monthCount * 2;
}

/**
 * 集計処理を実行
 * @param {Date} startDate - 開始日
 * @param {Date} endDate - 終了日
 * @param {number} threshold - 参加日数閾値（ダイアログで入力された値）
 * @return {Object} 処理結果
 */
function executeAggregation(startDate, endDate, threshold) {
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
    
    // 6. 閾値はダイアログで入力された値を使用
    Logger.log(`参加日数閾値: ${threshold}日`);
    
    // 7. 閾値以上の会員のみを抽出（前期/前々期順位があっても日数不足なら対象外）
    const qualifiedMembers = members
      .filter(m => m.participationDays >= threshold)
      .sort((a, b) => b.net - a.net);
    Logger.log(`閾値以上の会員: ${qualifiedMembers.length}名`);
    
    // 8. スコア集計シートに出力（全会員を出力）
    outputToSheet(members, guests, threshold, startDate, endDate);
    
    // 9. スコア集計（試合数加味）シートを作成
    outputWeightedSheet(qualifiedMembers, members, guests, threshold, startDate, endDate);
    
    // 参加日数シートはインプット用なので書き込みは行わない
    
    // 会員数はスコア集計シートのゲスト以外の人数（全会員数）を返す
    return {
      success: true,
      memberCount: members.length,
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
  // シート構造: 4行目がヘッダー、6行目からデータ
  // A列: ID, Q列(17列目): 参加日数合計
  const lastRow = sheet.getLastRow();
  const dataStartRow = 6;  // データ開始行
  
  if (lastRow < dataStartRow) {
    return {};
  }
  
  // A列(ID)とQ列(参加日数合計)を取得
  const numRows = lastRow - dataStartRow + 1;
  const idData = sheet.getRange(dataStartRow, 1, numRows, 1).getValues();    // A列
  const daysData = sheet.getRange(dataStartRow, 17, numRows, 1).getValues(); // Q列(17列目)
  const result = {};
  
  // 配列を一度だけループして辞書を構築
  for (let i = 0; i < numRows; i++) {
    const id = idData[i][0];
    const days = daysData[i][0];
    if (id !== '' && id !== null && days !== '' && days !== null) {
      result[id] = Number(days);
    }
  }
  
  Logger.log(`参加日数シートから ${Object.keys(result).length} 件のデータを取得しました（Q列参照）`);
  
  return result;
}

/**
 * HDCP情報を取得（辞書型で返す）
 * v1.0.6: P列(前々期)、Q列(前期)の取得を追加
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
  
  // A列=ID, N列(14列目)=新ハンディ, O列(15列目)=備考, P列(16列目)=前々期, Q列(17列目)=前期, T列(20列目)=会員フラグ
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
    const prevPrevPeriod = row[15] || '';  // P列（インデックス15）前々期
    const prevPeriod = row[16] || '';       // Q列（インデックス16）前期
    
    result[id] = {
      isMember: isMember,
      hdcp: Number(hdcp),
      remarks: String(remarks).trim(),
      prevPrevPeriod: String(prevPrevPeriod).trim(),  // 前々期（青色で表示）
      prevPeriod: String(prevPeriod).trim()           // 前期（赤色で表示）
    };
    
    Logger.log(`ID: ${id}, T列の値: ${memberFlag}, 会員判定: ${isMember}`);
  });
  
  return result;
}

/**
 * プレイヤーを会員とゲストに分類
 * v1.0.6: 前期/前々期情報の追加
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
    const hdcpInfo = hdcpData[id] || { isMember: false, hdcp: 0, remarks: '', prevPrevPeriod: '', prevPeriod: '' };
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
      remarks: hdcpInfo.remarks || '',
      prevPrevPeriod: hdcpInfo.prevPrevPeriod || '',  // 前々期（青色で表示）
      prevPeriod: hdcpInfo.prevPeriod || ''           // 前期（赤色で表示）
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
    .setBackground('#6AA84F')  // 緑色に変更
    .setFontColor('white')
    .setFontWeight('bold');
  currentRow++;
  
  // 会員データをNet降順でソート（非破壊）
  const sortedMembers = [...members].sort((a, b) => b.net - a.net);
  
  // 全会員の中でグロス最大のIDを特定
  let maxGross = -Infinity;
  let maxGrossId = null;
  sortedMembers.forEach(member => {
    if (member.gross > maxGross) {
      maxGross = member.gross;
      maxGrossId = member.id;
    }
  });
  
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
    
    // グロス1位の場合：F列を赤色+斜体
    if (member.id == maxGrossId) {
      outputSheet.getRange(currentRow, 6)
        .setFontColor('#FF0000')
        .setFontStyle('italic');
    }
    
    // グロス順位がある場合：I列を赤色+斜体
    if (member.grossRank) {
      outputSheet.getRange(currentRow, 9)
        .setFontColor('#FF0000')
        .setFontStyle('italic');
    }
    
    currentRow++;
  });
  
  // ゲストデータをNet降順でソート（非破壊）
  const sortedGuests = [...guests].sort((a, b) => b.net - a.net);
  
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
 * v1.0.6: 画像仕様に合わせた全面書き換え
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
  
  // ===== グロス1位のIDを特定 =====
  let maxGrossId = null;
  let maxGross = -Infinity;
  allMembers.forEach(m => {
    if (m.gross > maxGross) {
      maxGross = m.gross;
      maxGrossId = m.id;
    }
  });
  
  let currentRow = 1;
  
  // ===== 行1: ヘッダー行（緑背景 #6AA84F、白文字）=====
  const headers = ['順位', '会員ＩＤ', '会員名', '合計', '試合数', 'Gross', 'HDCP', 'Net', 'ｸﾞﾛｽ順位', '参加日数', '備考'];
  weightedSheet.getRange(currentRow, 1, 1, headers.length).setValues([headers]);
  weightedSheet.getRange(currentRow, 1, 1, headers.length)
    .setBackground('#6AA84F')
    .setFontColor('white')
    .setFontWeight('bold');
  currentRow++;
  
  // ===== 行2〜: 閾値以上の会員をNet順降順でソート =====
  const sortedQualified = [...qualifiedMembers].sort((a, b) => b.net - a.net);
  const qualifiedStartRow = currentRow;
  
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
      ''  // 備考は後でリッチテキストで設定
    ];
    weightedSheet.getRange(currentRow, 1, 1, row.length).setValues([row]);
    
    // グロス1位の場合：F列を赤色+斜体
    if (member.id == maxGrossId) {
      weightedSheet.getRange(currentRow, 6)
        .setFontColor('#FF0000')
        .setFontStyle('italic');
    }
    
    // グロス順位がある場合：I列を赤色+斜体
    if (member.grossRank) {
      weightedSheet.getRange(currentRow, 9)
        .setFontColor('#FF0000')
        .setFontStyle('italic');
    }
    
    // 備考欄（K列）のリッチテキスト設定
    setRemarksCell(weightedSheet, currentRow, 11, member.prevPrevPeriod, member.prevPeriod);
    
    currentRow++;
  });
  
  // ===== 閾値未満の会員を抽出してNet順降順でソート =====
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
      ''  // 備考は後でリッチテキストで設定
    ];
    weightedSheet.getRange(currentRow, 1, 1, row.length).setValues([row]);
    
    // グロス1位の場合：F列を赤色+斜体
    if (member.id == maxGrossId) {
      weightedSheet.getRange(currentRow, 6)
        .setFontColor('#FF0000')
        .setFontStyle('italic');
    }
    
    // グロス順位がある場合：I列を赤色+斜体
    if (member.grossRank) {
      weightedSheet.getRange(currentRow, 9)
        .setFontColor('#FF0000')
        .setFontStyle('italic');
    }
    
    // 備考欄（K列）のリッチテキスト設定
    setRemarksCell(weightedSheet, currentRow, 11, member.prevPrevPeriod, member.prevPeriod);
    
    currentRow++;
  });
  
  const memberEndRow = currentRow - 1;
  
  // ===== 規定日数行（緑背景 #D9EAD3、緑文字 #38761D、緑点線枠）=====
  const noteRowIndex = currentRow;
  weightedSheet.getRange(currentRow, 3).setValue('規定日数');
  weightedSheet.getRange(currentRow, 4).setValue('＝');
  weightedSheet.getRange(currentRow, 5).setValue(`日数 ${threshold}日以上をランキング対象とする`);
  
  // 規定日数行の書式
  weightedSheet.getRange(currentRow, 1, 1, 11)
    .setBackground('#D9EAD3')
    .setFontColor('#38761D');
  
  // 緑点線枠
  weightedSheet.getRange(currentRow, 1, 1, 11).setBorder(
    true, true, true, true, null, null,
    '#93C47D', SpreadsheetApp.BorderStyle.DASHED
  );
  currentRow++;
  
  // ===== ゲストセクション =====
  const sortedGuests = [...guests].sort((a, b) => b.net - a.net);
  const guestStartRow = currentRow;
  
  // ゲストデータを出力（A列に「ゲスト」赤文字）
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
      ''
    ];
    weightedSheet.getRange(currentRow, 1, 1, row.length).setValues([row]);
    
    // A列「ゲスト」を赤色に
    weightedSheet.getRange(currentRow, 1).setFontColor('#FF0000');
    
    // 備考欄（K列）のリッチテキスト設定
    setRemarksCell(weightedSheet, currentRow, 11, guest.prevPrevPeriod, guest.prevPeriod);
    
    currentRow++;
  });
  
  const guestEndRow = currentRow - 1;
  
  // ===== 書式設定 =====
  formatWeightedSheetV106(weightedSheet, qualifiedStartRow, memberEndRow, noteRowIndex, guestStartRow, guestEndRow, sortedQualified.length, sortedUnqualified.length, sortedGuests.length);
}

/**
 * 備考欄にリッチテキストで前々期（青）・前期（赤）を表示
 * 形式: 「前々期○位」「前期○位」
 * @param {Sheet} sheet - シートオブジェクト
 * @param {number} row - 行番号
 * @param {number} col - 列番号
 * @param {string} prevPrevPeriod - 前々期順位（数値のみ）
 * @param {string} prevPeriod - 前期順位（数値のみ）
 */
function setRemarksCell(sheet, row, col, prevPrevPeriod, prevPeriod) {
  const cell = sheet.getRange(row, col);
  
  // 順位テキストを生成
  const prevPrevText = prevPrevPeriod ? `前々期${prevPrevPeriod}位` : '';
  const prevText = prevPeriod ? `前期${prevPeriod}位` : '';
  
  if (prevPrevText && prevText) {
    // 両方ある場合はリッチテキストで色分け
    const text = prevPrevText + '\n' + prevText;
    const richText = SpreadsheetApp.newRichTextValue()
      .setText(text)
      .setTextStyle(0, prevPrevText.length, SpreadsheetApp.newTextStyle().setForegroundColor('#0000FF').build())
      .setTextStyle(prevPrevText.length + 1, text.length, SpreadsheetApp.newTextStyle().setForegroundColor('#FF0000').build())
      .build();
    cell.setRichTextValue(richText);
  } else if (prevPrevText) {
    // 前々期のみの場合は青色
    cell.setValue(prevPrevText).setFontColor('#0000FF');
  } else if (prevText) {
    // 前期のみの場合は赤色
    cell.setValue(prevText).setFontColor('#FF0000');
  }
}

/**
 * スコア集計（試合数加味）シートの書式を設定（v1.0.6版）
 * @param {Sheet} sheet - 対象シート
 * @param {number} qualifiedStartRow - 閾値以上会員開始行
 * @param {number} memberEndRow - 会員終了行
 * @param {number} noteRowIndex - 規定日数行
 * @param {number} guestStartRow - ゲスト開始行
 * @param {number} guestEndRow - ゲスト終了行
 * @param {number} qualifiedCount - 閾値以上会員数
 * @param {number} unqualifiedCount - 閾値未満会員数
 * @param {number} guestCount - ゲスト数
 */
function formatWeightedSheetV106(sheet, qualifiedStartRow, memberEndRow, noteRowIndex, guestStartRow, guestEndRow, qualifiedCount, unqualifiedCount, guestCount) {
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
  
  // 会員データ（ヘッダー含む）の書式
  const totalMemberRows = memberEndRow;
  if (totalMemberRows > 0) {
    // ヘッダー + 会員データに罫線
    sheet.getRange(1, 1, totalMemberRows, 11).setBorder(
      true, true, true, true, true, true,
      'black', SpreadsheetApp.BorderStyle.SOLID
    );
    
    // 数値列の配置
    if (memberEndRow >= qualifiedStartRow) {
      sheet.getRange(qualifiedStartRow, 1, memberEndRow - qualifiedStartRow + 1, 1).setHorizontalAlignment('center');  // 順位
      sheet.getRange(qualifiedStartRow, 2, memberEndRow - qualifiedStartRow + 1, 1).setHorizontalAlignment('center');  // ID
      sheet.getRange(qualifiedStartRow, 4, memberEndRow - qualifiedStartRow + 1, 7).setHorizontalAlignment('right');   // 数値列
    }
  }
  
  // ゲストデータの罫線（ゲストのみ、ヘッダーなし）
  if (guestCount > 0 && guestEndRow >= guestStartRow) {
    sheet.getRange(guestStartRow, 1, guestEndRow - guestStartRow + 1, 11).setBorder(
      true, true, true, true, true, true,
      'black', SpreadsheetApp.BorderStyle.SOLID
    );
    
    sheet.getRange(guestStartRow, 1, guestEndRow - guestStartRow + 1, 1).setHorizontalAlignment('center');
    sheet.getRange(guestStartRow, 2, guestEndRow - guestStartRow + 1, 1).setHorizontalAlignment('center');
    sheet.getRange(guestStartRow, 4, guestEndRow - guestStartRow + 1, 7).setHorizontalAlignment('right');
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
  const memberStartRow = 2;
  const guestStartRow = memberStartRow + memberCount;
  const totalDataRows = memberCount + guestCount;
  
  // C列（会員名）にHG丸ゴシックM-PROまたは代替フォントを設定
  // Google Sheetsで利用可能なフォントに近いものを使用
  if (totalDataRows > 0) {
    sheet.getRange(memberStartRow, 3, totalDataRows, 1).setFontFamily('M PLUS Rounded 1c');
  }
  
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
  const totalRows = 1 + memberCount + guestCount;
  if (totalRows > 1) {
    sheet.getRange(1, 1, totalRows, 11).setBorder(
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
