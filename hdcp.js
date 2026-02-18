/**
 * HDCP計算処理
 * v3.01 - 新機能追加（テスト表示）
 * v3.03 - HDCP計算ロジック実装
 * 
 * HDCPシートの列構成:
 *   A列(1)=ID, B列(2)=名前
 *   C列(3)=前々期合計, D列(4)=前々期試合数, E列(5)=前々期Gross
 *   F列(6)=前期合計,   G列(7)=前期試合数,   H列(8)=前期Gross
 *   I列(9)=2期合計,    J列(10)=2期試合数,    K列(11)=2期Gross
 *   L列(12)=前の期名,  M列(13)=今の期名(=新ﾊﾝﾃﾞｲ算出元)
 *   N列(14)=新ハンディ, O列(15)=備考
 *   P列(16)=前々期順位, Q列(17)=前期順位
 *   T列(20)=会員フラグ
 *
 * スコア集計シートの列構成:
 *   A=順位, B=会員ID, C=会員名, D=合計, E=試合数, F=Gross, G=HDCP, H=Net
 */

/**
 * HDCP計算ボタン押下時の処理
 * HDCPシートのボタンに割り当てる関数
 */
function calculateHDCP() {
  const ui = SpreadsheetApp.getUi();
  
  // 確認ダイアログ
  const response = ui.alert(
    'ハンディキャップ計算',
    'ハンディキャップ計算は半年に一度の処理です。開始しますか？',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response !== ui.Button.OK) {
    return;
  }
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hdcpSheet = ss.getSheetByName('HDCP');
    
    if (!hdcpSheet) {
      ui.alert('エラー', 'HDCPシートが見つかりません。', ui.ButtonSet.OK);
      return;
    }
    
    // ===== Step 1: HDCPシートのバックアップ =====
    const backupName = createHDCPBackup(ss, hdcpSheet);
    Logger.log(`バックアップ作成完了: ${backupName}`);
    
    // ===== Step 2: スコア集計シート・試合数加味シートのデータを取得 =====
    const scoreSheet = ss.getSheetByName('スコア集計');
    if (!scoreSheet) {
      ui.alert('エラー', 'スコア集計シートが見つかりません。\n先にスコア集計を実行してください。', ui.ButtonSet.OK);
      return;
    }
    
    const weightedSheet = ss.getSheetByName('スコア集計（試合数加味）');
    if (!weightedSheet) {
      ui.alert('エラー', 'スコア集計（試合数加味）シートが見つかりません。\n先にスコア集計を実行してください。', ui.ButtonSet.OK);
      return;
    }
    
    // スコア集計シートからID別データを辞書化（D=合計, E=試合数, F=Gross, H=Net）
    const scoreData = buildScoreDataDict(scoreSheet);
    
    // スコア集計（試合数加味）シートから上位3名を取得
    const top3 = getWeightedTop3(weightedSheet);
    
    // ===== Step 3: HDCPシートの更新処理 =====
    const lastRow = hdcpSheet.getLastRow();
    if (lastRow < 5) {
      ui.alert('エラー', 'HDCPシートのデータが不足しています。', ui.ButtonSet.OK);
      return;
    }
    
    const dataRows = lastRow - 2;   // 3行目以降の行数
    const calcRows = lastRow - 4;   // 5行目以降の行数
    
    // HDCPシートのA列(ID)を一括取得（3行目以降）
    const hdcpIds = hdcpSheet.getRange(3, 1, dataRows, 1).getValues();
    
    // --- 3a: F～H列(前期) → C～E列(前々期)にコピー（3行目以降） ---
    const prevPeriodValues = hdcpSheet.getRange(3, 6, dataRows, 3).getValues();
    hdcpSheet.getRange(3, 3, dataRows, 3).setValues(prevPeriodValues);
    
    // --- 3b: F～H列を0クリアし、スコア集計シートのデータを反映（3行目以降） ---
    const newFGH = Array.from({length: dataRows}, () => [0, 0, 0]);
    for (let i = 0; i < dataRows; i++) {
      const id = String(hdcpIds[i][0]);
      if (id === '' || id === 'null' || id === 'undefined') continue;
      if (scoreData[id]) {
        newFGH[i][0] = scoreData[id].total;     // F列: 合計
        newFGH[i][1] = scoreData[id].gameCount;  // G列: 試合数
        newFGH[i][2] = scoreData[id].gross;       // H列: Gross
      }
    }
    hdcpSheet.getRange(3, 6, dataRows, 3).setValues(newFGH);
    
    // --- 3c: I～K列の計算（5行目以降）---
    // I列=C列+F列, J列=D列+G列, K列=I列/J列
    if (calcRows > 0) {
      const cefValues = hdcpSheet.getRange(5, 3, calcRows, 6).getValues(); // C～H列 (5行目以降)
      const newIJK = [];
      for (let i = 0; i < calcRows; i++) {
        const cVal = Number(cefValues[i][0]) || 0; // C列
        const dVal = Number(cefValues[i][1]) || 0; // D列
        const fVal = Number(cefValues[i][3]) || 0; // F列
        const gVal = Number(cefValues[i][4]) || 0; // G列
        
        const iVal = cVal + fVal;
        const jVal = dVal + gVal;
        const kVal = jVal > 0 ? iVal / jVal : 0;
        
        newIJK.push([iVal, jVal, kVal]);
      }
      hdcpSheet.getRange(5, 9, calcRows, 3).setValues(newIJK);
    }
    
    // --- 3d: M列 → L列にコピー（3行目以降） ---
    const mColValues = hdcpSheet.getRange(3, 13, dataRows, 1).getValues();
    hdcpSheet.getRange(3, 12, dataRows, 1).setValues(mColValues);
    
    // --- 3e: M列3行目 = L列の次の期 ---
    const lVal = String(hdcpSheet.getRange(3, 12).getValue());
    const nextPeriod = getNextPeriod(lVal);
    hdcpSheet.getRange(3, 13).setValue(nextPeriod);
    
    // --- 3f: M列5行目以降 = (5 - K列) ---
    if (calcRows > 0) {
      const kValues = hdcpSheet.getRange(5, 11, calcRows, 1).getValues(); // K列
      const newM = kValues.map(row => [5 - (Number(row[0]) || 0)]);
      hdcpSheet.getRange(5, 13, calcRows, 1).setValues(newM);
    }
    
    // --- 3g: Q列 → P列にコピー（5行目以降） ---
    if (calcRows > 0) {
      const qValues = hdcpSheet.getRange(5, 17, calcRows, 1).getValues();
      hdcpSheet.getRange(5, 16, calcRows, 1).setValues(qValues);
    }
    
    // --- 3h: Q列にスコア集計（試合数加味）シートの上位3名を記載 ---
    if (calcRows > 0) {
      const newQ = Array.from({length: calcRows}, () => ['']);
      for (const entry of top3) {
        for (let i = 0; i < dataRows; i++) {
          if (String(hdcpIds[i][0]) === String(entry.id)) {
            // 5行目以降のインデックスに変換（3行目開始なのでi-2だが、calcRowsは5行目開始なのでi>=2のみ対象）
            if (i >= 2) {
              newQ[i - 2][0] = entry.rank;
            }
            break;
          }
        }
      }
      hdcpSheet.getRange(5, 17, calcRows, 1).setValues(newQ);
    }
    
    // --- 3i & 3j: 備考(O列)とN列（新ハンディ）の設定（5行目以降） ---
    // 重み: 1位→0.8, 2位→0.85, 3位→0.9
    const weightMap = {1: 0.8, 2: 0.85, 3: 0.9};
    
    if (calcRows > 0) {
      // 必要なデータを一括取得（5行目以降）
      const mValues = hdcpSheet.getRange(5, 13, calcRows, 1).getValues();  // M列(新ﾊﾝﾃﾞｲ算出元)
      const pValues = hdcpSheet.getRange(5, 16, calcRows, 1).getValues();  // P列
      const qValues = hdcpSheet.getRange(5, 17, calcRows, 1).getValues();  // Q列
      
      const newO = []; // 備考
      const newN = []; // 新ハンディ
      
      for (let i = 0; i < calcRows; i++) {
        const id = String(hdcpIds[i + 2] ? hdcpIds[i + 2][0] : ''); // 5行目=hdcpIds[2]
        const newHandy = Number(mValues[i][0]) || 0; // M列
        const pRank = Number(pValues[i][0]);
        const qRank = Number(qValues[i][0]);
        
        let remarks = '';
        let nVal = newHandy; // デフォルトはM列の値
        
        // P列に1～3の数字がある場合
        if (pRank >= 1 && pRank <= 3) {
          const weight = weightMap[pRank];
          nVal = newHandy * weight;
          remarks = `修正→${newHandy.toFixed(3)}×${weight}`;
        }
        
        // Q列に1～3の数字がある場合
        if (qRank >= 1 && qRank <= 3) {
          const weight = weightMap[qRank];
          const net = (id && scoreData[id]) ? (scoreData[id].net || 0) : 0;
          nVal = (newHandy - (net - 5.0)) * weight;
          if (remarks) remarks += '\n';
          remarks += `修正→{${newHandy.toFixed(3)}-(${net.toFixed(3)}-5.000)}×${weight}`;
        }
        
        newO.push([remarks]);
        newN.push([nVal]);
      }
      
      hdcpSheet.getRange(5, 15, calcRows, 1).setValues(newO); // O列(備考)
      hdcpSheet.getRange(5, 14, calcRows, 1).setValues(newN); // N列(新ハンディ)
    }
    
    // 完了メッセージ
    ui.alert(
      'ハンディキャップ計算完了',
      `ハンディキャップ計算が完了しました。\n\nバックアップ: ${backupName}\n期: ${nextPeriod}`,
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    Logger.log(`HDCP計算エラー: ${error.message}\n${error.stack}`);
    ui.alert('エラー', `HDCP計算中にエラーが発生しました。\n${error.message}`, ui.ButtonSet.OK);
  }
}

/**
 * HDCPシートのバックアップを作成
 * 既に「HDCPバックアップ」が存在する場合は末尾に(1),(2)...と添え字を付ける
 * @param {Spreadsheet} ss - スプレッドシート
 * @param {Sheet} hdcpSheet - HDCPシート
 * @return {string} 作成されたバックアップシート名
 */
function createHDCPBackup(ss, hdcpSheet) {
  const baseName = 'HDCPバックアップ';
  let backupName = baseName;
  let counter = 0;
  
  while (ss.getSheetByName(backupName)) {
    counter++;
    backupName = `${baseName}(${counter})`;
  }
  
  const backupSheet = hdcpSheet.copyTo(ss);
  backupSheet.setName(backupName);
  
  return backupName;
}

/**
 * スコア集計シートからID別データを辞書化
 * @param {Sheet} scoreSheet - スコア集計シート
 * @return {Object} ID別データ {id: {total, gameCount, gross, net}}
 */
function buildScoreDataDict(scoreSheet) {
  const lastRow = scoreSheet.getLastRow();
  if (lastRow <= 1) return {};
  
  const data = scoreSheet.getRange(2, 1, lastRow - 1, 8).getValues();
  const result = {};
  
  for (let i = 0; i < data.length; i++) {
    const id = data[i][1]; // B列 = 会員ID
    if (id === '' || id === null) continue;
    
    result[String(id)] = {
      total: Number(data[i][3]) || 0,      // D列 = 合計
      gameCount: Number(data[i][4]) || 0,   // E列 = 試合数
      gross: Number(data[i][5]) || 0,        // F列 = Gross
      net: Number(data[i][7]) || 0           // H列 = Net
    };
  }
  
  return result;
}

/**
 * スコア集計（試合数加味）シートから上位3名を取得
 * @param {Sheet} weightedSheet - スコア集計（試合数加味）シート
 * @return {Array} [{rank, id}, ...]
 */
function getWeightedTop3(weightedSheet) {
  const lastRow = weightedSheet.getLastRow();
  if (lastRow <= 1) return [];
  
  const data = weightedSheet.getRange(2, 1, lastRow - 1, 2).getValues(); // A～B列
  const result = [];
  
  for (let i = 0; i < data.length; i++) {
    const rank = Number(data[i][0]);
    if (rank >= 1 && rank <= 3) {
      result.push({
        rank: rank,
        id: data[i][1] // B列 = 会員ID
      });
    }
  }
  
  return result;
}

/**
 * 期の文字列から次の期を計算する
 * "2025年後期" → "2026年前期", "2025年前期" → "2025年後期"
 * @param {string} periodStr - 期の文字列
 * @return {string} 次の期の文字列
 */
function getNextPeriod(periodStr) {
  if (!periodStr) return '';
  
  const match = periodStr.match(/(\d{4})年(前期|後期)/);
  if (!match) {
    Logger.log(`期の解析に失敗: ${periodStr}`);
    return periodStr;
  }
  
  const year = parseInt(match[1]);
  const half = match[2];
  
  if (half === '後期') {
    return `${year + 1}年前期`;
  } else {
    return `${year}年後期`;
  }
}
