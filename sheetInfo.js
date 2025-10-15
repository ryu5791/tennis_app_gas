/**
* シート情報を管理するクラス
*/
class SheetInfo {
 /**
  * 試合位置テーブルを返すgetter
  * @return {Array} 試合位置の配列 [行オフセット, 列オフセット]
  */
 static get positions() {
   const positions = [];
   const colOffsets = [0, 5, 10, 15]; // 列B,G,L,Qの相対オフセット（5列間隔）
   const rowOffsets = [0, 5, 10, 15, 20, 25, 30]; // 7試合分の行（5行間隔）
   
   // 列優先で配置（列ごとに上から下へ）
   for (let col of colOffsets) {
     for (let row of rowOffsets) {
       positions.push([row, col]);
     }
   }
   return positions; // 28試合分
 }

 /**
  * 列オフセットを返すgetter
  */
 static get OFFSET_COL_TEAM() {
   return 0;  // チーム表示列（A/B固定表示）
 }

 static get OFFSET_COL_ID() {
   return 1;  // ID列のオフセット
 }

 static get OFFSET_COL_NAME() {
   return 2;  // 会員名列のオフセット
 }

 static get OFFSET_COL_POINT() {
   return 3;  // ポイント列のオフセット
 }

 /**
  * 行オフセットを返すgetter
  */
 static get OFFSET_ROW_UPPOINT() {
   return 0;  // 上段チームのゲーム数の行オフセット
 }

 static get OFFSET_ROW_LOPOINT() {
   return 2;  // 下段チームのゲーム数の行オフセット
 }

 /**
  * 日付セルの位置を返すgetter
  */
 static get DATE_CELL() {
   return "B1";  // 日付は常にB1
 }

 /**
  * 1ゲーム当たりの行数を返すgetter
  */
 static get ROWS_PER_GAME() {
   return 4;  // チームA: 2人、チームB: 2人
 }
 
 /**
  * ゲーム間の行間隔
  */
 static get ROWS_BETWEEN_GAMES() {
   return 5;  // 各ゲーム4行 + 空白1行
 }

 /**
  * ページ情報を返すgetter
  */
 static get pageInfo() {
   return [
     {
       pageIndex: 0,
       pageName: "スコア入力",
       position: "B3",  // 最初のゲームデータ開始位置（行3）
       startGameNo: 1,
       totalGames: 28
     }
   ];
 }
}
