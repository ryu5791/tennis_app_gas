/**
* シート情報を管理するクラス
*/
class SheetInfo {
 /**
  * 試合位置テーブルを返すgetter
  * @return {Array} 試合位置の配列 [行オフセット, 列オフセット]
  */
 static get positions() {
   return [
     [0, 0],   // 1試合目
     [4, 0],   // 2試合目
     [8, 0],   // 3試合目
     [12, 0],  // 4試合目
     [16, 0],  // 5試合目
     [20, 0],  // 6試合目
     [0, 6],   // 7試合目
     [4, 6],   // 8試合目
     [8, 6],   // 9試合目
     [12, 6],  // 10試合目
     [16, 6],  // 11試合目
     [20, 6]   // 12試合目
   ];
 }

 /**
  * 列オフセットを返すgetter
  */
 static get OFFSET_COL_DATE() {
   return 0;  // 日付列のオフセット
 }

 static get OFFSET_COL_NAME() {
   return 1;  // 会員名列のオフセット
 }

 static get OFFSET_COL_ID() {
   return 2;  // ID列のオフセット
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
  * 日付位置を返すgetter
  */
 static get OFFSET_DATE_POSITION() {
   return [0, 0];  // 日付位置のオフセット
 }

 /**
  * シートの開始位置を返すgetter
  * @param {number} page - ページ番号（0ベース）
  * @return {string} - セル参照（例: "B2"）
  */
 static get sheetPosition() {
   const positions = {
     0: "B2",    // 1ページ目の開始位置
     1: "B28",   // 2ページ目の開始位置
     2: "B54"    // 3ページ目の開始位置
   };
   
   return positions;
 }
}