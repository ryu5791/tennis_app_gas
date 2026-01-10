/**
 * アプリケーションバージョン
 * スクリプトプロパティ 'APP_VERSION' から取得
 * 未設定の場合はデフォルト値を使用
 */
const DEFAULT_VERSION = '1.0.3';

/**
 * バージョン情報を取得する
 * @return {string} - バージョン文字列
 */
function getVersion() {
  try {
    const props = PropertiesService.getScriptProperties();
    const version = props.getProperty('APP_VERSION');
    return version || DEFAULT_VERSION;
  } catch (e) {
    return DEFAULT_VERSION;
  }
}

/**
 * バージョン情報を設定する
 * @param {string} version - 設定するバージョン
 */
function setVersion(version) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('APP_VERSION', version);
  Logger.log(`バージョンを ${version} に設定しました`);
}

/**
 * バージョン情報をログに出力する
 */
function logVersion() {
  Logger.log(`テニススコア管理システム v${getVersion()}`);
}

/**
 * UI操作とログ出力を扱うヘルパークラス
 */
class UIHelper {
  // 静的プロパティは直接初期化するのではなく、
  // 静的メソッドでアクセス・設定するようにします
  static getDebugMode() {
    if (typeof UIHelper._IS_DEBUG_MODE === 'undefined') {
      UIHelper._IS_DEBUG_MODE = false; // デフォルト値
    }
    return UIHelper._IS_DEBUG_MODE;
  }
  
  /**
   * UIが利用可能であればアラートを表示する
   * @param {string} message - 表示するメッセージ
   */
  static showAlert(message) {
    // まずはログに出力
    Logger.log(message);
    
    if (!UIHelper.getDebugMode()) {
      try {
        SpreadsheetApp.getUi().alert(message);
      } catch (e) {
        Logger.log("UIの代わりにログを出力します：" + message);
      }
    } else {
      Logger.log("デバッグモード中: " + message);
    }
  }

  /**
   * ログにメッセージを出力する
   * @param {string} message - ログに出力するメッセージ
   */
  static log(message) {
    Logger.log(message);
  }

  /**
   * ログにメッセージを出力し、UIが利用可能であればアラートも表示する
   * @param {string} message - 出力するメッセージ
   */
  static alertWithLog(message) {
    this.log(message);
    this.showAlert(message);
  }
  
  /**
   * デバッグモードを設定する
   * @param {boolean} enabled - デバッグモードの有効/無効
   */
  static setDebugMode(enabled) {
    UIHelper._IS_DEBUG_MODE = enabled;
    Logger.log(`デバッグモードを${enabled ? '有効' : '無効'}にしました`);
  }
}

// 静的プロパティを初期化
UIHelper._IS_DEBUG_MODE = false;

/**
 * デバッグ用のログ出力
 * @param {string} message - ログメッセージ
 * @param {string} category - ログのカテゴリ (オプション)
 */
function debugLog(message, category = "DEBUG") {
  const timestamp = new Date().toISOString();
  Logger.log(`[${timestamp}] [${category}] ${message}`);
}

/**
 * テスト実行時のデバッグ情報出力
 * @param {string} testName - テスト名
 * @param {string} status - テストのステータス
 * @param {string} message - 追加情報
 */
function testLog(testName, status, message = "") {
  const prefix = status === "START" ? "開始" : 
                 status === "PASS" ? "成功" : 
                 status === "FAIL" ? "失敗" : 
                 status === "ERROR" ? "エラー" : "情報";
  
  debugLog(`${testName}: ${prefix} - ${message}`, "TEST");
}
