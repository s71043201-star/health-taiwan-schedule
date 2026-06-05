/**
 * 共用 Google Sheet → GitHub 即時觸發重建
 * ------------------------------------------------
 * 安裝步驟見 SETUP_雲端試算表.md。
 *
 * 原理：試算表內容一變動（onChange），就呼叫 GitHub 的 repository_dispatch，
 * 觸發 Actions 重新抓試算表並重建網站。
 */

// === 改成你的 repo（使用者名稱/儲存庫名稱）===
var GITHUB_REPO = 's71043201-star/health-taiwan-schedule';

function onChangeRebuild(e) {
  var token = PropertiesService.getScriptProperties().getProperty('GH_TOKEN');
  if (!token) {
    throw new Error('尚未設定 GH_TOKEN（專案設定 ▸ 指令碼屬性）');
  }
  var url = 'https://api.github.com/repos/' + GITHUB_REPO + '/dispatches';
  var res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Authorization': 'Bearer ' + token,
      'Accept': 'application/vnd.github+json',
      'X-GitHub-Api-Version': '2022-11-28'
    },
    payload: JSON.stringify({ event_type: 'rebuild' }),
    muteHttpExceptions: true
  });
  Logger.log('GitHub 回應：' + res.getResponseCode());
}

/** 手動測試用：在編輯器選這個函式按「執行」，看 Actions 有沒有被觸發。 */
function testDispatch() {
  onChangeRebuild(null);
}
