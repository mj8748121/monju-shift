// 文殊ダッシュボード連携：Comiruの欠席連絡メール → ダッシュボードの「休み連絡」
// 5分ごとに Gmail を確認し、新しい欠席連絡メールをダッシュボード（Firebase）に書き込みます。
const DB_URL = 'https://sift-sistem-default-rtdb.asia-southeast1.firebasedatabase.app';
const QUERY = 'from:noreply@comiru.jp subject:欠席連絡 newer_than:2d';

function syncComiruAbsences() {
  const props = PropertiesService.getScriptProperties();
  const done = JSON.parse(props.getProperty('done') || '[]');
  const token = ScriptApp.getOAuthToken();
  let added = 0;
  GmailApp.search(QUERY, 0, 50).forEach(function (th) {
    th.getMessages().forEach(function (m) {
      const mid = m.getId();
      if (done.indexOf(mid) >= 0) return;
      const subj = m.getSubject();
      const name = (subj.match(/文殊の(.+?)さん/) || [])[1];
      const url = (m.getPlainBody().match(/https:\/\/comiru\.jp\/\S*absence-contact\/(\d+)/) || []);
      if (!name || !url[0]) { done.push(mid); return; }
      const d = m.getDate();
      const ds = Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd');
      const rec = {
        name: name.trim(), grade: '', lesson: '', reason: '',
        via: 'Comiru', by: '自動', auto: true, url: url[0], at: d.getTime()
      };
      const res = UrlFetchApp.fetch(
        DB_URL + '/monju_dash/absences/' + ds + '/comiru_' + url[1] + '.json?access_token=' + token,
        { method: 'put', contentType: 'application/json', payload: JSON.stringify(rec), muteHttpExceptions: true });
      if (res.getResponseCode() === 200) { done.push(mid); added++; }
      else console.error(res.getResponseCode() + ' ' + res.getContentText());
    });
  });
  props.setProperty('done', JSON.stringify(done.slice(-500)));
  console.log('追加: ' + added + '件');
}

// 最初に1回だけ実行：5分ごとの自動実行を設定
function setup() {
  ScriptApp.getProjectTriggers().forEach(function (t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('syncComiruAbsences').timeBased().everyMinutes(5).create();
  syncComiruAbsences();
}
