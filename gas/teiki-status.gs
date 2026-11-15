// 文殊ダッシュボード連携：「定期テスト回収表」→ ダッシュボードの「定期テスト回収状況」
// 30分ごとに回収表を読み、学校×学年×教科の回収状況と未回収の生徒をダッシュボード（Firebase）に書き込みます。
// 最新のテストは「生徒別回収_YYYY年M月」タブの名前から自動で選びます（次のテストのタブを作れば自動で切り替わります）。
const DB_URL = 'https://sift-sistem-default-rtdb.asia-southeast1.firebasedatabase.app';
const SS_ID = '1_tcmbASqGguPBmyfuRHSU6VDd5CeF42khuPQ_tbGSVk';
const SUBJ = ['国語', '数学', '英語', '理科', '社会'];

function syncTeikiStatus() {
  const ss = SpreadsheetApp.openById(SS_ID);
  const ym = function (n) { const m = n.match(/(\d{4})年(\d{1,2})月/); return m ? Number(m[1]) * 100 + Number(m[2]) : 0; };
  const tabs = ss.getSheets().map(function (s) { return s.getName(); })
    .filter(function (n) { return /^生徒別回収_\d{4}年\d{1,2}月$/.test(n); })
    .sort(function (a, b) { return ym(b) - ym(a); });
  if (!tabs.length) { console.log('生徒別回収タブがありません'); return; }
  const test = tabs[0].replace('生徒別回収_', '');

  // 学校×学年のタブ（例「2026年11月」）：問題・解答作成・解答けし・問題用紙の部数
  const t = ss.getSheetByName(test).getDataRange().getValues();
  const label = String(t[1][0] || '').split(/[　\s]/)[0];
  const cols = {};
  const keyOf = { '問題': 'mondai', '解答作成': 'kaitou', '解答用紙': 'touan', '解答けし': 'hakushi' };
  let cur = null;
  for (let c = 1; c < t[2].length; c++) {
    const h3 = String(t[2][c]).trim(), h4 = String(t[3][c]).trim();
    if (/^問題用紙/.test(h3)) cur = '#bu';
    else if (SUBJ.indexOf(h3) >= 0) cur = h3;
    if (!cur) continue;
    if (cur === '#bu') { if (SUBJ.indexOf(h4) >= 0) (cols[h4] = cols[h4] || {}).bu = c; }
    else if (keyOf[h4]) { cols[cur] = cols[cur] || {}; if (cols[cur][keyOf[h4]] === undefined) cols[cur][keyOf[h4]] = c; }
  }
  const rows = {};
  for (let r = 4; r < t.length; r++) {
    const m = String(t[r][0]).trim().match(/^(仲西|神森)(\d)年$/);
    if (!m) continue;
    const row = { school: m[1], grade: Number(m[2]), enrolled: 0, subj: {}, missing: [] };
    SUBJ.forEach(function (s) {
      const c = cols[s] || {};
      const v = function (k) { return c[k] === undefined ? '' : t[r][c[k]]; };
      row.subj[s] = { got: 0, mondai: v('mondai') === true, kaitou: v('kaitou') === true, hakushi: v('hakushi') === true, bu: String(v('bu') || '') };
    });
    rows[m[1] + m[2]] = row;
  }

  // 生徒別：回収数と未回収の生徒
  const st = ss.getSheetByName(tabs[0]).getDataRange().getValues();
  const head = st[2];
  const ci = function (name) { return head.indexOf(name); };
  let collected = 0, target = 0;
  for (let r = 3; r < st.length; r++) {
    const name = String(st[r][ci('氏名')] || '').trim();
    const row = rows[String(st[r][ci('学校')]).trim() + String(st[r][ci('学年')]).trim()];
    if (!name || !row) continue;
    row.enrolled++;
    const miss = [];
    SUBJ.forEach(function (s) {
      target++;
      if (st[r][ci(s)] !== '' && st[r][ci(s)] !== null) { row.subj[s].got++; collected++; }
      else miss.push(s);
    });
    if (miss.length) row.missing.push({ n: name.replace(/[\s　]+/g, ' '), m: miss.length === 5 ? ['全教科'] : miss });
  }

  const out = {
    test: test, label: label, updated: Date.now(), collected: collected, target: target,
    rows: Object.keys(rows).map(function (k) { return rows[k]; })
  };
  const res = UrlFetchApp.fetch(DB_URL + '/monju_dash/teiki.json?access_token=' + ScriptApp.getOAuthToken(),
    { method: 'put', contentType: 'application/json', payload: JSON.stringify(out), muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) throw new Error(res.getResponseCode() + ' ' + res.getContentText());
  console.log(test + '：' + collected + '/' + target + ' を書き込みました');
}

// 最初に1回だけ実行：30分ごとの自動実行を設定
function setup() {
  ScriptApp.getProjectTriggers().forEach(function (t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('syncTeikiStatus').timeBased().everyMinutes(30).create();
  syncTeikiStatus();
}
