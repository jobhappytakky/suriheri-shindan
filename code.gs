/* ============================================================
   「職場の人に疲れた」のすり減り診断（7号機）
   設定配信＆ログ記録用 GAS
   設置手順：
     1. 新規スプレッドシートを作成
     2. 拡張機能 → Apps Script にこのコードを貼り付け
     3. setup() を1回実行（シート3枚と初期値が作られる）
     4. デプロイ → ウェブアプリ（実行ユーザー: 自分／アクセス: 全員）
     5. 発行されたURLを index.html と demo.html の GAS_URL に貼る
============================================================ */

const SHEET_LINKS = 'リンク設定';
const SHEET_CARDS = 'カード画像';
const SHEET_LOG   = 'ログ';

function doGet(e) {
  const out = { cta: {}, cards: {} };
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ls = ss.getSheetByName(SHEET_LINKS);
    if (ls) {
      const rows = ls.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        const [key, label, url, thumb, show] = rows[i];
        if (!key) continue;
        out.cta[String(key)] = {
          label: String(label || ''),
          url: String(url || ''),
          thumb: String(thumb || ''),
          show: String(show).toUpperCase() !== 'FALSE'
        };
      }
    }
    const cs = ss.getSheetByName(SHEET_CARDS);
    if (cs) {
      const rows = cs.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        const [name, url] = rows[i];
        if (name && url) out.cards[String(name)] = String(url);
      }
    }
  } catch (err) {}
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const p = JSON.parse(e.postData.contents || '{}');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const log = ss.getSheetByName(SHEET_LOG) || ss.insertSheet(SHEET_LOG);
    log.appendRow([
      new Date(),
      sanitize(p.event),
      sanitize(p.main),
      sanitize(p.sub),
      sanitize(p.card)
    ]);
  } catch (err) {}
  return ContentService.createTextOutput('ok');
}

/* 数式インジェクション対策＋50字制限 */
function sanitize(v) {
  let s = String(v == null ? '' : v).slice(0, 50);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

/* 初期セットアップ（1回だけ実行） */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const URL3 = 'https://coconala.com/services/3506681'; // ③職場の人間関係

  let ls = ss.getSheetByName(SHEET_LINKS) || ss.insertSheet(SHEET_LINKS);
  ls.clear();
  ls.getRange(1, 1, 8, 5).setValues([
    ['キー', 'ボタン文言', 'リンク先URL', 'サムネ画像URL', '表示'],
    ['kikubari', 'この気疲れの先をみてもらう', URL3, '', 'TRUE'],
    ['itabasami', '板挟みのほどき方をみてもらう', URL3, '', 'TRUE'],
    ['ukezara', 'あの人の本音をみてもらう', URL3, '', 'TRUE'],
    ['honne', '本音の伝え方をみてもらう', URL3, '', 'TRUE'],
    ['namae', 'この疲れの正体をみてもらう', URL3, '', 'TRUE'],
    ['marugoto', 'まるごとみてもらう', URL3, '', 'TRUE'],
    ['sekai', '', '', '', 'FALSE']
  ]);

  const cards = ['愚者','魔術師','女教皇','女帝','皇帝','教皇','恋人','戦車','力','隠者',
    '運命の輪','正義','吊るされた男','死神','節制','悪魔','塔','星','月','太陽','審判','世界'];
  let cs = ss.getSheetByName(SHEET_CARDS) || ss.insertSheet(SHEET_CARDS);
  cs.clear();
  cs.getRange(1, 1, 1, 2).setValues([['カード名', '画像URL（空ならリポジトリ同梱画像を使用）']]);
  cs.getRange(2, 1, cards.length, 1).setValues(cards.map(c => [c]));

  let log = ss.getSheetByName(SHEET_LOG) || ss.insertSheet(SHEET_LOG);
  if (log.getLastRow() === 0) {
    log.getRange(1, 1, 1, 5).setValues([['日時', 'event', 'main', 'sub', 'card']]);
  }
}
