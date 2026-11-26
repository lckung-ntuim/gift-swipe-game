/**
 * 心動禮物大挑戰 — 把遊玩結果寫進這份 Google 試算表
 * 安裝：試算表 → 擴充功能 → Apps Script → 貼上本檔 → 部署 → 新增部署作業
 *       類型「網頁應用程式」，執行身分「我」，存取權「所有人」→ 複製網址
 */
const TZ = 'Asia/Taipei';
const SUMMARY = '總覽';
const DETAIL = '明細';
const SUMMARY_HEAD = ['Run ID', '暱稱', '開始時間', '結束時間', '總秒數', '是否完成', '作答數',
  '喜歡數', '超想要數', '心動風格', '偏好價位', '偏好品類', '心動均價', '平均思考秒數',
  '推薦品牌', '推薦商品', '推薦價格', '推薦連結', '裝置', '最後更新'];
const DETAIL_HEAD = ['Run ID', '暱稱', '開始時間', '順序', '品牌', '商品名稱', '品類', '價位', '風格',
  '價格', '選擇', '思考秒數', '操作方式', '選擇時間', '連結'];
const CHOICE = { like: '喜歡', super: '超想要', nope: '不要' };

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const run = JSON.parse(e.postData.contents);
    if (!run || !run.runId || !Array.isArray(run.choices)) throw new Error('bad payload');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sum = sheet_(ss, SUMMARY, SUMMARY_HEAD);
    const det = sheet_(ss, DETAIL, DETAIL_HEAD);
    const id = String(run.runId).slice(0, 60);
    const nick = clean_(run.nickname);
    const start = t_(run.startedAt);
    const p = run.profile || {}, r = run.recommendation || {};

    const row = [id, nick, start, t_(run.finishedAt), run.totalSeconds || '', run.completed ? '是' : '否',
      run.choices.length,
      run.choices.filter(c => c.choice !== 'nope').length,
      run.choices.filter(c => c.choice === 'super').length,
      clean_(p.topStyle), clean_(p.topTier), clean_(p.topCategory), p.avgLikedPrice || '', p.avgSeconds || '',
      clean_(r.brand), clean_(r.name), r.price || '', clean_(r.url), clean_(run.device),
      Utilities.formatDate(new Date(), TZ, 'yyyy/MM/dd HH:mm:ss')];
    upsert_(sum, id, [row]);

    const rows = run.choices.slice(0, 60).map(c => [id, nick, start, c.order, clean_(c.brand), clean_(c.name),
      clean_(c.category), clean_(c.tier), clean_(c.style), c.price, CHOICE[c.choice] || clean_(c.choice),
      c.seconds, clean_(c.method), t_(c.decidedAt), clean_(c.url)]);
    upsert_(det, id, rows);
    return ContentService.createTextOutput('ok');
  } catch (err) {
    return ContentService.createTextOutput('error: ' + err.message);
  } finally {
    lock.releaseLock();
  }
}

/* ---------- 商品清單：遊戲每次開啟會讀「商品」分頁 ---------- */
const PRODUCTS = '商品';
const PRODUCT_HEAD = ['id', '品牌', '商品名稱', '品類', '價格', '風格', 'emoji', '描述', '商品連結', '圖片網址', '啟用'];

function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'products') {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PRODUCTS);
    const out = [];
    if (sh && sh.getLastRow() > 1) {
      const vals = sh.getRange(2, 1, sh.getLastRow() - 1, PRODUCT_HEAD.length).getValues();
      vals.forEach(v => {
        if (!v[0]) return;
        out.push({ id: String(v[0]).trim(), brand: v[1], name: v[2], cat: v[3], price: v[4], style: v[5],
          emoji: v[6], desc: v[7], link: v[8], img: v[9], active: String(v[10]).trim() !== '否' });
      });
    }
    return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
  }
  return ContentService.createTextOutput('心動禮物大挑戰：紀錄端點運作中');
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🎁 禮物遊戲')
    .addItem('建立「商品」分頁', 'setupProducts')
    .addItem('依商品連結自動抓圖片', 'fillImages')
    .addToUi();
}

function setupProducts() { sheet_(SpreadsheetApp.getActiveSpreadsheet(), PRODUCTS, PRODUCT_HEAD); }

// 對「圖片網址」空白、但有「商品連結」的列，讀取商品頁的 og:image 填入
function fillImages() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PRODUCTS);
  if (!sh || sh.getLastRow() < 2) { SpreadsheetApp.getUi().alert('「商品」分頁沒有資料'); return; }
  const n = sh.getLastRow() - 1;
  const links = sh.getRange(2, 9, n, 1).getValues();
  const imgs = sh.getRange(2, 10, n, 1).getValues();
  let ok = 0, fail = [];
  for (let i = 0; i < n; i++) {
    const url = String(links[i][0]).trim();
    if (!url || imgs[i][0]) continue;
    const img = ogImage_(url);
    if (img) { imgs[i][0] = img; ok++; } else fail.push(i + 2);
  }
  sh.getRange(2, 10, n, 1).setValues(imgs);
  SpreadsheetApp.getUi().alert('抓到 ' + ok + ' 張圖片' + (fail.length ? '\n抓不到的列：' + fail.join(', ') + '（請手動貼上圖片網址）' : ''));
}

function ogImage_(url) {
  try {
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true,
      headers: { 'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 Chrome/124 Safari/537.36' } });
    if (res.getResponseCode() >= 400) return '';
    const html = res.getContentText().slice(0, 400000);
    const m = html.match(/<meta[^>]+(?:property|name)=["'](?:og:image(?::secure_url)?|twitter:image)["'][^>]*content=["']([^"']+)["']/i)
           || html.match(/<meta[^>]+content=["']([^"']+)["'][^>]*(?:property|name)=["'](?:og:image|twitter:image)["']/i);
    if (!m) return '';
    let img = m[1].replace(/&amp;/g, '&');
    if (img.startsWith('//')) img = 'https:' + img;
    if (img.startsWith('/')) img = url.match(/^https?:\/\/[^/]+/)[0] + img;
    return img.startsWith('https://') ? img : '';
  } catch (e) { return ''; }
}

function sheet_(ss, name, head) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(head);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, head.length).setFontWeight('bold').setBackground('#f6e3e8');
  }
  return sh;
}

// 刪除同一個 Run ID 的舊資料列，再寫入最新的（中途存檔→完成時會覆蓋）
function upsert_(sh, id, rows) {
  const last = sh.getLastRow();
  if (last > 1) {
    const ids = sh.getRange(2, 1, last - 1, 1).getValues();
    for (let i = ids.length - 1; i >= 0; i--) if (ids[i][0] === id) sh.deleteRow(i + 2);
  }
  if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
}

function t_(iso) { return iso ? Utilities.formatDate(new Date(iso), TZ, 'yyyy/MM/dd HH:mm:ss') : ''; }
// 避免公式注入、限制長度
function clean_(v) { const s = String(v == null ? '' : v).slice(0, 300); return /^[=+\-@]/.test(s) ? "'" + s : s; }
