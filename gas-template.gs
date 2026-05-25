// ================================================
// ランボルギーニ 受発注システム GAS（最終版）
// ================================================
const LINE_TOKEN    = 'cGQcVpix9Q9nYIeMF8iHTpqsrt5EV0zQbZHQ1h2ikTpSSZGe8jXJKrP8/DREH5Yn0CtTR4/UoKgHkAKtz4Ku7Kt1XiuCLHXOvTbNFUvP3uRf9R90OD5XyQoyLbqein2+c+p6PSOPJj2uMpW8ooVrUAdB04t89/1O/w1cDnyilFU=';
const SHEET_ID      = '15PDRjw1QSJi0mXxpPNGHHvu--HTFJ3q-0bEsDpFLJbs';
const ADMIN_USER_ID = 'Uf9fead1dd3b45f150bc2c1a9b6740223';
// ↓ スタッフのLINE UIDをここに追加（在庫確認などスタッフ限定機能を使えるUID）
const STAFF_UIDS    = [
  'Uf9fead1dd3b45f150bc2c1a9b6740223', // 管理者（ADMIN_USER_IDと同じ）
  // 'Uxxxxxxxxxxxxxxxxxxxxxxxxxx',     // ← スタッフ追加時はここに記入
];
function isStaff(uid) { return STAFF_UIDS.includes(uid); }

function doGet(e) {
  const action = e.parameter.action;

  if (action === 'ping') return json({ status: 'ok' });

  // ─── 新規注文取得（管理画面ポーリング用） ──────
  if (action === 'getNewOrders') {
    const since = e.parameter.since || '';
    return json({ orders: getNewOrders(since) });
  }

  if (action === 'getOrders') {
    const rows = getSheet('注文受付').getDataRange().getValues();
    return json({ orders: rows.slice(1).map(r => ({
      date:r[0], branch:r[1], staff:r[2], store:r[3], area:r[4],
      items:r[5], amount:r[6], payMethod:r[7], payStatus:r[8],
      status:r[9], note:r[10], uid:r[11]
    }))});
  }

  if (action === 'getStock') {
    const rows = getSheet('在庫').getDataRange().getValues();
    return json({ stock: rows.map(r => ({ branch:r[0], product:r[1], qty:r[2] })) });
  }

  if (action === 'checkLineUser') {
    const uid = e.parameter.uid;
    if (!uid) return json({ ok:false, error:'uid missing' });
    const rows = getSheet('顧客').getDataRange().getValues();
    if (rows.length === 0) return json({ ok:false });
    const hdrs = rows[0].map(String);

    // ★ BS列(index 70)から動的列開始（詳細情報・備考・商品別掛率が混在）
    //   支払方法は R列(index 17)、配達先は Z〜AR列(25〜69) に固定。
    const prodRateStartIdx = 70;  // BS列から動的列開始
    // 詳細情報・備考のヘッダー名（商品別掛率と区別するため・支払方法は固定列なので除外）
    const DETAIL_HEADERS = ['業態','開店年月','席数','経営者年代','主要客層',
                             '主力商品傾向','競合関係','取引開始日','取引区分',
                             '営業日時','平均客単価','価格モード','税','備考'];

    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][14]) !== uid) continue;      // O列: LINE UID

      // ── 配達先5件（Z〜AR列、インデックス25〜69）────
      const deliveries = [];
      for (let n = 0; n < 5; n++) {
        const base = 25 + n * 9;
        const d = {
          name:    String(rows[i][base]   || ''),
          zip:     String(rows[i][base+1] || ''),
          pref:    String(rows[i][base+2] || ''),
          city:    String(rows[i][base+3] || ''),
          addr1:   String(rows[i][base+4] || ''),
          addr2:   String(rows[i][base+5] || ''),
          contact: String(rows[i][base+6] || ''),
          tel:     String(rows[i][base+7] || ''),
          memo:    String(rows[i][base+8] || '')
        };
        if (d.name || d.zip || d.addr1 || d.contact) {
          deliveries.push(d);
        }
      }

      // ── 商品別掛率（BS列以降、詳細情報・備考を除く）──
      const pr = [];
      for (let col = prodRateStartIdx; col < hdrs.length; col++) {
        if (DETAIL_HEADERS.indexOf(hdrs[col]) >= 0) continue; // 詳細情報・備考はスキップ
        if (hdrs[col] && rows[i][col] !== '' && rows[i][col] !== undefined) {
          pr.push({ key: hdrs[col], rate: String(rows[i][col]) });
        }
      }

      // ── 共通掛率：商品別掛率から代表値を取得 ──────
      // イエロー掛率を優先、なければ最初の商品別掛率、なければデフォルト7
      let defaultRate = '7';
      const yellowRate = pr.find(function(r) { return r.key === 'イエロー'; });
      if (yellowRate && yellowRate.rate) {
        defaultRate = yellowRate.rate;
      } else if (pr.length > 0 && pr[0].rate) {
        defaultRate = pr[0].rate;
      }

      // ── レスポンス ────────────────────────────────
      return json({
        ok:         true,
        branch:     rows[i][0],   // A: 拠点
        staff:      rows[i][1],   // B: 弊社担当
        name:       rows[i][2],   // C: 店舗名
        kana:       rows[i][3],   // D: フリガナ
        tel:        rows[i][4],   // E: 店舗電話
        owner:      rows[i][5],   // F: 代表者名
        ownerTel:   rows[i][6],   // G: 代表者電話
        zip:        rows[i][8],   // I: 郵便番号
        pref:       rows[i][9],   // J: 都道府県
        city:       rows[i][10],  // K: 市町村郡
        addr1:      rows[i][11],  // L: 番地
        addr2:      rows[i][12],  // M: 建物名
        lineId:     rows[i][13],  // N: LINE ID
        contact:    (()=>{ try{ const c=JSON.parse(rows[i][7]||'[]'); return c[0]?.name||''; }catch(e){return '';} })(),
        contactTel: (()=>{ try{ const c=JSON.parse(rows[i][7]||'[]'); return c[0]?.tel||''; }catch(e){return '';} })(),
        rate:       defaultRate,
        productRates: pr,
        payMethod:  rows[i][17] || '',   // R列: 支払方法（固定）
        deliveries: deliveries,          // 配達先5件（空除外）
      });
    }
    return json({ ok:false });
  }

  if (action === 'getCustomers') {
    const rows = getSheet('顧客').getDataRange().getValues();
    if (rows.length === 0) return json({ customers: [] });
    const hdrs = rows[0].map(String);
    // 詳細情報のヘッダー名 → JSプロパティ名 のマッピング（支払方法は固定列なので除外）
    const DETAIL_KEYS = {
      '業態':         'bizType',
      '開店年月':     'openDate',
      '席数':         'seats',
      '経営者年代':   'ownerAge',
      '主要客層':     'customerBase',
      '主力商品傾向': 'productTrend',
      '競合関係':     'competitor',
      '取引開始日':   'dealStart',
      '取引区分':     'dealChannel',
      '営業日時':     'bizHours',
      '平均客単価':   'avgSpend',
      '価格モード':   'priceMode',
      '税':           'taxType'
    };
    return json({ customers: rows.slice(1).map(r => {
      const pr = [];
      const details = {};
      let note = '';
      // 配達先5件をZ〜AR列（インデックス25〜69）から読み取り
      const deliveries = [];
      for (let n = 0; n < 5; n++) {
        const base = 25 + n * 9;
        const d = {
          name:    String(r[base]   || ''),
          zip:     String(r[base+1] || ''),
          pref:    String(r[base+2] || ''),
          city:    String(r[base+3] || ''),
          addr1:   String(r[base+4] || ''),
          addr2:   String(r[base+5] || ''),
          contact: String(r[base+6] || ''),
          tel:     String(r[base+7] || ''),
          memo:    String(r[base+8] || '')
        };
        // いずれかのフィールドに値があれば配達先として保持
        if (d.name || d.zip || d.addr1 || d.contact) {
          deliveries.push(d);
        }
      }
      // BS列以降（インデックス70〜）を「詳細情報」「備考」「商品別掛率」に分離
      for (let col = 70; col < hdrs.length; col++) {
        const h = hdrs[col];
        if (h === '備考') {
          note = (r[col] !== undefined && r[col] !== null) ? String(r[col]) : '';
        } else if (DETAIL_KEYS[h]) {
          details[DETAIL_KEYS[h]] = (r[col] !== undefined && r[col] !== null) ? String(r[col]) : '';
        } else if (h && r[col] !== '' && r[col] !== undefined) {
          pr.push({ key: h, rate: String(r[col]) });
        }
      }
      return Object.assign({
        branch:r[0], staff:r[1], name:r[2], kana:r[3], tel:r[4],
        owner:r[5], ownerTel:r[6], contacts:r[7],
        zip:r[8], pref:r[9], city:r[10], addr1:r[11], addr2:r[12],
        lineId:r[13], lineUid:r[14],
        referrer:r[15], invName:r[16],
        payMethod: r[17] || '',                              // R列: 支払方法（固定）
        bankNames:[r[18],r[19],r[20]].filter(Boolean), bankName:r[18]||'',
        note: note,
        deliveries: deliveries,                              // 配達先5件（空は除外）
        productRates: pr,
        rate: pr.length > 0 ? (pr[0].rate || '') : '',
        sponsorEnabled: r[21] === 'あり',  // V列: 協賛
        sponsorTrigger: r[22] || '',       // W列: 購入条件
        sponsorReward:  r[23] || '',       // X列: 協賛本数
        sponsorProduct: r[24] || ''        // Y列: 対象商品
      }, details);
    })});
  }

  // ★ ブラウザno-cors対策：GETのpayloadパラメータで書き込みを受け取る
  if (e.parameter.payload) {
    try {
      const body = JSON.parse(e.parameter.payload);
      if (body.action === 'saveCustomer')    { saveCustomerToSheet(body); return json({ok:true}); }
      if (body.action === 'saveOrder')       { saveOrderToSheet(body);    return json({ok:true}); }
      if (body.action === 'updatePayStatus') { updatePayInSheet(body.orderId, body.payStatus); return json({ok:true}); }
      if (body.action === 'updateStock')     { updateStockInSheet(body.branch, body.product, body.qty); return json({ok:true}); }
      if (body.action === 'updateAllStock')  { updateAllStockInSheet(body.stockData); return json({ok:true}); }
    } catch(err) {
      return json({ error: 'parse error: ' + err.message });
    }
  }

  if (action === 'getPendingUids') {
    const rows = getSheet('未登録UID').getDataRange().getValues();
    return json({ pending: rows.slice(1).map(r => ({
      date:r[0], uid:r[1], name:r[2], message:r[3], status:r[4]
    }))});
  }

  // 顧客自身の支払い状況照会（LIFF用）
  if (action === 'myPayments') {
    const uid = e.parameter.uid;
    if (!uid) return json({ error: 'uid required' });
    // 顧客名取得
    const custRows = getSheet('顧客').getDataRange().getValues();
    const custRow  = custRows.slice(1).find(r => String(r[9]) === uid);
    const storeName = custRow ? custRow[0] : '';
    // 注文データ取得
    const ordRows = getSheet('注文受付').getDataRange().getValues();
    const mine = ordRows.slice(1).filter(r => String(r[11]) === uid);
    const unpaid = mine.filter(r => r[8] === '未収' || r[8] === '売掛').map(r => ({
      date: String(r[0]).substring(0,10), items: r[5], amount: r[6]||0, payStatus: r[8], payMethod: r[7]
    }));
    const paid = mine.filter(r => r[8] === '集金済み').slice(-5).map(r => ({
      date: String(r[0]).substring(0,10), items: r[5], amount: r[6]||0, payStatus: r[8]
    }));
    const totalUnpaid = unpaid.reduce((s,o) => s + Number(o.amount||0), 0);
    return json({ ok:true, storeName, unpaid, paid, totalUnpaid });
  }

  // 顧客自身の注文履歴照会（LIFF用）
  if (action === 'myOrders') {
    const uid = e.parameter.uid;
    if (!uid) return json({ error: 'uid required' });
    const ordRows = getSheet('注文受付').getDataRange().getValues();
    const mine = ordRows.slice(1).filter(r => String(r[11]) === uid).slice(-20).reverse().map(r => ({
      date: String(r[0]).substring(0,10), items: r[5], amount: r[6]||0, payStatus: r[8], status: r[9]
    }));
    return json({ ok:true, orders: mine });
  }

  // 顧客にUIDを紐付け
  if (action === 'linkUid') {
    const uid       = e.parameter.uid;
    const storeName = e.parameter.store;
    if (!uid || !storeName) return json({ error: 'uid and store required' });
    const sh = getSheet('顧客');
    const d  = sh.getDataRange().getValues();
    for (let i = 1; i < d.length; i++) {
      if (d[i][0] === storeName) {
        sh.getRange(i+1, 10).setValue(uid); // J列: LINE UID
        // 未登録UIDシートから該当行を削除
        const ps = getSheet('未登録UID');
        const pd = ps.getDataRange().getValues();
        for (let j = 1; j < pd.length; j++) {
          if (pd[j][1] === uid) { ps.deleteRow(j+1); break; }
        }
        return json({ ok: true });
      }
    }
    return json({ error: '顧客が見つかりません' });
  }

  // 未登録UIDをスプシから削除
  if (action === 'deleteCustomer') {
    var nm = e.parameter.name;
    if (!nm) return json({ error: 'name required' });
    var sh = getSheet('顧客');
    var d  = sh.getDataRange().getValues();
    for (var i = 1; i < d.length; i++) {
      if (d[i][0] === nm) { sh.deleteRow(i+1); break; }
    }
    return json({ ok: true });
  }

  if (action === 'deleteUid') {
    const uid = e.parameter.uid;
    if (!uid) return json({ error: 'uid required' });
    const ps = getSheet('未登録UID');
    const pd = ps.getDataRange().getValues();
    for (let j = 1; j < pd.length; j++) {
      if (pd[j][1] === uid) { ps.deleteRow(j+1); break; }
    }
    return json({ ok: true });
  }

  return json({ error: 'unknown action' });
}

// ================================================
// doOptions：CORSプリフライト対応
// ================================================
function doOptions(e) {
  return ContentService.createTextOutput(JSON.stringify({ok:true})).setMimeType(ContentService.MimeType.JSON);
}

// ================================================
// doPost：POSTリクエスト処理
// ================================================
function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch(err) { return json({ error:'invalid JSON' }); }

  if (body.action === 'saveOrder')       { saveOrderToSheet(body);    return json({ok:true}); }
  if (body.action === 'updatePayStatus') { updatePayInSheet(body.orderId, body.payStatus); return json({ok:true}); }
  if (body.action === 'updateStock')     { updateStockInSheet(body.branch, body.product, body.qty); return json({ok:true}); }
  if (body.action === 'saveCustomer')    { saveCustomerToSheet(body); return json({ok:true}); }

  // LINE Webhook
  (body.events||[]).forEach(ev => {
    if (ev.type !== 'message' || ev.message.type !== 'text') return;
    const text        = ev.message.text;
    const uid         = ev.source.userId;
    const displayName = ev.source.displayName || '';

    // ★ 未登録UIDチェック：顧客シートにUIDが存在するか確認
    const custRows = getSheet('顧客').getDataRange().getValues();
    const registered = custRows.slice(1).some(r => String(r[9]) === uid);

    if (!registered) {
      // 未登録リストに記録（重複は追加しない）
      const pendingSheet = getSheet('未登録UID');
      const pendingData  = pendingSheet.getDataRange().getValues();
      const alreadyPending = pendingData.some(r => String(r[1]) === uid);
      if (!alreadyPending) {
        if (pendingSheet.getLastRow() === 0) {
          pendingSheet.appendRow(['受信日時','LINE UID','表示名','メッセージ','状態']);
        }
        pendingSheet.appendRow([new Date().toLocaleString('ja-JP'), uid, displayName, text, '未設定']);
        if (ADMIN_USER_ID && ADMIN_USER_ID !== 'YOUR_LINE_USER_ID') {
          push(ADMIN_USER_ID, '🔔 未登録ユーザーからメッセージ\n名前：' + displayName + '\nUID：' + uid + '\n管理画面で顧客に紐付けてください');
        }
      }
      reply(ev.replyToken, 'メッセージありがとうございます。\n担当者より折り返しご連絡いたします。');
      return;
    }

    if (text.startsWith('【注文】')) {
      const o = parseOrder(text, uid);
      saveOrderToSheet(o); notifyAdmin(o);
      reply(ev.replyToken, '✅ 注文受付しました！\n' + o.items + '\n' + o.store);
    } else if (text === '在庫確認') {
      // スタッフのみ利用可能
      if (isStaff(uid)) {
        reply(ev.replyToken, stockMsg());
      } else {
        reply(ev.replyToken, '⚠️ 在庫確認はスタッフ専用です。\nご注文内容はマイページよりご確認ください。');
      }
    } else if (text === '注文確認') { reply(ev.replyToken, myOrders(uid)); }
      else if (text === '支払確認') { reply(ev.replyToken, myPayments(uid)); }
      else if (text === 'ヘルプ' || text === 'help') { reply(ev.replyToken, helpMsg(isStaff(uid))); }
  });
  return json({ ok: true });
}

function parseOrder(text, uid) {
  const lines = text.split('\n');
  const o = { uid, timestamp: new Date().toLocaleString('ja-JP'), store:'', items:'', payMethod:'売掛', note:'', status:'新規' };
  lines.forEach(l => {
    if (l.includes('店舗名：')) o.store     = l.split('：')[1]?.trim();
    if (l.includes('商品：'))  o.items     = l.split('：')[1]?.trim();
    if (l.includes('支払：'))  o.payMethod = l.split('：')[1]?.trim();
    if (l.includes('備考：'))  o.note      = l.split('：')[1]?.trim();
  });
  return o;
}

function getSheet(name) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

function saveOrderToSheet(o) {
  getSheet('注文受付').appendRow([
    o.timestamp || new Date().toLocaleString('ja-JP'),
    o.branch    || '',
    o.staff     || '',
    o.store     || '',
    o.area      || '',
    typeof o.items === 'object' ? JSON.stringify(o.items) : (o.items || ''),
    o.amount    || '',
    o.payMethod || '',
    o.payStatus || '未収',
    o.status    || '新規',
    o.note      || '',
    o.deliveryName    || '',  // 配達先名称（空＝店舗住所と同じ）
    o.deliveryAddress || '',  // 配達先住所
    o.uid       || ''
  ]);
}

function saveCustomerToSheet(c) {
  const sh = getSheet('顧客');
  const d  = sh.getDataRange().getValues();

  // 電話番号の先頭シングルクォート除去
  const tel      = String(c.tel      || '').replace(/^'+/, '');
  const ownerTel = String(c.ownerTel || '').replace(/^'+/, '');

  // 担当者リスト（contacts）をJSON文字列化
  let contacts = '';
  if (c.contacts && Array.isArray(c.contacts)) {
    contacts = JSON.stringify(c.contacts);
  } else if (typeof c.contacts === 'string') {
    contacts = c.contacts;
  }
  // ヘッダー（70列固定）= 基本25列 + 配達先45列（5件×9項目）
  // ※BS列以降は詳細情報+備考+商品別掛率（動的追加）
  const HEADER = [
    '拠点', '弊社担当', '店舗名', 'フリガナ', '店舗電話',
    '代表者名', '代表者電話', '担当者リスト',
    '郵便番号', '都道府県', '市町村郡', '番地', '建物名',
    'LINE ID', 'LINE UID',
    '紹介者', '請求書宛名',
    '支払方法',
    '振込名義①', '振込名義②', '振込名義③',
    '協賛', '購入条件', '協賛本数', '対象商品',
    // 配達先1〜5（各9項目 = 45列）
    '配達先1_名称', '配達先1_郵便番号', '配達先1_都道府県', '配達先1_市町村郡', '配達先1_番地', '配達先1_建物名', '配達先1_担当者', '配達先1_電話', '配達先1_メモ',
    '配達先2_名称', '配達先2_郵便番号', '配達先2_都道府県', '配達先2_市町村郡', '配達先2_番地', '配達先2_建物名', '配達先2_担当者', '配達先2_電話', '配達先2_メモ',
    '配達先3_名称', '配達先3_郵便番号', '配達先3_都道府県', '配達先3_市町村郡', '配達先3_番地', '配達先3_建物名', '配達先3_担当者', '配達先3_電話', '配達先3_メモ',
    '配達先4_名称', '配達先4_郵便番号', '配達先4_都道府県', '配達先4_市町村郡', '配達先4_番地', '配達先4_建物名', '配達先4_担当者', '配達先4_電話', '配達先4_メモ',
    '配達先5_名称', '配達先5_郵便番号', '配達先5_都道府県', '配達先5_市町村郡', '配達先5_番地', '配達先5_建物名', '配達先5_担当者', '配達先5_電話', '配達先5_メモ'
    // ↑ 25 + 45 = 70列。BS列以降は詳細情報・備考・商品別掛率（動的追加）
  ];

  const bankN = Array.isArray(c.bankNames) ? c.bankNames : (c.bankName ? [c.bankName] : []);
  const baseRow = [
    c.branch   || '',   // A: 拠点
    c.staff    || '',   // B: 弊社担当
    c.name     || '',   // C: 店舗名
    c.kana     || '',   // D: フリガナ
    tel,                // E: 店舗電話
    c.owner    || '',   // F: 代表者名
    ownerTel,           // G: 代表者電話
    contacts,           // H: 担当者リスト
    c.zip      || '',   // I: 郵便番号
    c.pref     || '',   // J: 都道府県
    c.city     || '',   // K: 市町村郡
    c.addr1    || '',   // L: 番地
    c.addr2    || '',   // M: 建物名
    c.lineId   || '',   // N: LINE ID
    c.lineUid  || '',   // O: LINE UID
    c.referrer || '',   // P: 紹介者
    c.invName  || '',   // Q: 請求書宛名
    c.payMethod || '',  // R: 支払方法
    bankN[0]   || '',   // S: 振込名義①
    bankN[1]   || '',   // T: 振込名義②
    bankN[2]   || '',   // U: 振込名義③
    c.sponsorEnabled ? 'あり' : 'なし',  // V: 協賛
    c.sponsorTrigger || '',              // W: 購入条件
    c.sponsorReward  || '',              // X: 協賛本数
    c.sponsorProduct || ''               // Y: 対象商品
  ];
  // 配達先5件分（各9列）を追加 — Z列以降
  const deliveries = Array.isArray(c.deliveries) ? c.deliveries : [];
  for (let n = 0; n < 5; n++) {
    const d = deliveries[n] || {};
    baseRow.push(
      d.name    || '',  // 配達先N_名称
      d.zip     || '',  // 配達先N_郵便番号
      d.pref    || '',  // 配達先N_都道府県
      d.city    || '',  // 配達先N_市町村郡
      d.addr1   || '',  // 配達先N_番地
      d.addr2   || '',  // 配達先N_建物名
      d.contact || '',  // 配達先N_担当者
      d.tel     || '',  // 配達先N_電話
      d.memo    || ''   // 配達先N_メモ
    );
  }
  // baseRow は 25 + 45 = 70要素
  // BS列以降は詳細情報 → 商品別掛率 → 備考（動的追加）

  // 商品別掛率マップ
  const ratesMap = {};
  if (c.productRates && Array.isArray(c.productRates)) {
    c.productRates.forEach(function(pr) { if (pr.key) ratesMap[pr.key] = pr.rate || ''; });
  }

  // 詳細情報（分析用）— 動的列として扱う（支払方法は固定列に移動済み）
  const detailFields = {
    '業態':         c.bizType      || '',
    '開店年月':     c.openDate     || '',
    '席数':         c.seats        || '',
    '経営者年代':   c.ownerAge     || '',
    '主要客層':     c.customerBase || '',
    '主力商品傾向': c.productTrend || '',
    '競合関係':     c.competitor   || '',
    '取引開始日':   c.dealStart    || '',
    '取引区分':     c.dealChannel  || '',
    '営業日時':     c.bizHours     || '',
    '平均客単価':   c.avgSpend     || '',
    '価格モード':   c.priceMode    || '新値',
    '税':           c.taxType      || '税込'
  };

  // 動的列マップ（詳細情報 → 商品別掛率 → 備考の順で最後に並ぶ）
  // ※ Object.assignはキー追加順を保持。備考を最後に置くと一番右の列になる
  const dynamicFields = Object.assign({}, detailFields, ratesMap, { '備考': c.note || '' });

  // ヘッダー行を取得（リセット後・空シートも正しく判定）
  // ★ A1セルが空 = リセット直後 と判定し、基本HEADERを書き込む
  let headerRow;
  const a1Value = (d.length > 0 && d[0].length > 0) ? String(d[0][0]).trim() : '';
  const isResetOrEmpty = d.length === 0 || !a1Value;
  if (isResetOrEmpty) {
    // リセット直後 or 完全に空のシート → 全クリアして基本HEADERを書き込む
    sh.clearContents();
    sh.appendRow(HEADER);
    headerRow = HEADER.slice();
  } else {
    // 既存ヘッダーを使用
    headerRow = d[0].map(String);
  }

  // 動的列があればヘッダーに追加（detailFields → ratesMap → 備考の順）
  Object.keys(dynamicFields).forEach(function(key) {
    if (headerRow.indexOf(key) < 0) {
      headerRow.push(key);
      sh.getRange(1, headerRow.length).setValue(key);
    }
  });

  // 最終行 = 基本70列（25 + 配達先45） + 動的列（詳細情報 + 商品別掛率 + 備考）
  const row = baseRow.slice();
  for (let col = 70; col < headerRow.length; col++) {
    row.push(dynamicFields[headerRow[col]] || '');
  }

  // 既存行を更新（店舗名=C列=index2 で照合）
  // ★ isResetOrEmpty なら既存行はないので走査スキップ
  if (!isResetOrEmpty) {
    for (let i = 1; i < d.length; i++) {
      if (d[i][2] === c.name) {
        sh.getRange(i+1, 1, 1, row.length).setValues([row]);
        sh.getRange(i+1, 5).setNumberFormat('@'); // 店舗電話(E列)：文字列
        sh.getRange(i+1, 7).setNumberFormat('@'); // 代表者電話(G列)：文字列
        return;
      }
    }
  }

  // 新規追加
  const newRow = sh.getLastRow() + 1;
  sh.appendRow(row);
  sh.getRange(newRow, 5).setNumberFormat('@'); // 店舗電話(E列)：文字列
  sh.getRange(newRow, 7).setNumberFormat('@'); // 代表者電話(G列)：文字列
}

function updatePayInSheet(orderId, status) {
  const sh = getSheet('注文受付');
  const d  = sh.getDataRange().getValues();
  for (let i = 1; i < d.length; i++) {
    if (String(d[i][11]) === String(orderId)) { sh.getRange(i+1, 9).setValue(status); break; }
  }
}

function updateStockInSheet(branch, product, qty) {
  const sh = getSheet('在庫');
  const d  = sh.getDataRange().getValues();
  for (let i = 1; i < d.length; i++) {
    if (d[i][0] === branch && d[i][1] === product) { sh.getRange(i+1, 3).setValue(qty); return; }
  }
  sh.appendRow([branch, product, qty]);
}

function updateAllStockInSheet(stockData) {
  const sh = getSheet('在庫');
  sh.clearContents();
  const rows = stockData.map(s => [s.branch, s.product, s.qty]);
  if (rows.length > 0) {
    sh.getRange(1, 1, rows.length, 3).setValues(rows);
  }
}

function stockMsg() {
  const rows = getSheet('在庫').getDataRange().getValues();
  return '📦 在庫一覧\n' + rows.slice(1).map(r => r[0] + ' ' + r[1] + '：' + r[2] + '本').join('\n');
}

function myOrders(uid) {
  const rows = getSheet('注文受付').getDataRange().getValues();
  const mine = rows.filter(r => r[11] === uid).slice(-3);
  return mine.length
    ? '📋 直近の注文\n' + mine.map(r => r[0] + ' ' + r[3] + ' [' + r[8] + ']').join('\n')
    : '注文履歴がありません';
}

function myPayments(uid) {
  const rows = getSheet('注文受付').getDataRange().getValues();
  const mine = rows.filter(r => r[11] === uid);
  if (!mine.length) return '💴 ご注文の履歴がありません';

  // 未収・売掛を抽出
  const unpaid = mine.filter(r => r[8] === '未収' || r[8] === '売掛');
  // 集金済みを直近3件
  const paid   = mine.filter(r => r[8] === '集金済み').slice(-3);

  let msg = '💴 お支払い状況\n';
  msg += '─────────────\n';

  if (unpaid.length) {
    // 未収合計（金額カラムが存在する場合）
    const totalUnpaid = unpaid.reduce((s, r) => s + (Number(r[6]) || 0), 0);
    msg += '【未収・売掛】' + unpaid.length + '件\n';
    unpaid.forEach(r => {
      const dateStr = String(r[0]).substring(0, 10);
      const amt     = r[6] ? '¥' + Number(r[6]).toLocaleString('ja-JP') : '金額未登録';
      msg += '  ' + dateStr + ' ' + amt + '［' + r[8] + '］\n';
    });
    if (totalUnpaid > 0) {
      msg += '  合計：¥' + totalUnpaid.toLocaleString('ja-JP') + '\n';
    }
  } else {
    msg += '✅ 未収・売掛なし\n';
  }

  if (paid.length) {
    msg += '─────────────\n';
    msg += '【直近の入金済み】\n';
    paid.forEach(r => {
      const dateStr = String(r[0]).substring(0, 10);
      const amt     = r[6] ? '¥' + Number(r[6]).toLocaleString('ja-JP') : '';
      msg += '  ' + dateStr + ' ' + amt + ' ✓\n';
    });
  }

  msg += '─────────────\n';
  msg += 'ご不明点は担当者までお問い合わせください。';
  return msg;
}

function helpMsg(staff) {
  let msg = '📖 使えるコマンド一覧\n─────────────\n【注文】\n  店舗名：○○\n  商品：イエロー×6\n  支払：売掛\n  備考：なし\n─────────────\n支払確認\n　未収・売掛の金額を確認\n注文確認\n　直近3件の注文を確認\nヘルプ\n　このメニューを表示';
  if (staff) {
    msg += '\n─────────────\n【スタッフ専用】\n在庫確認\n　現在の在庫本数を確認';
  }
  msg += '\n─────────────\nご不明点は担当者まで';
  return msg;
}

function notifyAdmin(o) {
  // 通知メッセージ作成（より詳細に）
  const itemsStr = o.items || '商品情報なし';
  const msg = '📩 新規注文\n' +
              '🏪 ' + (o.store || '無名') + '\n' +
              '📦 ' + itemsStr + '\n' +
              '💴 ' + (o.payMethod || '未設定') + '\n' +
              '📍 ' + (o.branch || '未設定') + '\n' +
              '⏰ ' + (o.timestamp || new Date().toLocaleString('ja-JP'));

  // STAFF_UIDS全員にPush通知
  STAFF_UIDS.forEach(uid => {
    if (uid && !uid.startsWith('YOUR_')) {
      try {
        push(uid, msg);
      } catch(e) {
        console.error('Push failed for ' + uid + ': ' + e.message);
      }
    }
  });
}

// ─── 新規注文取得API（管理画面のポーリング用） ──────────
function getNewOrders(since) {
  const rows = getSheet('注文受付').getDataRange().getValues();
  if (rows.length <= 1) return [];

  const newOrders = [];
  // ヘッダー後の行を逆順で確認（新しいものから）
  for (let i = rows.length - 1; i >= 1; i--) {
    const r = rows[i];
    const orderId = String(r[0] || '') + '_' + String(r[3] || '');
    if (since && orderId === since) break;
    newOrders.unshift({
      id: orderId,
      timestamp: r[0],
      branch: r[1],
      staff: r[2],
      store: r[3],
      area: r[4],
      items: r[5],
      amount: r[6],
      payMethod: r[7],
      payStatus: r[8],
      status: r[9],
    });
    if (newOrders.length >= 20) break; // 最大20件
  }
  return newOrders;
}

function reply(token, text) {
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
    method: 'post',
    headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + LINE_TOKEN },
    payload: JSON.stringify({ replyToken: token, messages: [{ type: 'text', text }] })
  });
}

function push(to, text) {
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method: 'post',
    headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + LINE_TOKEN },
    payload: JSON.stringify({ to, messages: [{ type: 'text', text }] })
  });
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}