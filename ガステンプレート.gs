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

// ══════════ OneSignal Webプッシュ（アプリアイコンの赤バッジ通知） ══════════
// OneSignalダッシュボード「Settings > Keys & IDs」の値を貼り付けてください
const ONESIGNAL_APP_ID   = 'YOUR_ONESIGNAL_APP_ID';
const ONESIGNAL_REST_KEY = 'YOUR_ONESIGNAL_REST_API_KEY';
function pushToDevices(title, message, url, badge){
  if (!ONESIGNAL_APP_ID || ONESIGNAL_APP_ID.indexOf('YOUR_') === 0) return; // 未設定なら何もしない
  try {
    UrlFetchApp.fetch('https://onesignal.com/api/v1/notifications', {
      method: 'post',
      contentType: 'application/json',
      headers: { 'Authorization': 'Key ' + ONESIGNAL_REST_KEY },
      payload: JSON.stringify({
        app_id: ONESIGNAL_APP_ID,
        included_segments: ['Total Subscriptions'],
        headings: { en: title, ja: title },
        contents: { en: message, ja: message },
        url: url || '',
        data: { badge: (typeof badge === 'number' ? badge : 1) }
      }),
      muteHttpExceptions: true
    });
  } catch(e) {}
}

// ══════════ コピー対策（4点セット） ══════════
// 秘密キー: 空文字 = 無効（既存動作のまま）。クライアントの PROTECT.SECRET_KEY と同じ値を入れると書込にキー必須化
const SECRET_KEY = '';
// 本番で稼働を許可するドメイン（ビーコンがこれ以外で来たら管理者へ通知）
const PROTECT_ALLOWED_HOSTS = ['minacel2026.github.io'];
function _keyOK(k){ return !SECRET_KEY || String(k) === String(SECRET_KEY); }

// ビーコン記録（コピー検知）。未許可ドメインは管理者へLINE通知（同ホストは1回だけ）
function logBeacon(p){
  try{
    const sh = getSheet('アクセスログ');
    if (sh.getLastRow() === 0) sh.getRange(1,1,1,8).setValues([['日時','種別','ホスト','URL','WM','参照元','UA','通知済']]);
    const host = String((p && p.host) || '');
    const unknown = host !== '' && PROTECT_ALLOWED_HOSTS.indexOf(host) < 0 && host !== 'localhost' && host !== '127.0.0.1';
    let alerted = '';
    if (unknown){
      const d = sh.getDataRange().getValues(); let already = false;
      for (let i=1;i<d.length;i++){ if (String(d[i][2])===host && String(d[i][7])==='1'){ already=true; break; } }
      if (!already && ADMIN_USER_ID && ADMIN_USER_ID.indexOf('YOUR_') < 0){
        try{ push(ADMIN_USER_ID, '🚨 未許可ドメインで稼働を検知\nhost: '+host+'\nURL: '+String((p&&p.href)||'')+'\nWM: '+String((p&&p.wm)||'')); alerted='1'; }catch(e){}
      }
    }
    sh.appendRow([new Date().toLocaleString('ja-JP'), String((p&&p.kind)||''), host, String((p&&p.href)||''), String((p&&p.wm)||''), String((p&&p.ref)||''), String((p&&p.ua)||''), alerted]);
  }catch(e){}
  return json({ ok:true });
}

function doGet(e) {
  // ── 診断: シート構成と注文受付の中身を確認（action=diag） ──
  if (e && e.parameter && e.parameter.action === 'diag') {
    try {
      const ss = SpreadsheetApp.openById(SHEET_ID);
      const sheets = ss.getSheets().map(function(sh){ return { name: sh.getName(), rows: sh.getLastRow(), cols: sh.getLastColumn() }; });
      const target = ss.getSheetByName('注文受付');
      let header = null, sample = null;
      if (target) {
        const v = target.getDataRange().getValues();
        header = v[0] || [];
        sample = v.length > 1 ? v.slice(1).slice(-3) : [];
      }
      return ContentService.createTextOutput(JSON.stringify({
        ok:true, spreadsheetName: ss.getName(), sheetId: SHEET_ID,
        sheets: sheets, orderSheetFound: !!target, orderRows: target ? target.getLastRow() : 0,
        orderHeader: header, orderLastRows: sample
      }, null, 2)).setMimeType(ContentService.MimeType.JSON);
    } catch(err) {
      return ContentService.createTextOutput(JSON.stringify({ ok:false, error:String(err) })).setMimeType(ContentService.MimeType.JSON);
    }
  }
  const action = e.parameter.action;

  if (action === 'ping') return json({ status: 'ok', version: 'stock-sync-v2', updateAllStock: true });

  if (action === 'beacon') return logBeacon(e.parameter);

  // ─── 新規注文取得（管理画面ポーリング用） ──────
  if (action === 'getNewOrders') {
    const since = e.parameter.since || '';
    return json({ orders: getNewOrders(since) });
  }

  // ===== 注文フォーム: 本人(UID)の協賛残数（商品別）を返す =====
  if (action === 'getMySponsor') { return json(_computeMySponsor(e.parameter.uid)); }

  // ===== 注文フォーム: 認証+協賛残数+在庫 を1往復でまとめて返す（高速化） =====
  if (action === 'getOrderInit') {
    const uid = e.parameter.uid;
    if (!uid) return json({ ok:false, error:'uid missing' });
    const auth = _computeAuth(uid);
    const sponsor = _computeMySponsor(uid);   // 協賛非対象は ok:false / total:0
    const stock = _computeStock();
    return json({ ok:true, auth: auth, sponsor: sponsor, stock: stock, shipping: _computeShippingConfig() });
  }

  // ===== マイページ: 本人(LINE UID)の注文だけを返す =====
  if (action === 'getMyOrders') {
    const uid = e.parameter.uid;
    if (!uid) return json({ ok:false, error:'uid missing' });
    // 顧客シートで UID → 店舗名 を特定
    const crows = getSheet('顧客').getDataRange().getValues();
    if (crows.length === 0) return json({ ok:false, error:'no customers' });
    const chd = crows[0].map(String);
    const uidCol = chd.indexOf('LINE UID');
    const nameCol = chd.indexOf('店舗名');
    if (uidCol < 0 || nameCol < 0) return json({ ok:false, error:'headers' });
    let store = '';
    for (let i = 1; i < crows.length; i++) {
      if (String(crows[i][uidCol]).trim() === String(uid).trim()) { store = String(crows[i][nameCol]).trim(); break; }
    }
    if (!store) return json({ ok:false, error:'unregistered' });
    // 注文受付から 店舗一致 or UID一致 を抽出（新しい順）
    const rows = getSheet('注文受付').getDataRange().getValues();
    const orders = rows.slice(1).filter(r =>
      String(r[3]).trim() === store || (r[13] && String(r[13]).trim() === String(uid).trim())
    ).map(r => ({
      date: r[0], items: r[5], amount: r[6], payMethod: r[7],
      payStatus: r[8], status: r[9], id: r[14]
    })).reverse();
    let unpaid = 0;
    orders.forEach(o => { if (String(o.payStatus).indexOf('未') === 0) unpaid += Number(o.amount) || 0; });
    // 協賛: 残数（商品別）＋ 履歴（付与/使用）を同梱
    const _sp = _computeMySponsor(uid);   // { ok, custId, total, byProduct }
    let _spLogs = [];
    if (_sp && _sp.ok && _sp.custId) {
      const _lsh = getSheet('協賛ログ');
      if (_lsh.getLastRow() > 1) {
        const _lr = _lsh.getDataRange().getValues();
        for (let _i = 1; _i < _lr.length; _i++) {
          const _r = _lr[_i];
          if (!_r[0]) continue;
          if (String(_r[1]).trim() !== String(_sp.custId)) continue;  // 顧客ID一致
          _spLogs.push({ date: String(_r[5] || ''), type: String(_r[3] || ''), qty: (typeof _r[4] === 'number') ? _r[4] : (parseInt(_r[4]) || 0), product: String(_r[9] || ''), orderId: String(_r[6] || '') });
        }
        _spLogs.reverse();              // 新しい順
        _spLogs = _spLogs.slice(0, 30); // 最新30件
      }
    }
    return json({ ok:true, store: store, orders: orders, unpaidTotal: unpaid,
      sponsor: (_sp && _sp.ok) ? { total: _sp.total, byProduct: _sp.byProduct } : { total: 0, byProduct: {} },
      sponsorLogs: _spLogs });
  }

  if (action === 'getOrders') {
    const rows = getSheet('注文受付').getDataRange().getValues();
    return json({ orders: rows.slice(1).map(r => ({
      date:r[0], branch:r[1], staff:r[2], store:r[3], area:r[4],
      items:r[5], amount:r[6], payMethod:r[7], payStatus:r[8],
      status:r[9], note:r[10], uid:r[13], id:r[14]
    }))});
  }

  if (action === 'getStock') { return json({ stock: _computeStock() }); }
  if (action === 'getStaff')   return getStaff();
  if (action === 'getCompany') return getCompany();
  if (action === 'getProducts') return getProducts();

  if (action === 'checkLineUser') { return json(_computeAuth(e.parameter.uid)); }

  if (action === 'getCustomers') {
    const rows = getSheet('顧客').getDataRange().getValues();
    if (rows.length === 0) return json({ customers: [] });
    const hdrs = rows[0].map(String);
    // ヘッダー名 → 列インデックス（列の物理位置に依存せず読むための要）
    const colIdx = {};
    hdrs.forEach(function(h, i){ if (colIdx[h] === undefined) colIdx[h] = i; });
    // 協賛 名称⇔ID 変換マップ
    const spoMaps = getSponsorMaps();
    // 固定ヘッダー集合（これ以外を動的列＝詳細情報/商品別掛率/顧客ID/備考 として処理）
    // ※旧「購入条件/協賛本数/対象商品」も既知扱いにして無視する（旧データ互換）
    const FIXED = {};
    ['拠点','弊社担当','店舗名','フリガナ','店舗電話','代表者名','代表者電話','担当者リスト',
     '郵便番号','都道府県','市町村郡','番地','建物名','LINE ID','LINE UID','紹介者','請求書宛名',
     '支払方法','振込名義①','振込名義②','振込名義③','協賛','適用協賛',
     '購入条件','協賛本数','対象商品','適用協賛ID'
    ].forEach(function(h){ FIXED[h] = true; });
    for (let n = 1; n <= 5; n++) {
      ['名称','郵便番号','都道府県','市町村郡','番地','建物名','担当者','電話','メモ'].forEach(function(s){
        FIXED['配達先'+n+'_'+s] = true;
      });
    }
    // 詳細情報ヘッダー → JSプロパティ名
    const DETAIL_KEYS = {
      '業態':'bizType','開店年月':'openDate','席数':'seats','経営者年代':'ownerAge',
      '主要客層':'customerBase','主力商品傾向':'productTrend','競合関係':'competitor',
      '取引開始日':'dealStart','取引区分':'dealChannel','営業日時':'bizHours','平均客単価':'avgSpend',
      '共通価格モード':'commonPriceMode','共通税':'commonTaxType','掛率':'_rateRaw','顧客ID':'_custId',
      '協賛即時反映':'_spInstant','送料無料':'_freeShip','個別送料':'_shipFee','協賛を送料本数に含める':'_shipIncSp'
    };
    return json({ customers: rows.slice(1).map(function(r){
      const val = function(name){
        const i = colIdx[name];
        return (i !== undefined && r[i] !== undefined && r[i] !== null) ? String(r[i]) : '';
      };
      // 配達先5件（ヘッダー名で読む＝Z列でもX列でも正しく取得）
      const deliveries = [];
      for (let n = 1; n <= 5; n++) {
        const d = {
          name:    val('配達先'+n+'_名称'),
          zip:     val('配達先'+n+'_郵便番号'),
          pref:    val('配達先'+n+'_都道府県'),
          city:    val('配達先'+n+'_市町村郡'),
          addr1:   val('配達先'+n+'_番地'),
          addr2:   val('配達先'+n+'_建物名'),
          contact: val('配達先'+n+'_担当者'),
          tel:     val('配達先'+n+'_電話'),
          memo:    val('配達先'+n+'_メモ')
        };
        if (d.name || d.zip || d.addr1 || d.contact) deliveries.push(d);
      }
      // 動的列（固定ヘッダー以外）: 詳細情報 / 商品別掛率 / 備考 / 顧客ID
      const details = {}; let note = ''; const prodMap = {};
      hdrs.forEach(function(h, col){
        if (!h || FIXED[h]) return;
        const cell = (r[col] !== undefined && r[col] !== null) ? String(r[col]) : '';
        if (h === '備考') { note = cell; }
        else if (DETAIL_KEYS[h]) { details[DETAIL_KEYS[h]] = cell; }
        else if (cell !== '') {
          if (h.length > 4 && h.slice(-4) === '_モード') {
            const pn = h.slice(0, -4); if (!prodMap[pn]) prodMap[pn] = {}; prodMap[pn].priceMode = cell;
          } else if (h.length > 2 && h.slice(-2) === '_税') {
            const pn = h.slice(0, -2); if (!prodMap[pn]) prodMap[pn] = {}; prodMap[pn].taxType = cell;
          } else {
            if (!prodMap[h]) prodMap[h] = {}; prodMap[h].rate = cell;
          }
        }
      });
      const pr = [];
      Object.keys(prodMap).forEach(function(pn){
        const e = prodMap[pn];
        if (e.rate) pr.push({ key: pn, rate: e.rate, priceMode: e.priceMode || '新値', taxType: e.taxType || '税込' });
      });
      // 適用協賛（名称カンマ区切り）→ ID配列に逆変換
      const spoNames = val('適用協賛');
      const sponsorIds = spoNames
        ? spoNames.split(',').map(function(x){return x.trim();}).filter(Boolean)
            .map(function(nm){ return spoMaps.name2id[nm] || nm; }).filter(Boolean)
        : [];
      const result = Object.assign({
        branch: val('拠点'), staff: val('弊社担当'), name: val('店舗名'), kana: val('フリガナ'), tel: val('店舗電話'),
        owner: val('代表者名'), ownerTel: val('代表者電話'), contacts: val('担当者リスト'),
        zip: val('郵便番号'), pref: val('都道府県'), city: val('市町村郡'), addr1: val('番地'), addr2: val('建物名'),
        lineId: val('LINE ID'), lineUid: val('LINE UID'),
        referrer: val('紹介者'), invName: val('請求書宛名'),
        payMethod: val('支払方法'),
        bankNames: [val('振込名義①'),val('振込名義②'),val('振込名義③')].filter(Boolean),
        bankName: val('振込名義①'),
        note: note,
        deliveries: deliveries,
        productRates: pr,
        rate: pr.length > 0 ? (pr[0].rate || '') : '',
        sponsorEnabled: val('協賛') === 'あり',
        sponsorInstant: val('協賛即時反映') === 'あり',
        freeShipping: val('送料無料') === 'あり',
        shipFee: val('個別送料'),
        shipIncludeSponsor: val('協賛を送料本数に含める'),
        sponsorIds: sponsorIds
      }, details);
      // 共通掛率: 「掛率」列があればそれを優先（なければ商品別掛率の最初をフォールバック）
      if (result._rateRaw) result.rate = result._rateRaw;
      delete result._rateRaw;
      // 顧客ID: スプシ保存値 → なければ店舗名から安定ID生成
      if (result._custId) {
        result.id = result._custId;
      } else {
        var _nm = String(result.name || ''); var _h = 0;
        for (var _ci = 0; _ci < _nm.length; _ci++) { _h = ((_h << 5) - _h + _nm.charCodeAt(_ci)) | 0; }
        result.id = 'cust-sync-' + Math.abs(_h).toString(36);
      }
      delete result._custId;
      delete result._spInstant; delete result._freeShip; delete result._shipFee; delete result._shipIncSp;
      return result;
    })});
  }

  // ★ ブラウザno-cors対策：GETのpayloadパラメータで書き込みを受け取る
  if (e.parameter.payload) {
    try {
      const body = JSON.parse(e.parameter.payload);
      if (!_keyOK(body.key)) return json({ error: 'auth' });
      if (body.action === 'saveCustomer')    { saveCustomerToSheet(body); return json({ok:true}); }
      if (body.action === 'saveStaff')       { return saveStaff(body); }
      if (body.action === 'deleteStaff')     { return deleteStaff(body); }
      if (body.action === 'saveCompany')     { return saveCompany(body); }
      if (body.action === 'saveProducts')    { return saveProducts(body); }
      if (body.action === 'saveCustomersBulk') { return saveCustomersBulk(body); }
      if (body.action === 'deleteCustomer')  { deleteCustomerFromSheet(body.id, body.name); return json({ok:true}); }
      if (body.action === 'saveOrder')       { saveOrderToSheet(body);    return json({ok:true}); }
      if (body.action === 'saveOrdersBulk')  { return saveOrdersBulk(body); }
      if (body.action === 'updatePayStatus') { updatePayInSheet(body.orderId, body.payStatus, body.collector, body.collectedDate); return json({ok:true}); }
      if (body.action === 'updateOrderShip') { updateShipInSheet(body.orderId, body.tracking, body.status); return json({ok:true}); }
      if (body.action === 'updateStock')     { updateStockInSheet(body.branch, body.product, body.qty); return json({ok:true}); }
      if (body.action === 'updateAllStock')  { updateAllStockInSheet(body.stockData); return json({ok:true}); }
      if (body.action === 'saveReferrer')    { saveReferrerToSheet(body); return json({ok:true}); }
      if (body.action === 'deleteReferrer')  { deleteReferrerFromSheet(body.id); return json({ok:true}); }
      if (body.action === 'saveSponsor')     { saveSponsorToSheet(body); return json({ok:true}); }
      if (body.action === 'deleteSponsor')   { deleteSponsorFromSheet(body.id); return json({ok:true}); }
      if (body.action === 'saveSponsorLog')  { saveSponsorLogToSheet(body); return json({ok:true}); }
      if (body.action === 'deleteOrder')     { deleteOrderFromSheet(body.id, body.timestamp, body.store, body.amount); return json({ok:true}); }
      if (body.action === 'deleteSponsorLog'){ deleteSponsorLogFromSheet(body.id); return json({ok:true}); }
    } catch(err) {
      return json({ error: 'parse error: ' + err.message });
    }
  }

  if (action === 'getSponsors') {
    const sh = getSheet('協賛');
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, 12).setValues([['ID','ルール名','対象商品','購入条件数量','購入単位','報酬数量','報酬商品','開始日','終了日','状態','メモ','更新日時']]);
    }
    const rows = sh.getDataRange().getValues();
    if (rows.length <= 1) return json({ sponsors: [] });
    return json({ sponsors: rows.slice(1).filter(r => r[0]).map(r => {
      // 日付は ISO 形式の "YYYY-MM-DD" に統一
      function toDateStr(v) {
        if (!v) return '';
        if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        return String(v);
      }
      return {
        id:            String(r[0] || ''),
        name:          String(r[1] || ''),
        targetProduct: String(r[2] || 'all'),
        triggerQty:    (typeof r[3] === 'number') ? r[3] : (parseFloat(r[3]) || 0),
        triggerUnit:   String(r[4] || 'ケース'),
        rewardQty:     (typeof r[5] === 'number') ? r[5] : (parseFloat(r[5]) || 0),
        rewardProduct: String(r[6] || 'same'),
        startDate:     toDateStr(r[7]),
        endDate:       toDateStr(r[8]),
        enabled:       String(r[9] || '有効'),
        memo:          String(r[10] || ''),
        updatedAt:     r[11] ? String(r[11]) : ''
      };
    })});
  }

  if (action === 'getSponsorLogs') {
    const sh = getSheet('協賛ログ');
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, 14).setValues([['ID','顧客ID','顧客名','種別','本数','日付','注文ID','協賛ID','協賛名','商品','拠点','メモ','作成日時','作成者']]);
    }
    const rows = sh.getDataRange().getValues();
    if (rows.length <= 1) return json({ logs: [] });
    return json({ logs: rows.slice(1).filter(r => r[0]).map(function(r){
      return {
        id:         String(r[0] || ''),
        custId:     String(r[1] || ''),
        custName:   String(r[2] || ''),
        type:       String(r[3] || ''),
        qty:        (typeof r[4] === 'number') ? r[4] : (parseInt(r[4]) || 0),
        date:       (r[5] instanceof Date) ? Utilities.formatDate(r[5], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(r[5] || ''),
        orderId:    String(r[6] || ''),
        sponsorId:  String(r[7] || ''),
        sponsorName:String(r[8] || ''),
        product:    String(r[9] || ''),
        branch:     String(r[10] || ''),
        note:       String(r[11] || ''),
        createdAt:  r[12] ? String(r[12]) : '',
        createdBy:  String(r[13] || '')
      };
    })});
  }

  if (action === 'getReferrers') {
    const sh = getSheet('紹介者');
    // ヘッダーが空ならセットアップ
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, 9).setValues([['ID','名前','電話番号','住所','メール','LINE ID','計算方法','パラメータ','更新日時']]);
    }
    const rows = sh.getDataRange().getValues();
    if (rows.length <= 1) return json({ referrers: [] });
    return json({ referrers: rows.slice(1).filter(r => r[0]).map(r => ({
      id:        String(r[0] || ''),
      name:      String(r[1] || ''),
      tel:       String(r[2] || ''),
      address:   String(r[3] || ''),
      email:     String(r[4] || ''),
      lineId:    String(r[5] || ''),
      calcType:  String(r[6] || 'A'),
      calcParam: (typeof r[7] === 'number') ? r[7] : (parseFloat(r[7]) || 0),
      updatedAt: r[8] ? new Date(r[8]).toISOString() : ''
    }))});
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
    // ヘッダー名で「店舗名」「LINE UID」の列を特定（列位置に依存しない）
    const hdr = d[0].map(String);
    const storeCol = hdr.indexOf('店舗名') >= 0 ? hdr.indexOf('店舗名') : 2;   // C列
    const uidCol   = hdr.indexOf('LINE UID') >= 0 ? hdr.indexOf('LINE UID') : 14; // O列
    for (let i = 1; i < d.length; i++) {
      if (String(d[i][storeCol]).trim() === String(storeName).trim()) {
        sh.getRange(i+1, uidCol+1).setValue(uid); // LINE UID列（ヘッダー名で特定）
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
    deleteCustomerFromSheet(e.parameter.id, e.parameter.name);
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
  if (!body.events && !_keyOK(body.key)) return json({ error: 'auth' });

  if (body.action === 'saveOrder')       { saveOrderToSheet(body);    return json({ok:true}); }
  if (body.action === 'saveOrdersBulk')  { return saveOrdersBulk(body); }
  if (body.action === 'updatePayStatus') { updatePayInSheet(body.orderId, body.payStatus, body.collector, body.collectedDate); return json({ok:true}); }
  if (body.action === 'updateOrderShip') { updateShipInSheet(body.orderId, body.tracking, body.status); return json({ok:true}); }
  if (body.action === 'updateStock')     { updateStockInSheet(body.branch, body.product, body.qty); return json({ok:true}); }
  if (body.action === 'updateAllStock')  { updateAllStockInSheet(body.stockData); return json({ok:true}); }
  if (body.action === 'saveCustomer')    { saveCustomerToSheet(body); return json({ok:true}); }
  if (body.action === 'saveStaff')       { return saveStaff(body); }
  if (body.action === 'deleteStaff')     { return deleteStaff(body); }
  if (body.action === 'saveCompany')     { return saveCompany(body); }
  if (body.action === 'saveProducts')    { return saveProducts(body); }
  if (body.action === 'saveCustomersBulk') { return saveCustomersBulk(body); }
  if (body.action === 'deleteCustomer')  { deleteCustomerFromSheet(body.id, body.name); return json({ok:true}); }
  if (body.action === 'saveReferrer')    { saveReferrerToSheet(body); return json({ok:true}); }
  if (body.action === 'deleteReferrer')  { deleteReferrerFromSheet(body.id); return json({ok:true}); }
  if (body.action === 'saveSponsor')     { saveSponsorToSheet(body); return json({ok:true}); }
  if (body.action === 'deleteSponsor')   { deleteSponsorFromSheet(body.id); return json({ok:true}); }
  if (body.action === 'saveSponsorLog')  { saveSponsorLogToSheet(body); return json({ok:true}); }
  if (body.action === 'deleteOrder')     { deleteOrderFromSheet(body.id, body.timestamp, body.store, body.amount); return json({ok:true}); }
  if (body.action === 'deleteSponsorLog'){ deleteSponsorLogFromSheet(body.id); return json({ok:true}); }

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
      try { pushToDevices('🔔 未登録ユーザーから連絡', (displayName || 'お客様') + ' さんからメッセージが届きました', '', getSheet('未登録UID').getLastRow() - 1); } catch(e){}
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

// ══════════ 会社情報・スタッフ（別PC共有用） ══════════
var STAFF_COLS = ['id','name','email','branch','role','tel','lineUid','active','joined','note','updatedAt'];

function getStaff() {
  const sh = getSheet('スタッフ');
  const d = sh.getDataRange().getValues();
  if (d.length < 2) return json({ staff: [] });
  const h = d[0].map(String);
  const idx = {}; STAFF_COLS.forEach(c => { idx[c] = h.indexOf(c); });
  const out = [];
  for (let i = 1; i < d.length; i++) {
    const r = d[i];
    const o = {};
    STAFF_COLS.forEach(c => { o[c] = (idx[c] >= 0) ? r[idx[c]] : ''; });
    if (o.active === 'true' || o.active === true) o.active = true;
    else if (o.active === 'false' || o.active === false) o.active = false;
    if (o.id !== '' && o.id != null) out.push(o);
  }
  return json({ staff: out });
}

function saveStaff(payload) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch (e) { return json({ ok: false, error: 'lock-timeout' }); }
  try {
    const sh = getSheet('スタッフ');
    let d = sh.getDataRange().getValues();
    if (d.length === 0 || String(d[0][0]) !== 'id') {
      sh.clear();
      sh.getRange(1, 1, 1, STAFF_COLS.length).setValues([STAFF_COLS]);
      d = [STAFF_COLS.slice()];
    }
    const row = STAFF_COLS.map(c => { const v = payload[c]; return (v === undefined || v === null) ? '' : v; });
    let found = 0;
    for (let i = 1; i < d.length; i++) { if (String(d[i][0]) === String(payload.id)) { found = i + 1; break; } }
    const targetRow = found ? found : (sh.getLastRow() + 1);
    sh.getRange(targetRow, 1, 1, STAFF_COLS.length).setValues([row]);
    // 電話番号は先頭0が消えないよう文字列書式で書き直す（顧客電話と同じ対策）
    const telCol = STAFF_COLS.indexOf('tel') + 1;
    sh.getRange(targetRow, telCol).setNumberFormat('@');
    sh.getRange(targetRow, telCol).setValue(String(payload.tel == null ? '' : payload.tel).replace(/^'+/, ''));
    return json({ ok: true });
  } finally { lock.releaseLock(); }
}

function deleteStaff(payload) {
  const sh = getSheet('スタッフ');
  const d = sh.getDataRange().getValues();
  for (let i = d.length - 1; i >= 1; i--) {
    if (String(d[i][0]) === String(payload.id)) sh.deleteRow(i + 1);
  }
  return json({ ok: true });
}

function saveCompany(payload) {
  const sh = getSheet('会社');
  const obj = {};
  Object.keys(payload).forEach(k => { if (k !== 'action') obj[k] = payload[k]; });
  sh.getRange(1, 1).setValue('company_json');
  sh.getRange(2, 1).setValue(JSON.stringify(obj));
  return json({ ok: true });
}

function getCompany() {
  const sh = getSheet('会社');
  const d = sh.getDataRange().getValues();
  if (d.length < 2 || !d[1][0]) return json({ company: {} });
  try { return json({ company: JSON.parse(d[1][0]) }); } catch (e) { return json({ company: {} }); }
}

// ══════════ 商品マスター（全端末共有 / 注文フォーム・在庫の単一ソース） ══════════
var PRODUCT_COLS = ['name','fullName','category','cost','costOld','price','priceOld','reorder','color','img','active','sortOrder','updatedAt'];
var PRODUCT_NUM = { cost:1, costOld:1, price:1, priceOld:1, reorder:1, sortOrder:1 };

function getProducts() {
  const sh = getSheet('商品');
  const d = sh.getDataRange().getValues();
  if (d.length < 2) return json({ products: [] });
  const h = d[0].map(String);
  const idx = {}; PRODUCT_COLS.forEach(c => { idx[c] = h.indexOf(c); });
  const out = [];
  for (let i = 1; i < d.length; i++) {
    const r = d[i]; const o = {};
    PRODUCT_COLS.forEach(c => {
      let v = (idx[c] >= 0) ? r[idx[c]] : '';
      if (PRODUCT_NUM[c]) v = (typeof v === 'number') ? v : (parseFloat(v) || 0);
      o[c] = v;
    });
    o.key = o.name;
    o.active = !(o.active === 'false' || o.active === false);
    if (o.name !== '' && o.name != null) out.push(o);
  }
  out.sort(function(a, b){ return (a.sortOrder || 0) - (b.sortOrder || 0); });
  return json({ products: out });
}

function saveProducts(payload) {
  const list = (payload && payload.products) || [];
  const lock = LockService.getScriptLock();
  try { lock.waitLock(60000); } catch (e) { return json({ ok: false, error: 'lock-timeout' }); }
  try {
    const sh = getSheet('商品');
    sh.clear();
    sh.getRange(1, 1, 1, PRODUCT_COLS.length).setValues([PRODUCT_COLS]);
    if (list.length) {
      const rows = list.map(function(p, i){
        return PRODUCT_COLS.map(function(c){
          if (c === 'sortOrder') return (p.sortOrder != null ? p.sortOrder : i);
          if (c === 'updatedAt') return new Date().toISOString();
          if (c === 'active')    return (p.active === false ? 'false' : 'true');
          const v = p[c]; return (v === undefined || v === null) ? '' : v;
        });
      });
      sh.getRange(2, 1, rows.length, PRODUCT_COLS.length).setValues(rows);
    }
    return json({ ok: true, count: list.length });
  } finally { lock.releaseLock(); }
}

// ───────────────────────────────────────────────
// 紹介者シートへの書き込み・削除
// ───────────────────────────────────────────────
function saveReferrerToSheet(r) {
  const sh = getSheet('紹介者');
  // ヘッダーが空ならセットアップ
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, 9).setValues([['ID','名前','電話番号','住所','メール','LINE ID','計算方法','パラメータ','更新日時']]);
  }
  const rows = sh.getDataRange().getValues();
  const newRow = [
    r.id        || ('ref-' + Date.now()),
    r.name      || '',
    r.tel       || '',
    r.address   || '',
    r.email     || '',
    r.lineId    || '',
    r.calcType  || 'A',
    (typeof r.calcParam === 'number') ? r.calcParam : (parseFloat(r.calcParam) || 0),
    new Date().toLocaleString('ja-JP')
  ];
  // 既存IDがあれば更新、なければ追加
  let updated = false;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(newRow[0])) {
      sh.getRange(i + 1, 1, 1, 9).setValues([newRow]);
      updated = true;
      break;
    }
  }
  if (!updated) sh.appendRow(newRow);
}

function deleteReferrerFromSheet(id) {
  const sh = getSheet('紹介者');
  const rows = sh.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][0]) === String(id)) {
      sh.deleteRow(i + 1);
      return;
    }
  }
}

// ───────────────────────────────────────────────
// 協賛ログ（earned/used）の追記（残数の元データ）
// ※IDで重複チェック → 既存なら何もしない（複数端末からの二重追記防止）
// ───────────────────────────────────────────────
function saveSponsorLogToSheet(l) {
  const sh = getSheet('協賛ログ');
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, 14).setValues([['ID','顧客ID','顧客名','種別','本数','日付','注文ID','協賛ID','協賛名','商品','拠点','メモ','作成日時','作成者']]);
  }
  const id = String(l.id || ('sp-' + Date.now()));
  // 重複チェック（A列のID）
  const ids = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
  for (let i = 1; i < ids.length; i++) {
    if (String(ids[i][0]) === id) return; // 既に存在 → 二重追記しない
  }
  sh.appendRow([
    id,
    String(l.custId || ''),
    String(l.custName || ''),
    String(l.type || ''),
    (typeof l.qty === 'number') ? l.qty : (parseInt(l.qty) || 0),
    String(l.date || ''),
    String(l.orderId || ''),
    String(l.sponsorId || ''),
    String(l.sponsorName || ''),
    String(l.product || ''),
    String(l.branch || ''),
    String(l.note || ''),
    String(l.createdAt || new Date().toISOString()),
    String(l.createdBy || '')
  ]);
}

// ───────────────────────────────────────────────
// 協賛シートへの書き込み・削除
// ───────────────────────────────────────────────
function saveSponsorToSheet(s) {
  const sh = getSheet('協賛');
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, 12).setValues([['ID','ルール名','対象商品','購入条件数量','購入単位','報酬数量','報酬商品','開始日','終了日','状態','メモ','更新日時']]);
  }
  const rows = sh.getDataRange().getValues();
  const newRow = [
    s.id            || ('spo-' + Date.now()),
    s.name          || '',
    s.targetProduct || 'all',
    (typeof s.triggerQty === 'number') ? s.triggerQty : (parseFloat(s.triggerQty) || 0),
    s.triggerUnit   || 'ケース',
    (typeof s.rewardQty  === 'number') ? s.rewardQty  : (parseFloat(s.rewardQty)  || 0),
    s.rewardProduct || 'same',
    s.startDate     || '',
    s.endDate       || '',
    s.enabled       || '有効',
    s.memo          || '',
    new Date().toLocaleString('ja-JP')
  ];
  let updated = false;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(newRow[0])) {
      sh.getRange(i + 1, 1, 1, 12).setValues([newRow]);
      updated = true;
      break;
    }
  }
  if (!updated) sh.appendRow(newRow);
}

function deleteSponsorFromSheet(id) {
  const sh = getSheet('協賛');
  const rows = sh.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][0]) === String(id)) {
      sh.deleteRow(i + 1);
      return;
    }
  }
}

// ───────────────────────────────────────────────
// 手動実行用：協賛シートをセットアップ
// （Apps Script の関数選択で「initSponsorSheet」を選んで実行）
// ───────────────────────────────────────────────
function initSponsorSheet() {
  const sh = getSheet('協賛');
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, 12).setValues([['ID','ルール名','対象商品','購入条件数量','購入単位','報酬数量','報酬商品','開始日','終了日','状態','メモ','更新日時']]);
    Logger.log('✅ 協賛シートを作成しました（ヘッダー設定完了）');
  } else {
    Logger.log('協賛シートは既に存在します（行数: ' + sh.getLastRow() + '）');
  }
}

// ───────────────────────────────────────────────
// 手動実行用：協賛ログシートをセットアップ
// （Apps Script の関数選択で「initSponsorLogSheet」を選んで実行）
// ───────────────────────────────────────────────
function initSponsorLogSheet() {
  const sh = getSheet('協賛ログ');
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, 14).setValues([['ID','顧客ID','顧客名','種別','本数','日付','注文ID','協賛ID','協賛名','商品','拠点','メモ','作成日時','作成者']]);
    Logger.log('✅ 協賛ログシートを作成しました（ヘッダー設定完了）');
  } else {
    Logger.log('協賛ログシートは既に存在します（行数: ' + sh.getLastRow() + '）');
  }
}

// ───────────────────────────────────────────────
// 手動実行用：紹介者シートをセットアップ
// ───────────────────────────────────────────────
function initReferrerSheet() {
  const sh = getSheet('紹介者');
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, 9).setValues([['ID','名前','電話番号','住所','メール','LINE ID','計算方法','パラメータ','更新日時']]);
    Logger.log('✅ 紹介者シートを作成しました（ヘッダー設定完了）');
  } else {
    Logger.log('紹介者シートは既に存在します（行数: ' + sh.getLastRow() + '）');
  }
}

// ── 注文の一括取込（実績の過去注文などを一括追記）──
// 注文ID(O列=idx14)で重複チェック → 既存はスキップ、新規だけ追記（冪等）。ロック1回・まとめ書き。
function saveOrdersBulk(payload) {
  const list = (payload && payload.orders) || [];
  if (!list.length) return json({ ok: true, added: 0, skipped: 0, total: 0 });
  const lock = LockService.getScriptLock();
  try { lock.waitLock(300000); } catch (e) { return json({ ok: false, error: 'lock-timeout' }); }
  try {
    const sh = getSheet('注文受付');
    const d = sh.getDataRange().getValues();
    const W = Math.max(15, d.length ? d[0].length : 15);
    const have = {};
    for (let i = 1; i < d.length; i++) {
      const id = String(d[i][14] || '').trim();
      if (id) have[id] = true;
    }
    const newRows = []; let skipped = 0;
    list.forEach(function(o){
      const id = String(o.id || '').trim();
      if (id && have[id]) { skipped++; return; }   // 既存ID → 追記しない（冪等）
      if (id) have[id] = true;                       // 同一バッチ内の重複も防ぐ
      const row = [
        o.timestamp || o.date || '',
        o.branch || '', o.staff || '', o.store || '', o.area || '',
        (typeof o.items === 'object') ? JSON.stringify(o.items) : (o.items || ''),
        o.amount || '', o.payMethod || '',
        o.payStatus || '未収', o.status || '新規', o.note || '',
        o.deliveryName || '', o.deliveryAddress || '', o.uid || '', o.id || ''
      ];
      while (row.length < W) row.push('');
      newRows.push(row);
    });
    if (newRows.length) {
      const start = sh.getLastRow() + 1;
      sh.getRange(start, 1, newRows.length, W).setValues(newRows);
    }
    return json({ ok: true, added: newRows.length, skipped: skipped, total: list.length });
  } finally {
    lock.releaseLock();
  }
}

function saveOrderToSheet(o) {
  const sh = getSheet('注文受付');
  const rows = sh.getDataRange().getValues();
  const oid = String(o.id || '').trim();
  const newRow = [
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
    o.uid       || '',
    o.id        || ''         // O列: 注文ID（削除・更新時の照合に使用）
  ];
  // upsert: 同じ注文ID（O列）があればその行を上書き（編集の重複防止）。なければ追加。
  if (oid) {
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][14] || '').trim() === oid) {
        sh.getRange(i + 1, 1, 1, newRow.length).setValues([newRow]);
        return;
      }
    }
  } else {
    // ID無しの旧データ互換: 日時+店舗+金額が一致すれば重複とみなしスキップ
    const ots = String(o.timestamp || '').trim();
    const ostore = String(o.store || '').trim();
    const oamt = String(o.amount || '').trim();
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      if (ots && String(r[0] || '').trim() === ots
              && String(r[3] || '').trim() === ostore
              && String(r[6] || '').trim() === oamt) return;
    }
  }
  sh.appendRow(newRow);

  // 協賛注文（sponsorItems）があれば「used」ログを記録（無料出荷・残数消費）
  if (o.sponsorItems && typeof o.sponsorItems === 'object') {
    const _today = (o.timestamp ? String(o.timestamp).slice(0,10) : new Date().toISOString().slice(0,10));
    Object.keys(o.sponsorItems).forEach(function(prod){
      const q = parseInt(o.sponsorItems[prod]) || 0;
      if (q > 0) {
        saveSponsorLogToSheet({
          id: 'spu-' + String(o.id || Date.now()) + '-' + prod,
          custId: o.custId || '',
          custName: o.store || '',
          type: 'used',
          qty: q,
          date: _today,
          orderId: o.id || '',
          product: prod,
          branch: o.branch || '',
          note: 'LINE協賛注文',
          createdBy: 'LINEフォーム'
        });
      }
    });
  }
}

// ───────────────────────────────────────────────
// 注文をスプシから削除（注文IDで照合・なければ timestamp+store で照合）
// ───────────────────────────────────────────────
function deleteOrderFromSheet(id, timestamp, store, amount) {
  const sh = getSheet('注文受付');
  const rows = sh.getDataRange().getValues();
  // ① 注文ID（O列）で照合（最優先・確実）
  if (id) {
    for (let i = rows.length - 1; i >= 1; i--) {
      if (String(rows[i][14] || '') === String(id)) { sh.deleteRow(i + 1); return; }
    }
  }
  // ② timestamp(A列) + store(D列) で照合
  if (timestamp && store) {
    for (let i = rows.length - 1; i >= 1; i--) {
      if (String(rows[i][0]) === String(timestamp) && String(rows[i][3]) === String(store)) { sh.deleteRow(i + 1); return; }
    }
  }
  // ③ store(D列) + amount(G列) で照合（旧注文の最終手段・同一店舗同一金額の最新1件）
  if (store && amount) {
    for (let i = rows.length - 1; i >= 1; i--) {
      if (String(rows[i][3]) === String(store) && String(rows[i][6]) === String(amount)) { sh.deleteRow(i + 1); return; }
    }
  }
}

// 協賛ログをスプシから削除（ID照合）
function deleteSponsorLogFromSheet(id) {
  const sh = getSheet('協賛ログ');
  if (sh.getLastRow() === 0) return;
  const rows = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][0]) === String(id)) {
      sh.deleteRow(i + 1);
      return;
    }
  }
}

// 顧客をスプシから削除（ヘッダー名で「顧客ID」優先・なければ「店舗名」で照合）
// ※列構成に依存しないようヘッダー名で列を特定する
function deleteCustomerFromSheet(id, name) {
  const sh = getSheet('顧客');
  const d = sh.getDataRange().getValues();
  if (d.length === 0) return;
  const hdrs = d[0].map(String);
  const idCol   = hdrs.indexOf('顧客ID');
  const nameCol = hdrs.indexOf('店舗名');
  for (let i = d.length - 1; i >= 1; i--) {
    if (id && idCol >= 0 && String(d[i][idCol]) === String(id)) { sh.deleteRow(i + 1); return; }
    if (name && nameCol >= 0 && String(d[i][nameCol]) === String(name)) { sh.deleteRow(i + 1); return; }
  }
}

// ─── 協賛 名称⇔ID 変換マップ（協賛シート A=ID, B=ルール名 を参照） ───
function getSponsorMaps() {
  const id2name = {}, name2id = {};
  try {
    const sh = getSheet('協賛');
    const d = sh.getDataRange().getValues();
    for (let i = 1; i < d.length; i++) {
      const id   = String(d[i][0] || '').trim();
      const name = String(d[i][1] || '').trim();
      if (id) {
        id2name[id] = name;
        if (name) name2id[name] = id;
      }
    }
  } catch (e) { /* 協賛シートが無い場合は空マップ */ }
  return { id2name: id2name, name2id: name2id };
}

// 電話列（店舗/代表者/配達先1〜5）を文字列書式にして先頭0を保持。値も文字列で書き直すので初回保存から0が残る
function _fixPhoneCells(sh, rowNum, headerRow, rowData) {
  for (let col = 0; col < headerRow.length; col++) {
    if (String(headerRow[col]).indexOf('電話') < 0) continue;
    const cell = sh.getRange(rowNum, col + 1);
    cell.setNumberFormat('@');
    cell.setValue((rowData && rowData[col] != null) ? String(rowData[col]).replace(/^'+/, '') : '');
  }
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
  // ヘッダー（68列固定）= 基本22列（協賛ON/OFF + 適用協賛名称含む） + 配達先45列
  // ※旧「購入条件/協賛本数/対象商品」は廃止。W列に「適用協賛」(名称)を配置。
  // ※BP列以降は詳細情報+顧客ID+備考+商品別掛率（動的追加）
  const HEADER = [
    '拠点', '弊社担当', '店舗名', 'フリガナ', '店舗電話',
    '代表者名', '代表者電話', '担当者リスト',
    '郵便番号', '都道府県', '市町村郡', '番地', '建物名',
    'LINE ID', 'LINE UID',
    '紹介者', '請求書宛名',
    '支払方法',
    '振込名義①', '振込名義②', '振込名義③',
    '協賛', '適用協賛',
    // 配達先1〜5（各9項目 = 45列）
    '配達先1_名称', '配達先1_郵便番号', '配達先1_都道府県', '配達先1_市町村郡', '配達先1_番地', '配達先1_建物名', '配達先1_担当者', '配達先1_電話', '配達先1_メモ',
    '配達先2_名称', '配達先2_郵便番号', '配達先2_都道府県', '配達先2_市町村郡', '配達先2_番地', '配達先2_建物名', '配達先2_担当者', '配達先2_電話', '配達先2_メモ',
    '配達先3_名称', '配達先3_郵便番号', '配達先3_都道府県', '配達先3_市町村郡', '配達先3_番地', '配達先3_建物名', '配達先3_担当者', '配達先3_電話', '配達先3_メモ',
    '配達先4_名称', '配達先4_郵便番号', '配達先4_都道府県', '配達先4_市町村郡', '配達先4_番地', '配達先4_建物名', '配達先4_担当者', '配達先4_電話', '配達先4_メモ',
    '配達先5_名称', '配達先5_郵便番号', '配達先5_都道府県', '配達先5_市町村郡', '配達先5_番地', '配達先5_建物名', '配達先5_担当者', '配達先5_電話', '配達先5_メモ'
    // ↑ 22 + 45 = 67列。BP列以降は詳細情報・顧客ID・備考・商品別掛率（動的追加）
  ];

  const bankN = Array.isArray(c.bankNames) ? c.bankNames : (c.bankName ? [c.bankName] : []);
  // 協賛: sponsorIds(ID配列) → 協賛名のカンマ区切りに変換（スプシ表示用）
  const _spoMaps = getSponsorMaps();
  const sponsorNames = (Array.isArray(c.sponsorIds) ? c.sponsorIds : [])
    .map(function(id){ return _spoMaps.id2name[id] || id; })
    .filter(Boolean).join(',');
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
    c.sponsorEnabled ? 'あり' : 'なし',  // V: 協賛(ON/OFF)
    sponsorNames                         // W: 適用協賛(名称カンマ区切り)
  ];
  // 配達先5件分（各9列）を追加 — X列以降
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
  // baseRow は 22 + 45 = 67要素
  // BP列以降は詳細情報 → 商品別掛率 → 備考（動的追加）

  // 商品別掛率マップ
  const ratesMap = {};
  if (c.productRates && Array.isArray(c.productRates)) {
    c.productRates.forEach(function(pr) {
      if (pr.key) {
        ratesMap[pr.key] = pr.rate || '';
        // 商品別の価格モード（新値/旧値）を「商品名_モード」として保存
        if (pr.priceMode) {
          ratesMap[pr.key + '_モード'] = pr.priceMode;
        }
        // 商品別の税（税込/税抜/未設定）を「商品名_税」として保存
        if (pr.taxType) {
          ratesMap[pr.key + '_税'] = pr.taxType;
        }
      }
    });
  }

  // 詳細情報（分析用）— 動的列として扱う（支払方法は固定列、税・価格モードは商品別なので除外）
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
    '共通価格モード': c.commonPriceMode || '新値',
    '共通税':       c.commonTaxType  || '税込',
    '掛率':         c.rate || '',
    '顧客ID':       c.id || '',
    '協賛即時反映': c.sponsorInstant ? 'あり' : 'なし',
    '送料無料':     c.freeShipping ? 'あり' : 'なし',
    '個別送料':     (c.shipFee !== undefined && c.shipFee !== null && String(c.shipFee).trim() !== '') ? c.shipFee : '',
    '協賛を送料本数に含める': (c.shipIncludeSponsor === 'あり' || c.shipIncludeSponsor === 'なし') ? c.shipIncludeSponsor : ''
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

  // 最終行 = 基本67列（22 + 配達先45） + 動的列（詳細情報 + 顧客ID + 商品別掛率 + 備考）
  const row = baseRow.slice();
  for (let col = baseRow.length; col < headerRow.length; col++) {
    row.push(dynamicFields[headerRow[col]] || '');
  }

  // 既存行を更新（顧客IDがあればID、無ければ店舗名=C列で照合）
  // ★ isResetOrEmpty なら既存行はないので走査スキップ
  if (!isResetOrEmpty) {
    const _idCol = headerRow.indexOf('顧客ID');
    for (let i = 1; i < d.length; i++) {
      const _match = c.id ? (_idCol >= 0 && d[i][_idCol] === c.id) : (d[i][2] === c.name);
      if (_match) {
        sh.getRange(i+1, 1, 1, row.length).setValues([row]);
        _fixPhoneCells(sh, i+1, headerRow, row);
        return;
      }
    }
  }

  // 新規追加
  const newRow = sh.getLastRow() + 1;
  sh.appendRow(row);
  _fixPhoneCells(sh, newRow, headerRow, row);
}

// ── 顧客の基本67列ヘッダー（saveCustomerToSheet と同一定義）──
function _custHeader() {
  return [
    '拠点', '弊社担当', '店舗名', 'フリガナ', '店舗電話',
    '代表者名', '代表者電話', '担当者リスト',
    '郵便番号', '都道府県', '市町村郡', '番地', '建物名',
    'LINE ID', 'LINE UID',
    '紹介者', '請求書宛名',
    '支払方法',
    '振込名義①', '振込名義②', '振込名義③',
    '協賛', '適用協賛',
    '配達先1_名称', '配達先1_郵便番号', '配達先1_都道府県', '配達先1_市町村郡', '配達先1_番地', '配達先1_建物名', '配達先1_担当者', '配達先1_電話', '配達先1_メモ',
    '配達先2_名称', '配達先2_郵便番号', '配達先2_都道府県', '配達先2_市町村郡', '配達先2_番地', '配達先2_建物名', '配達先2_担当者', '配達先2_電話', '配達先2_メモ',
    '配達先3_名称', '配達先3_郵便番号', '配達先3_都道府県', '配達先3_市町村郡', '配達先3_番地', '配達先3_建物名', '配達先3_担当者', '配達先3_電話', '配達先3_メモ',
    '配達先4_名称', '配達先4_郵便番号', '配達先4_都道府県', '配達先4_市町村郡', '配達先4_番地', '配達先4_建物名', '配達先4_担当者', '配達先4_電話', '配達先4_メモ',
    '配達先5_名称', '配達先5_郵便番号', '配達先5_都道府県', '配達先5_市町村郡', '配達先5_番地', '配達先5_建物名', '配達先5_担当者', '配達先5_電話', '配達先5_メモ'
  ];
}

// ── 顧客1件 → {base:基本67列, dyn:動的列マップ}（saveCustomerToSheet と同一ロジック）──
function _custBaseAndDyn(c, spoMaps) {
  const tel      = String(c.tel      || '').replace(/^'+/, '');
  const ownerTel = String(c.ownerTel || '').replace(/^'+/, '');
  let contacts = '';
  if (c.contacts && Array.isArray(c.contacts)) contacts = JSON.stringify(c.contacts);
  else if (typeof c.contacts === 'string')     contacts = c.contacts;
  const bankN = Array.isArray(c.bankNames) ? c.bankNames : (c.bankName ? [c.bankName] : []);
  const sponsorNames = (Array.isArray(c.sponsorIds) ? c.sponsorIds : [])
    .map(function(id){ return spoMaps.id2name[id] || id; }).filter(Boolean).join(',');
  const baseRow = [
    c.branch||'', c.staff||'', c.name||'', c.kana||'', tel,
    c.owner||'', ownerTel, contacts,
    c.zip||'', c.pref||'', c.city||'', c.addr1||'', c.addr2||'',
    c.lineId||'', c.lineUid||'',
    c.referrer||'', c.invName||'',
    c.payMethod||'',
    bankN[0]||'', bankN[1]||'', bankN[2]||'',
    c.sponsorEnabled ? 'あり' : 'なし', sponsorNames
  ];
  const deliveries = Array.isArray(c.deliveries) ? c.deliveries : [];
  for (let n = 0; n < 5; n++) {
    const d = deliveries[n] || {};
    baseRow.push(d.name||'', d.zip||'', d.pref||'', d.city||'', d.addr1||'', d.addr2||'', d.contact||'', d.tel||'', d.memo||'');
  }
  const ratesMap = {};
  if (c.productRates && Array.isArray(c.productRates)) {
    c.productRates.forEach(function(pr) {
      if (pr.key) {
        ratesMap[pr.key] = pr.rate || '';
        if (pr.priceMode) ratesMap[pr.key + '_モード'] = pr.priceMode;
        if (pr.taxType)   ratesMap[pr.key + '_税']   = pr.taxType;
      }
    });
  }
  const detailFields = {
    '業態': c.bizType||'', '開店年月': c.openDate||'', '席数': c.seats||'',
    '経営者年代': c.ownerAge||'', '主要客層': c.customerBase||'', '主力商品傾向': c.productTrend||'',
    '競合関係': c.competitor||'', '取引開始日': c.dealStart||'', '取引区分': c.dealChannel||'',
    '営業日時': c.bizHours||'', '平均客単価': c.avgSpend||'',
    '共通価格モード': c.commonPriceMode||'新値', '共通税': c.commonTaxType||'税込',
    '掛率': c.rate||'', '顧客ID': c.id||'',
    '協賛即時反映': c.sponsorInstant ? 'あり' : 'なし',
    '送料無料': c.freeShipping ? 'あり' : 'なし',
    '個別送料': (c.shipFee !== undefined && c.shipFee !== null && String(c.shipFee).trim() !== '') ? c.shipFee : '',
    '協賛を送料本数に含める': (c.shipIncludeSponsor === 'あり' || c.shipIncludeSponsor === 'なし') ? c.shipIncludeSponsor : ''
  };
  const dyn = Object.assign({}, detailFields, ratesMap, { '備考': c.note || '' });
  return { base: baseRow, dyn: dyn };
}

// ── 顧客の一括保存（メモリ上で全件マージ → 1回だけ書込。顧客IDで照合）──
// CSV取込など大量件数を1リクエストで安全・高速に書き込む。
// ※1件ずつ更新すると遅くロック待ちを超えてバッチ欠落するため、全体を一括setValuesする。
function saveCustomersBulk(payload) {
  const list = (payload && payload.customers) || [];
  if (!list.length) return json({ ok: true, added: 0, updated: 0, total: 0 });
  const lock = LockService.getScriptLock();
  try { lock.waitLock(300000); } catch (e) { return json({ ok: false, error: 'lock-timeout' }); }
  try {
    const sh = getSheet('顧客');
    const HEADER = _custHeader();
    let d = sh.getDataRange().getValues();
    let headerRow;
    const a1 = (d.length > 0 && d[0].length > 0) ? String(d[0][0]).trim() : '';
    if (d.length === 0 || !a1) { headerRow = HEADER.slice(); d = [headerRow.slice()]; }
    else { headerRow = d[0].map(String); }

    const spoMaps = getSponsorMaps();
    const parts = list.map(function(c){ return _custBaseAndDyn(c, spoMaps); });
    // 動的列をヘッダーへ集約
    parts.forEach(function(p){ Object.keys(p.dyn).forEach(function(k){ if (headerRow.indexOf(k) < 0) headerRow.push(k); }); });
    const W = headerRow.length;
    const idCol = headerRow.indexOf('顧客ID');   // ★ ユニークキー = 顧客ID

    // 既存行をメモリに展開（W列に正規化）。顧客ID・店舗名で索引。
    const out = [headerRow.slice()];
    const idToRow = {}, nameToRow = {};
    for (let i = 1; i < d.length; i++) {
      const r = d[i].slice();
      while (r.length < W) r.push('');
      if (r.length > W) r.length = W;
      out.push(r);
      const ri = out.length - 1;
      if (idCol >= 0) { const id = r[idCol]; if (id !== '' && id != null && idToRow[id] === undefined) idToRow[id] = ri; }
      const nm = r[2]; if (nm !== '' && nm != null && nameToRow[nm] === undefined) nameToRow[nm] = ri;
    }

    let added = 0, updated = 0;
    parts.forEach(function(p, i){
      const c = list[i];
      const row = p.base.slice();
      for (let col = p.base.length; col < W; col++) row.push(p.dyn[headerRow[col]] || '');
      while (row.length < W) row.push('');
      // 照合：顧客IDがあればIDのみ（同名でも別行）／IDが無い場合のみ店舗名
      let ri = -1;
      if (c.id) { if (idToRow[c.id] !== undefined) ri = idToRow[c.id]; }
      else      { if (nameToRow[c.name] !== undefined) ri = nameToRow[c.name]; }
      if (ri >= 0) { out[ri] = row; updated++; }
      else {
        out.push(row); const nri = out.length - 1;
        if (c.id) idToRow[c.id] = nri; else nameToRow[c.name] = nri; // 同一バッチ内の重複キーは集約
        added++;
      }
    });

    // 全体を1回で書き込み（高速・タイムアウト回避）
    sh.clearContents();
    sh.getRange(1, 1, out.length, W).setValues(out);
    // 電話列（店舗/代表者/配達先1〜5）すべて文字列化し先頭0を保持
    const _hdr0 = out[0] || [];
    for (let _c = 0; _c < W; _c++) {
      if (String(_hdr0[_c]).indexOf('電話') < 0) continue;
      sh.getRange(1, _c + 1, out.length, 1).setNumberFormat('@');
      sh.getRange(1, _c + 1, out.length, 1).setValues(out.map(function(r){ return [ (r[_c] != null) ? String(r[_c]).replace(/^'+/, '') : '' ]; }));
    }
    return json({ ok: true, added: added, updated: updated, total: list.length, rows: out.length - 1 });
  } finally {
    lock.releaseLock();
  }
}

// 注文受付の拡張列ヘッダーを確保（P=伝票No, Q=集金者, R=集金日）
function _ensureOrderHeaders(sh) {
  var need = { 16: '伝票No', 17: '集金者', 18: '集金日' };
  Object.keys(need).forEach(function(c){
    var cell = sh.getRange(1, Number(c));
    if (!cell.getValue()) cell.setValue(need[c]);
  });
}

// 集金（支払状況）更新：注文ID＝O列(15列目, idx14)で照合。集金者・集金日も保存。
function updatePayInSheet(orderId, status, collector, collectedDate) {
  const sh = getSheet('注文受付');
  _ensureOrderHeaders(sh);
  const d  = sh.getDataRange().getValues();
  for (let i = 1; i < d.length; i++) {
    if (String(d[i][14] || '').trim() === String(orderId).trim()) {
      sh.getRange(i+1, 9).setValue(status || '集金済み');                         // I列 支払状況
      if (collector    !== undefined && collector    !== '') sh.getRange(i+1, 17).setValue(collector);    // Q列 集金者
      if (collectedDate!== undefined && collectedDate!== '') sh.getRange(i+1, 18).setValue(collectedDate); // R列 集金日
      return;
    }
  }
}

// 発送処理：伝票No（送り状番号）とステータスを保存。注文ID＝O列で照合。
function updateShipInSheet(orderId, tracking, status) {
  const sh = getSheet('注文受付');
  _ensureOrderHeaders(sh);
  const d  = sh.getDataRange().getValues();
  for (let i = 1; i < d.length; i++) {
    if (String(d[i][14] || '').trim() === String(orderId).trim()) {
      if (tracking !== undefined && tracking !== '') sh.getRange(i+1, 16).setValue(tracking); // P列 伝票No
      if (status) sh.getRange(i+1, 10).setValue(status);                                      // J列 ステータス
      return;
    }
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
  // 空データで在庫を消さない安全策（S.stockが未読込のときの誤消去を防ぐ）
  if (!stockData || !Array.isArray(stockData) || stockData.length === 0) return;
  const sh = getSheet('在庫');
  sh.clearContents();
  sh.getRange(1, 1, 1, 3).setValues([['拠点', '商品', '在庫数']]);  // ヘッダー行
  const rows = stockData.map(s => [s.branch, s.product, s.qty]);
  if (rows.length > 0) {
    sh.getRange(2, 1, rows.length, 3).setValues(rows);  // データは2行目から
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

  // アプリアイコンへ赤バッジ通知（OneSignal）
  try { pushToDevices('📩 新規注文', (o.store || 'お客様') + ' から注文が入りました', '', getNewOrders('').length); } catch(e){}
}

// ─── 新規注文取得API（管理画面のポーリング用） ──────────
function getNewOrders(since) {
  const rows = getSheet('注文受付').getDataRange().getValues();
  if (rows.length <= 1) return [];

  const newOrders = [];
  // ヘッダー後の行を逆順で確認（新しいものから）
  for (let i = rows.length - 1; i >= 1; i--) {
    const r = rows[i];
    // 注文IDは O列(r[14]) の実IDを優先（手動登録 ord-xxx / LIFF liff-xxx と一致させ重複取込を防ぐ）
    // O列が空の旧データのみ timestamp_store で代替
    const realId = String(r[14] || '').trim();
    const orderId = realId || (String(r[0] || '') + '_' + String(r[3] || ''));
    if (since && orderId === since) break;
    // ▼ お知らせは「新規受付（=新着・未対応）」のみ。発注済（対応中/出荷済など）は出さない
    var _st = String(r[9] || '').trim();
    if (_st === '対応中' || _st === '出荷済' || _st === '発送' || _st === '完了' || _st === '発注済' || _st === 'キャンセル') continue;
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

// ===== 注文フォーム高速化: 共通計算ヘルパー（既存ハンドラと getOrderInit が共用） =====
function _computeStock() {
    const rows = getSheet('在庫').getDataRange().getValues();
    // 1行目がヘッダー（拠点/branch）ならスキップ。ヘッダーなしの旧データも許容
    const data = (rows.length > 0 && (String(rows[0][0]) === '拠点' || String(rows[0][0]) === 'branch')) ? rows.slice(1) : rows;
    return (data.filter(r => r[0] && r[1]).map(r => ({ branch:r[0], product:r[1], qty:r[2] })) );
}

function _computeMySponsor(uid) {
    if (!uid) return ({ ok:false, error:'uid missing' });
    // UID → 顧客ID・店舗名 を特定
    const crows = getSheet('顧客').getDataRange().getValues();
    if (crows.length === 0) return ({ ok:false, error:'no customers' });
    const chd = crows[0].map(String);
    const uidCol  = chd.indexOf('LINE UID');
    const nameCol = chd.indexOf('店舗名');
    const idCol   = chd.indexOf('顧客ID');
    if (uidCol < 0) return ({ ok:false, error:'headers' });
    let custId = '', store = '';
    for (let i = 1; i < crows.length; i++) {
      if (String(crows[i][uidCol]).trim() === String(uid).trim()) {
        store  = nameCol >= 0 ? String(crows[i][nameCol]).trim() : '';
        custId = idCol >= 0 ? String(crows[i][idCol]).trim() : '';
        // 顧客IDが空なら店舗名から安定ID生成（checkLineUserと同じ式）
        if (!custId && store) {
          let _h = 0; for (let _ci = 0; _ci < store.length; _ci++) { _h = ((_h << 5) - _h + store.charCodeAt(_ci)) | 0; }
          custId = 'cust-sync-' + Math.abs(_h).toString(36);
        }
        break;
      }
    }
    if (!custId) return ({ ok:false, error:'unregistered' });
    // 協賛ログを集計（商品別 earned − used）
    const sh = getSheet('協賛ログ');
    const byProd = {};
    if (sh.getLastRow() > 1) {
      const lr = sh.getDataRange().getValues();
      for (let i = 1; i < lr.length; i++) {
        const r = lr[i];
        if (!r[0]) continue;
        if (String(r[1]).trim() !== custId) continue;     // 顧客ID一致
        const type = String(r[3]);
        const qty  = (typeof r[4] === 'number') ? r[4] : (parseInt(r[4]) || 0);
        const prod = String(r[9] || '商品指定なし');
        if (!(prod in byProd)) byProd[prod] = 0;
        if (type === 'earned') byProd[prod] += qty;
        if (type === 'used')   byProd[prod] -= qty;
      }
    }
    // 残数>0 の商品だけ返す
    const balance = {};
    let total = 0;
    Object.keys(byProd).forEach(function(p){ if (byProd[p] > 0) { balance[p] = byProd[p]; total += byProd[p]; } });
    return ({ ok:true, custId: custId, store: store, total: total, byProduct: balance });
}

function _computeAuth(uid) {
    if (!uid) return ({ ok:false, error:'uid missing' });
    const rows = getSheet('顧客').getDataRange().getValues();
    if (rows.length === 0) return ({ ok:false });
    const hdrs = rows[0].map(String);

    // ヘッダー名 → 列インデックス（列の物理位置に依存せず読む）
    const colIdx = {};
    hdrs.forEach(function(h, idx){ if (colIdx[h] === undefined) colIdx[h] = idx; });
    // 固定ヘッダー集合（これ以外を商品別掛率として処理）
    const FIXED_H = {};
    ['拠点','弊社担当','店舗名','フリガナ','店舗電話','代表者名','代表者電話','担当者リスト',
     '郵便番号','都道府県','市町村郡','番地','建物名','LINE ID','LINE UID','紹介者','請求書宛名',
     '支払方法','振込名義①','振込名義②','振込名義③','協賛','適用協賛',
     '購入条件','協賛本数','対象商品','適用協賛ID'
    ].forEach(function(h){ FIXED_H[h] = true; });
    for (let n = 1; n <= 5; n++) {
      ['名称','郵便番号','都道府県','市町村郡','番地','建物名','担当者','電話','メモ'].forEach(function(s){
        FIXED_H['配達先'+n+'_'+s] = true;
      });
    }
    // 詳細情報・備考のヘッダー名（商品別掛率と区別するため）
    const DETAIL_HEADERS = ['業態','開店年月','席数','経営者年代','主要客層',
                             '主力商品傾向','競合関係','取引開始日','取引区分',
                             '営業日時','平均客単価','共通価格モード','共通税','掛率','顧客ID','備考',
                             '協賛即時反映','送料無料','個別送料','協賛を送料本数に含める'];

    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][14]) !== uid) continue;      // O列: LINE UID
      const cv = function(name){ const ci = colIdx[name]; return (ci !== undefined && rows[i][ci] != null) ? String(rows[i][ci]) : ''; };

      // ── 配達先5件（ヘッダー名ベース＝列位置が変わっても正しく読む）────
      const deliveries = [];
      for (let n = 1; n <= 5; n++) {
        const d = {
          name:    cv('配達先'+n+'_名称'),
          zip:     cv('配達先'+n+'_郵便番号'),
          pref:    cv('配達先'+n+'_都道府県'),
          city:    cv('配達先'+n+'_市町村郡'),
          addr1:   cv('配達先'+n+'_番地'),
          addr2:   cv('配達先'+n+'_建物名'),
          contact: cv('配達先'+n+'_担当者'),
          tel:     cv('配達先'+n+'_電話'),
          memo:    cv('配達先'+n+'_メモ')
        };
        if (d.name || d.zip || d.addr1 || d.contact) {
          deliveries.push(d);
        }
      }

      // ── 商品別掛率 + 価格モード + 税（BS列以降、詳細情報・備考を除く）──
      // 「商品名」= 掛率、「商品名_モード」= 価格モード、「商品名_税」= 税
      const prodMap = {};   // { 商品名: { rate, priceMode, taxType } }
      for (let col = 0; col < hdrs.length; col++) {
        const h = hdrs[col];
        if (!h || FIXED_H[h] || DETAIL_HEADERS.indexOf(h) >= 0) continue; // 固定列・詳細情報・備考はスキップ
        if (rows[i][col] === '' || rows[i][col] === undefined) continue;
        // サフィックスで分類
        if (h.length > 4 && h.slice(-4) === '_モード') {
          const prodName = h.slice(0, -4);
          if (!prodMap[prodName]) prodMap[prodName] = {};
          prodMap[prodName].priceMode = String(rows[i][col]);
        } else if (h.length > 2 && h.slice(-2) === '_税') {
          const prodName = h.slice(0, -2);
          if (!prodMap[prodName]) prodMap[prodName] = {};
          prodMap[prodName].taxType = String(rows[i][col]);
        } else {
          if (!prodMap[h]) prodMap[h] = {};
          prodMap[h].rate = String(rows[i][col]);
        }
      }
      const pr = [];
      Object.keys(prodMap).forEach(function(prodName) {
        const entry = prodMap[prodName];
        if (entry.rate) {
          pr.push({
            key: prodName,
            rate: entry.rate,
            priceMode: entry.priceMode || '新値',
            taxType:   entry.taxType   || '税込'
          });
        }
      });

      // ── 共通掛率：顧客シートの「掛率」列（CD列）を最優先で読む ──────
      // 列位置ではなくヘッダー名「掛率」で検索するため、列が移動しても正しく読める
      let defaultRate = String(cv('掛率') || '').trim();
      if (!defaultRate) {
        const yellowRate = pr.find(function(r) { return r.key === 'イエロー'; });
        if (yellowRate && yellowRate.rate) defaultRate = yellowRate.rate;
        else if (pr.length > 0 && pr[0].rate) defaultRate = pr[0].rate;
        else defaultRate = '7';
      }

      // 共通価格モード・共通税（ヘッダー名「共通価格モード」または「共通価格」に対応）
      const commonPriceMode = cv('共通価格モード') || cv('共通価格') || '新値';
      const commonTaxType   = cv('共通税') || '税込';

      // ── レスポンス ────────────────────────────────
      return ({
        ok:         true,
        sponsorEnabled: cv('協賛') === 'あり',  // 注文フォーム協賛セクション判定
        sponsorInstant: cv('協賛即時反映') === 'あり',  // 即時反映ON/OFF（同じ注文内で使えるか）
        sponsorRules: _getSponsorRulesForRow(cv),       // 適用協賛ルール（即時反映の付与計算用）
        freeShipping: cv('送料無料') === 'あり',          // 常時送料無料の顧客
        shipFee:      cv('個別送料'),                     // 個別送料（空ならデフォルトを使う）
        shipIncludeSponsor: cv('協賛を送料本数に含める'),  // 送料本数に協賛分を含めるか（''=全体設定に従う）
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
        commonPriceMode: commonPriceMode, // 共通価格モード（個別未指定時のフォールバック）
        commonTaxType:   commonTaxType,   // 共通税（個別未指定時のフォールバック）
      });
    }
    return ({ ok:false });
}

// 送料設定（全体共通）を読む。無ければデフォルトで自動作成。
// 返却: { defaultFee, freeQty, includeSponsor }
function _computeShippingConfig() {
  const sh = getSheet('送料設定');
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 4, 2).setValues([
      ['設定項目', '値'],
      ['デフォルト送料', 950],
      ['無料になる本数', 6],
      ['協賛無料を本数に含める', 'あり']
    ]);
  }
  const rows = sh.getDataRange().getValues();
  const map = {};
  for (let i = 1; i < rows.length; i++) { if (rows[i][0]) map[String(rows[i][0]).trim()] = rows[i][1]; }
  const num = function(v, d) {
    const n = (typeof v === 'number') ? v : parseInt(String(v == null ? '' : v).replace(/[^0-9.\-]/g, ''));
    return isNaN(n) ? d : n;
  };
  return {
    defaultFee:     num(map['デフォルト送料'], 950),
    freeQty:        num(map['無料になる本数'], 6),
    includeSponsor: String(map['協賛無料を本数に含める'] == null ? 'あり' : map['協賛無料を本数に含める']).trim() === 'あり'
  };
}

// 手動実行用：送料設定シートをセットアップ
function initShippingSheet() {
  const sh = getSheet('送料設定');
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 4, 2).setValues([['設定項目', '値'], ['デフォルト送料', 950], ['無料になる本数', 6], ['協賛無料を本数に含める', 'あり']]);
    Logger.log('✅ 送料設定シートを作成しました');
  } else {
    Logger.log('送料設定シートは既に存在します（行数: ' + sh.getLastRow() + '）');
  }
}

// 協賛ルールの日付を yyyy-MM-dd に正規化（日付型セル/スラッシュ表記の両方に対応）
// ※これを怠ると order.html 側の文字列比較で「未開始」と誤判定されルールが除外される
function _spDateStr(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v).slice(0, 10).replace(/\//g, '-');
}

// 顧客の「適用協賛」名称 → 協賛シートのルール配列を返す（即時反映の付与計算用）
function _getSponsorRulesForRow(cv) {
  const namesStr = (typeof cv === 'function') ? (cv('適用協賛') || '') : '';
  if (!String(namesStr).trim()) return [];
  const wanted = String(namesStr).split(',').map(function(s){ return s.trim(); }).filter(Boolean);
  if (!wanted.length) return [];
  const sh = getSheet('協賛');
  if (sh.getLastRow() <= 1) return [];
  const rows = sh.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0]) continue;
    const name = String(r[1] || '');
    if (wanted.indexOf(name) < 0) continue;
    out.push({
      name:          name,
      targetProduct: String(r[2] || 'all'),
      triggerQty:    (typeof r[3] === 'number') ? r[3] : (parseFloat(r[3]) || 0),
      rewardQty:     (typeof r[5] === 'number') ? r[5] : (parseFloat(r[5]) || 0),
      rewardProduct: String(r[6] || 'same'),
      startDate:     _spDateStr(r[7]),
      endDate:       _spDateStr(r[8]),
      enabled:       String(r[9] || '有効')
    });
  }
  return out;
}
