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

  if (action === 'ping') return json({ status: 'ok', version: 'stock-sync-v2', updateAllStock: true });

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
      status:r[9], note:r[10], uid:r[13], id:r[14]
    }))});
  }

  if (action === 'getStock') {
    const rows = getSheet('在庫').getDataRange().getValues();
    // 1行目がヘッダー（拠点/branch）ならスキップ。ヘッダーなしの旧データも許容
    const data = (rows.length > 0 && (String(rows[0][0]) === '拠点' || String(rows[0][0]) === 'branch')) ? rows.slice(1) : rows;
    return json({ stock: data.filter(r => r[0] && r[1]).map(r => ({ branch:r[0], product:r[1], qty:r[2] })) });
  }

  if (action === 'checkLineUser') {
    const uid = e.parameter.uid;
    if (!uid) return json({ ok:false, error:'uid missing' });
    const rows = getSheet('顧客').getDataRange().getValues();
    if (rows.length === 0) return json({ ok:false });
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
                             '営業日時','平均客単価','共通価格モード','共通税','掛率','顧客ID','備考'];

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

      // ── 共通掛率：商品別掛率から代表値を取得 ──────
      // イエロー掛率を優先、なければ最初の商品別掛率、なければデフォルト7
      let defaultRate = '7';
      const yellowRate = pr.find(function(r) { return r.key === 'イエロー'; });
      if (yellowRate && yellowRate.rate) {
        defaultRate = yellowRate.rate;
      } else if (pr.length > 0 && pr[0].rate) {
        defaultRate = pr[0].rate;
      }

      // 共通価格モード・共通税を読み取る
      const commonPriceMode = cv('共通価格モード') || '新値';
      const commonTaxType   = cv('共通税') || '税込';

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
        commonPriceMode: commonPriceMode, // 共通価格モード（個別未指定時のフォールバック）
        commonTaxType:   commonTaxType,   // 共通税（個別未指定時のフォールバック）
      });
    }
    return json({ ok:false });
  }

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
      '共通価格モード':'commonPriceMode','共通税':'commonTaxType','掛率':'_rateRaw','顧客ID':'_custId'
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
      return result;
    })});
  }

  // ★ ブラウザno-cors対策：GETのpayloadパラメータで書き込みを受け取る
  if (e.parameter.payload) {
    try {
      const body = JSON.parse(e.parameter.payload);
      if (body.action === 'saveCustomer')    { saveCustomerToSheet(body); return json({ok:true}); }
      if (body.action === 'deleteCustomer')  { deleteCustomerFromSheet(body.id, body.name); return json({ok:true}); }
      if (body.action === 'saveOrder')       { saveOrderToSheet(body);    return json({ok:true}); }
      if (body.action === 'updatePayStatus') { updatePayInSheet(body.orderId, body.payStatus); return json({ok:true}); }
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

  if (body.action === 'saveOrder')       { saveOrderToSheet(body);    return json({ok:true}); }
  if (body.action === 'updatePayStatus') { updatePayInSheet(body.orderId, body.payStatus); return json({ok:true}); }
  if (body.action === 'updateStock')     { updateStockInSheet(body.branch, body.product, body.qty); return json({ok:true}); }
  if (body.action === 'updateAllStock')  { updateAllStockInSheet(body.stockData); return json({ok:true}); }
  if (body.action === 'saveCustomer')    { saveCustomerToSheet(body); return json({ok:true}); }
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
    '顧客ID':       c.id || ''
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