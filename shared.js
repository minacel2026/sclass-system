// ============================================================
// shared.js  Sclassランボルギーニ 全システム共通データ
// ★ ポータル・マスター・業務システム全てで共通利用
// ★ このファイルを変更すると全システムに反映される
// ============================================================

// ─── 商品マスタ ─────────────────────────────────────
// 商品名、価格、原価、補充閾値、画像パス
const PROD_LIST = [
  {key:'イエロー',     name:'イエロー',     fullName:'001: Extra Brut（イエロー）',     category:'スパークリング', cost:5000,  price:10500,  reorder:10, img:''},
  {key:'オレンジ',     name:'オレンジ',     fullName:'002: Rosé（オレンジ）',            category:'スパークリング', cost:10000, price:21000,  reorder:10, img:''},
  {key:'ブルー',       name:'ブルー',       fullName:'003: Blue Label',                  category:'スパークリング', cost:10000, price:21000,  reorder:10, img:''},
  {key:'ブラック',     name:'ブラック',     fullName:'004: Black Label',                 category:'スパークリング', cost:10000, price:21000,  reorder:10, img:''},
  {key:'プラチナ',     name:'プラチナ',     fullName:'005: Platinum',                    category:'スパークリング', cost:15000, price:33000,  reorder:6,  img:''},
  {key:'ゴールド',     name:'ゴールド',     fullName:'006: Gold',                        category:'スパークリング', cost:20000, price:44000,  reorder:6,  img:''},
  {key:'ルミナス',     name:'ルミナス',     fullName:'007: Luminous',                    category:'スパークリング', cost:25000, price:55000,  reorder:6,  img:''},
  {key:'レジェンド',   name:'レジェンド',   fullName:'008: Legend',                      category:'スパークリング', cost:25000, price:55000,  reorder:6,  img:''},
  {key:'さくら',       name:'さくら',       fullName:'009: Sakura',                      category:'スパークリング', cost:12000, price:25000,  reorder:6,  img:''},
  {key:'チェングレット', name:'チェングレット', fullName:'010: Cheng Leggero',             category:'スティルワイン', cost:3000,  price:6800,   reorder:12, img:''},
  {key:'チェングロッソ', name:'チェングロッソ', fullName:'011: Cheng Grosso',              category:'スティルワイン', cost:3000,  price:6800,   reorder:12, img:''},
  {key:'トレスコーネ', name:'トレスコーネ', fullName:'012: Trescone',                    category:'スティルワイン', cost:3500,  price:7800,   reorder:12, img:''},
  {key:'エラ',         name:'エラ',         fullName:'013: Ella',                        category:'スティルワイン', cost:4500,  price:10000,  reorder:10, img:''},
  {key:'トラミ',       name:'トラミ',       fullName:'014: Trami',                       category:'スティルワイン', cost:6000,  price:13000,  reorder:10, img:''},
  {key:'カンポレオーネ', name:'カンポレオーネ', fullName:'015: Campoleone',                category:'スティルワイン', cost:4500,  price:10000,  reorder:10, img:''},
  {key:'GIFTBOX',      name:'GIFTBOX',      fullName:'016: GIFTBOX',                     category:'ギフト',       cost:9000,  price:20000,  reorder:5,  img:''},
  {key:'コラチョロサ', name:'コラチョロサ', fullName:'017: Cora Chollosa',               category:'スティルワイン', cost:8000,  price:17000,  reorder:6,  img:''},
];

// ─── 商品マスタから派生する定数 ────────────────────
const PRODS   = PROD_LIST.map(p => p.key);
const PRICE   = Object.fromEntries(PROD_LIST.map(p => [p.key, p.price]));
const COST    = Object.fromEntries(PROD_LIST.map(p => [p.key, p.cost]));
const REORDER = Object.fromEntries(PROD_LIST.map(p => [p.key, p.reorder]));

// ─── 共通状態 ───────────────────────────────────────
const S = {
  branch:'both',        // 'both' / '仙台' / '京都'
  orders:[],            // 注文履歴
  customers:[],         // 顧客一覧
  stock:{仙台:{}, 京都:{}},
  stockDetail:{仙台:{}, 京都:{}},
  slog:[],              // 在庫ログ
  collections:[],       // 集金履歴
  invoices:[],          // 請求書
  bankEntries:[],       // 通帳データ
  bankHistory:[],       // 通帳履歴
  company:{},           // 弊社情報
  purchases:[],         // 仕入履歴
  suppliers:['マルシェジャパン','オーク酒販'],
  cashEntries:[],       // 現金出納帳
  dailyReps:[],         // 営業日報
  gasUrl: localStorage.getItem('gasUrl') || '',
};

// ─── 共通ヘルパー関数 ───────────────────────────────
function today()       { return new Date().toISOString().slice(0,10); }
function fmt(n)        { return (Number(n)||0).toLocaleString('ja-JP'); }
function yen(n)        { return '¥' + fmt(n); }
function calcAmt(o) {
  const tax = o.tax === 'taxあり' ? 1.1 : 1;
  const cond = (o.condition || 10) / 10;
  let total = 0;
  Object.entries(o.items || {}).forEach(([p,q]) => {
    total += Math.round((PRICE[p] || 0) * cond) * tax * (q || 0);
  });
  return Math.round(total);
}

// ─── LocalStorage 読み書き ──────────────────────────
function saveState() {
  try {
    localStorage.setItem('lmb_orders',      JSON.stringify(S.orders));
    localStorage.setItem('lmb_customers',   JSON.stringify(S.customers));
    localStorage.setItem('lmb_stock',       JSON.stringify(S.stock));
    localStorage.setItem('lmb_stockDetail', JSON.stringify(S.stockDetail));
    localStorage.setItem('lmb_slog',        JSON.stringify(S.slog));
    localStorage.setItem('lmb_collections', JSON.stringify(S.collections));
    localStorage.setItem('lmb_invoices',    JSON.stringify(S.invoices));
    localStorage.setItem('lmb_bankEntries', JSON.stringify(S.bankEntries));
    localStorage.setItem('lmb_bankHistory', JSON.stringify(S.bankHistory));
    localStorage.setItem('lmb_company',     JSON.stringify(S.company));
    localStorage.setItem('lmb_purchases',   JSON.stringify(S.purchases));
    localStorage.setItem('lmb_cashEntries', JSON.stringify(S.cashEntries));
    localStorage.setItem('lmb_dailyReps',   JSON.stringify(S.dailyReps));
  } catch(e) { console.warn('saveState error', e); }
}

function loadState() {
  try {
    const keys = ['orders','customers','stock','stockDetail','slog','collections',
                  'invoices','bankEntries','bankHistory','company','purchases',
                  'cashEntries','dailyReps'];
    keys.forEach(k => {
      const v = localStorage.getItem('lmb_' + k);
      if (v) {
        try { S[k] = JSON.parse(v); } catch(e) {}
      }
    });
    S.gasUrl = localStorage.getItem('gasUrl') || '';
  } catch(e) { console.warn('loadState error', e); }
}

// ─── トースト通知 ───────────────────────────────────
function toast(msg, type) {
  let t = document.getElementById('toast');
  if (!t) {
    t = document.createElement('div');
    t.id = 'toast';
    t.style.cssText = 'position:fixed;bottom:2rem;left:50%;transform:translateX(-50%);background:#161616;border:1px solid rgba(201,168,76,.4);color:#E8C97A;padding:.8rem 1.5rem;border-radius:4px;font-size:.85rem;z-index:9999;opacity:0;transition:opacity .2s;letter-spacing:.05em';
    document.body.appendChild(t);
  }
  t.textContent = msg;
  t.style.opacity = '1';
  if (type === 'error') t.style.borderColor = 'rgba(224,82,82,.5)';
  else t.style.borderColor = 'rgba(201,168,76,.4)';
  setTimeout(() => { t.style.opacity = '0'; }, 2500);
}

// ─── GAS API 共通関数 ─────────────────────────────
async function gasGet(action, params) {
  if (!S.gasUrl) { toast('GAS URLが未設定です', 'error'); return null; }
  const q = new URLSearchParams({ action, ...params }).toString();
  try {
    const r = await fetch(S.gasUrl + '?' + q);
    return await r.json();
  } catch(e) {
    toast('通信エラー: ' + e.message, 'error');
    return null;
  }
}

async function gasPost(data) {
  if (!S.gasUrl) { toast('GAS URLが未設定です', 'error'); return null; }
  try {
    await fetch(S.gasUrl, {
      method: 'POST',
      mode: 'no-cors',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(data)
    });
    return { ok: true };
  } catch(e) {
    toast('送信エラー: ' + e.message, 'error');
    return null;
  }
}



// ─── グローバル公開（他スクリプトからの参照用） ─────────
if (typeof window !== 'undefined') {
  window.S         = S;
  window.PROD_LIST = PROD_LIST;
  window.PRODS     = PRODS;
  window.PRICE     = PRICE;
  window.COST      = COST;
  window.REORDER   = REORDER;
  window.calcAmt   = calcAmt;
  window.today     = today;
  window.fmt       = fmt;
  window.yen       = yen;
  window.saveState = saveState;
  window.loadState = loadState;
  window.toast     = toast;
  window.gasGet    = gasGet;
  window.gasPost   = gasPost;
}


// ─── 起動時に自動実行 ──────────────────────────────
if (typeof window !== 'undefined') {
  loadState();
}
