// OneSignal 本体（プッシュ受信・通知表示を担当）
importScripts("https://cdn.onesignal.com/sdks/web/v16/OneSignalSDK.sw.js");

// ▼ アプリアイコンの赤い数字バッジ（Badging API）
//   プッシュ受信時、payload の data.badge を読んでアイコンに件数を表示
self.addEventListener('push', function(event){
  try {
    var p = {};
    try { p = event.data ? event.data.json() : {}; } catch(_) {}
    // OneSignalは独自データを custom.a に格納するため両方を確認
    var b = null;
    if (p && p.custom && p.custom.a && typeof p.custom.a.badge !== 'undefined') b = parseInt(p.custom.a.badge);
    else if (p && typeof p.badge !== 'undefined') b = parseInt(p.badge);
    if (b !== null && !isNaN(b) && self.registration && self.navigator && self.navigator.setAppBadge) {
      event.waitUntil(self.navigator.setAppBadge(b).catch(function(){}));
    } else if (self.navigator && self.navigator.setAppBadge) {
      // 件数不明でも「●」を出す
      event.waitUntil(self.navigator.setAppBadge().catch(function(){}));
    }
  } catch(e) {}
});

// 通知をタップしたらバッジを消す
self.addEventListener('notificationclick', function(event){
  try { if (self.navigator && self.navigator.clearAppBadge) event.waitUntil(self.navigator.clearAppBadge().catch(function(){})); } catch(e){}
});
