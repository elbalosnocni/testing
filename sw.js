/* Service worker: nhận thông báo đẩy (Firebase Cloud Messaging) và mở lại trang lịch khi bấm. */
importScripts('config.js');
const C = self.LLV_CONFIG || {};
if (C.FIREBASE && C.FIREBASE.projectId) {
  importScripts('https://www.gstatic.com/firebasejs/10.12.2/firebase-app-compat.js',
                'https://www.gstatic.com/firebasejs/10.12.2/firebase-messaging-compat.js');
  firebase.initializeApp(C.FIREBASE);
  firebase.messaging().onBackgroundMessage(function (p) {
    const d = p.data || {};
    return self.registration.showNotification(d.title || 'Lịch làm việc', {
      body: d.body || '', tag: d.tag || 'llv', icon: 'icon-192.png', badge: 'icon-192.png',
      data: { url: d.url || 'index.html' }
    });
  });
}
self.addEventListener('install', function () { self.skipWaiting(); });
self.addEventListener('activate', function (e) { e.waitUntil(self.clients.claim()); });
self.addEventListener('notificationclick', function (e) {
  e.notification.close();
  const url = new URL((e.notification.data && e.notification.data.url) || 'index.html', self.registration.scope).href;
  e.waitUntil(self.clients.matchAll({ type: 'window', includeUncontrolled: true }).then(function (list) {
    for (const c of list) { if ('focus' in c) { if (c.navigate) c.navigate(url); return c.focus(); } }
    return self.clients.openWindow(url);
  }));
});
