/* Cấu hình dùng chung cho index.html và sw.js (điền khi bật thông báo đẩy).
 * Lấy trong Firebase Console → Project settings → General → "Your apps" (Web app),
 * và Cloud Messaging → Web configuration → Web Push certificates → Key pair (VAPID).
 * Để trống FIREBASE/VAPID_KEY thì web vẫn chạy, chỉ tắt nút "Bật thông báo". */
self.LLV_CONFIG = {
  FIREBASE: { apiKey: '', authDomain: '', projectId: '', messagingSenderId: '', appId: '' },
  VAPID_KEY: ''
};
