/*
 * File: service-worker.js
 * Mô tả: Xử lý các tác vụ nền cho PWA, bao gồm cả Push Notifications.
 */

// Lắng nghe sự kiện 'push' được gửi từ server
self.addEventListener('push', event => {
  console.log('[Service Worker] Push Received.');
  // Phân tích dữ liệu gửi về từ server (dạng JSON)
  const data = event.data.json();
  console.log('[Service Worker] Push data:', data);

  const title = data.title || 'Thông báo mới';
  const options = {
    body: data.body || 'Bạn có một tin nhắn mới.',
    icon: '/logo192.png', // Icon nhỏ hiển thị trên thông báo
    badge: '/logo72.png',  // Icon hiển thị trên thanh trạng thái Android
    tag: data.tag || 'default-tag', // Giúp gộp các thông báo có cùng tag
    data: {
        url: self.location.origin, // Lưu URL của trang web để mở lại khi click
    }
  };

  // Hiển thị thông báo
  event.waitUntil(self.registration.showNotification(title, options));
});

// Lắng nghe sự kiện người dùng click vào thông báo
self.addEventListener('notificationclick', event => {
  console.log('[Service Worker] Notification click Received.');
  
  // Đóng thông báo lại
  event.notification.close();

  // Mở lại tab ứng dụng hoặc focus vào tab đang mở
  const urlToOpen = event.notification.data.url;
  
  event.waitUntil(clients.matchAll({
    type: "window"
  }).then(clientList => {
    for (const client of clientList) {
      if (client.url === urlToOpen && 'focus' in client) {
        return client.focus();
      }
    }
    if (clients.openWindow) {
      return clients.openWindow(urlToOpen);
    }
  }));
});
