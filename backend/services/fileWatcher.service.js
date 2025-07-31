/*
 * File: services/fileWatcher.service.js
 * Mô tả: Service sử dụng Chokidar để theo dõi sự thay đổi của các file trong thư mục /data
 * và tự động tải lại chúng vào bộ nhớ.
 * Đã cập nhật để sử dụng kỹ thuật debounce, giúp ổn định hơn khi file đang được ghi.
 */
const chokidar = require('chokidar');
const path = require('path');
const { reloadJsonFile, DATA_DIR } = require('./json.service');

// Sử dụng một object để quản lý các bộ đếm thời gian (timers) cho mỗi file
const debounceTimers = {};
const DEBOUNCE_DELAY = 500; // Chờ 250ms sau khi file thay đổi

const handleFileChange = (filePath) => {
    // Xóa bộ đếm thời gian cũ nếu có để reset
    if (debounceTimers[filePath]) {
        clearTimeout(debounceTimers[filePath]);
    }

    // Tạo một bộ đếm thời gian mới. Chỉ khi hết thời gian chờ, file mới được đọc.
    debounceTimers[filePath] = setTimeout(() => {
        console.log(`>>> [Watcher] File ${path.basename(filePath)} đã ổn định. Đang tải lại...`);
        reloadJsonFile(filePath);
        // Xóa timer sau khi đã thực thi
        delete debounceTimers[filePath];
    }, DEBOUNCE_DELAY);
};

const initializeWatcher = () => {
    console.log(`>>> [Watcher] Đang theo dõi các thay đổi trong thư mục: ${DATA_DIR}`);

    const watcher = chokidar.watch(DATA_DIR, {
        ignored: /(^|[\/\\])\../,
        persistent: true,
        ignoreInitial: true,
    });

    // Sử dụng chung một handler đã được debounce cho các sự kiện
    watcher
        .on('change', (filePath) => {
            console.log(`>>> [Watcher] Phát hiện thay đổi trong file: ${path.basename(filePath)}. Chờ ổn định...`);
            handleFileChange(filePath);
        })
        .on('add', (filePath) => {
            console.log(`>>> [Watcher] Phát hiện file mới: ${path.basename(filePath)}. Chờ ổn định...`);
            handleFileChange(filePath);
        })
        .on('unlink', (filePath) => {
            console.log(`>>> [Watcher] Phát hiện file bị xóa: ${path.basename(filePath)}.`);
            // Xóa thì không cần debounce, thực hiện ngay
            reloadJsonFile(filePath, { deleted: true });
        })
        .on('error', error => console.error(`>>> [Watcher] Lỗi: ${error}`));
};

module.exports = {
    initializeWatcher
};
