/*
 * File: server.js
 * Mô tả: File khởi chạy chính của server.
 * Chịu trách nhiệm khởi tạo Express, áp dụng các middleware,
 * và liên kết đến các routes chính của ứng dụng.
 */

// --- 1. IMPORT CÁC THƯ VIỆN CẦN THIẾT ---
require('dotenv').config();
const express = require('express');
const https = require('https');
const cors = require('cors');
const path = require('path');
const fs = require('fs').promises;
const session = require('express-session');
const FileStore = require('session-file-store')(session);
const cookieParser = require('cookie-parser');

// --- 2. IMPORT CẤU HÌNH VÀ ROUTER ---
const mainRouter = require('./routes'); // Router tổng
const { loadAllJsonData } = require('./services/json.service');
const { initializeWatcher } = require('./services/fileWatcher.service');

// --- 3. CẤU HÌNH ỨNG DỤNG ---
const app = express();
const port = process.env.PORT || 5000;

const allowedOrigins = [
    'https://nhansulapthang.io.vn',
    'http://172.22.169.126',
    'http://localhost:3000',
];
const corsOptions = {
  origin: function (origin, callback) {
    if (!origin || allowedOrigins.indexOf(origin) !== -1) {
      callback(null, true);
    } else {
      callback(new Error('Not allowed by CORS'));
    }
  },
  credentials: true,
  optionsSuccessStatus: 200
};

const sessionSecret = process.env.SESSION_SECRET;
if (!sessionSecret) {
    console.error("!!! LỖI: Biến môi trường SESSION_SECRET chưa được thiết lập.");
    process.exit(1);
}

// --- 4. SỬ DỤNG MIDDLEWARE ---
if (process.env.NODE_ENV === 'production') {
    app.set('trust proxy', 1);
}
app.use(cors(corsOptions));

// Tăng giới hạn kích thước cho dữ liệu JSON và form để cho phép upload file lớn
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ limit: '50mb', extended: true }));

app.use(cookieParser());
app.use(session({
    name: 'app.sid',
    store: new FileStore({ path: path.join(__dirname, 'sessions') }),
    secret: sessionSecret,
    resave: false,
    saveUninitialized: false,
    cookie: {
        maxAge: 180 * 24 * 60 * 60 * 1000,
        httpOnly: true,
        secure: process.env.NODE_ENV === 'production',
        sameSite: process.env.NODE_ENV === 'production' ? 'none' : 'lax'
    }
}));
app.use(express.static('public'));

// --- 5. SỬ DỤNG ROUTER ---
app.use('/api', mainRouter);

// --- 6. PHỤC VỤ FILE TĨNH VÀ ROUTING CHO FRONTEND ---
const frontendBuildPath = path.join(__dirname, '..', 'frontend');
console.log(`>>> Đường dẫn tới thư mục giao diện frontend: ${frontendBuildPath}`);
app.use(express.static(frontendBuildPath));

app.get('*', (req, res) => {
    if (!req.originalUrl.startsWith('/api')) {
        res.sendFile(path.join(frontendBuildPath, 'index.html'));
    } else {
        res.status(404).json({ message: 'API endpoint not found' });
    }
});

// --- 7. KHỞI CHẠY SERVER ---
const startServer = async () => {
    try {
        const sessionsDir = path.join(__dirname, 'sessions');
        await fs.mkdir(sessionsDir, { recursive: true });
        console.log(`Thư mục session đã sẵn sàng tại: ${sessionsDir}`);
        
        await loadAllJsonData();
        initializeWatcher();

        if (process.env.NODE_ENV === 'production') {
            console.log('>>> Chế độ Production: Đang cố gắng khởi động server HTTPS...');
            const certPath = path.join(__dirname, 'cert', 'cert.pem');
            const keyPath = path.join(__dirname, 'cert', 'key.pem');

            try {
                await fs.access(certPath);
                await fs.access(keyPath);

                const httpsOptions = {
                    key: await fs.readFile(keyPath),
                    cert: await fs.readFile(certPath)
                };

                https.createServer(httpsOptions, app).listen(port, () => {
                    console.log(`>>> ✅ Backend server đang chạy an toàn (HTTPS) tại: https://localhost:${port}`);
                });

            } catch (certError) {
                console.error('------------------------------------------------------------------');
                console.error('!!! LỖI: Không tìm thấy file certificate (cert.pem) hoặc key (key.pem).');
                console.error(`>>> Vui lòng tạo Origin Certificate từ Cloudflare và đặt file vào thư mục: ${path.join(__dirname, 'cert')}`);
                console.error('------------------------------------------------------------------');
                process.exit(1);
            }
        } else {
            app.listen(port, () => {
                console.log(`>>> Backend server (Development) đang chạy (HTTP) tại: http://localhost:${port}`);
            });
        }

    } catch (error) {
        console.error('!!! Không thể khởi động server:', error);
    }
};

startServer();
