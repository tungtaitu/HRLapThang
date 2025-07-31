/*
 * File: config/db.js
 * Mô tả: Cấu hình và khởi tạo kết nối đến SQL Server.
 */
const sql = require('mssql');

// Đảm bảo các biến môi trường đã được load (thường là trong file server.js chính)
const dbConfig = {
    user: process.env.DB_USER,
    password: process.env.DB_PASSWORD,
    server: process.env.DB_SERVER,
    database: process.env.DB_DATABASE,
    port: parseInt(process.env.DB_PORT, 10) || 1433,
    options: {
        encrypt: false,
        trustServerCertificate: true,
        connectionTimeout: 15000,
        pool: { max: 10, min: 0, idleTimeoutMillis: 30000 }
    }
};

if (!dbConfig.user || !dbConfig.password || !dbConfig.server || !dbConfig.database) {
    console.error("!!! LỖI: Các biến môi trường của Database chưa được thiết lập.");
    process.exit(1);
}

const pool = new sql.ConnectionPool(dbConfig);
const poolConnect = pool.connect().then(p => {
    console.log('SQL Connection Pool đã được tạo thành công.');
    return p;
}).catch(err => console.error('Tạo Connection Pool thất bại:', err));

pool.on('error', err => {
    console.error('Lỗi SQL Connection Pool:', err);
});

module.exports = { pool, poolConnect };
