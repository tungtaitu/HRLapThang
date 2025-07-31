/*
 * File: server.js
 * Mô tả: API Backend cho ứng dụng chấm công và lương (Phiên bản đơn giản hóa).
 * Tác giả: Gemini
 * Ngày tạo: 06/06/2025
 * Cập nhật lần cuối: 27/06/2025 - Thêm sắp xếp theo bộ phận cho các chức năng xuất file.
 */

// --- 1. IMPORT CÁC THƯ VIỆN CẦN THIẾT ---
const express = require('express');
const sql = require('mssql');
const cors = require('cors');
const path = require('path');
const fs = require('fs').promises;
const multer = require('multer'); // Thư viện xử lý file upload
const xlsx = require('xlsx');     // Thư viện đọc file Excel
const session = require('express-session');
const FileStore = require('session-file-store')(session); // Thêm thư viện để lưu session vào file
const cookieParser = require('cookie-parser');
const { v4: uuidv4 } = require('uuid');
const { 
    generateRegistrationOptions, 
    verifyRegistrationResponse,
    generateAuthenticationOptions,
    verifyAuthenticationResponse,
} = require('@simplewebauthn/server');

// --- 2. CẤU HÌNH ỨNG DỤNG ---
const app = express();
const port = 5000;
const LEAVE_DATA_FILE = path.join(__dirname, 'leave_data.json');
const PAYROLL_APPROVAL_FILE = path.join(__dirname, 'payroll_approvals.json');
const USER_PASSWORDS_FILE = path.join(__dirname, 'user_passwords.json'); // File lưu mật khẩu người dùng
const ADMIN_SUBMITTED_LEAVE_FILE = path.join(__dirname, 'admin_submitted_leave.json');
const upload = multer({ dest: 'uploads/' });
const SESSIONS_DIR = path.join(__dirname, 'sessions'); 

const rpID = 'nhansulapthang.io.vn';
const rpName = 'Hệ thống Nhân sự Lập Thắng';
const origin = `https://${rpID}`;

let userAuthenticators = [];
const AUTHENTICATORS_FILE = path.join(__dirname, 'authenticators.json');

const readAuthenticators = async () => {
    try {
        await fs.access(AUTHENTICATORS_FILE);
        const data = await fs.readFile(AUTHENTICATORS_FILE, 'utf-8');
        userAuthenticators = JSON.parse(data);
    } catch (error) {
        userAuthenticators = [];
    }
};
const writeAuthenticators = async () => {
    await fs.writeFile(AUTHENTICATORS_FILE, JSON.stringify(userAuthenticators, null, 2));
};
readAuthenticators(); 
const dbConfig = {
    user: 'sa',
    password: 'MIS@dbbox',
    server: '172.22.166.33',
    database: 'MISDBEMP',
    port: 1433,
    options: {
        encrypt: false,
        trustServerCertificate: true,
        connectionTimeout: 15000,
        pool: {
            max: 10,
            min: 0,
            idleTimeoutMillis: 30000
        }
    }
};
const allowedOrigins = [
    'https://nhansulapthang.io.vn', 
    'http://localhost:3000',       
    'http://172.22.169.126'        
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
const pool = new sql.ConnectionPool(dbConfig);
const poolConnect = pool.connect().then(p => {
    console.log('SQL Connection Pool đã được tạo thành công.');
    return p;
}).catch(err => console.error('Tạo Connection Pool thất bại:', err));

pool.on('error', err => {
    console.error('Lỗi SQL Connection Pool:', err);
});


const readJsonFile = async (filePath) => {
    try {
        await fs.access(filePath);
        const data = await fs.readFile(filePath, 'utf-8');
        return JSON.parse(data);
    } catch (error) {
        return []; 
    }
};

const writeJsonFile = async (filePath, data) => {
    await fs.writeFile(filePath, JSON.stringify(data, null, 2));
};

async function getLeaveSummary(connectionPool, userId, year) {
    const selectedYear = parseInt(year);
    const today = new Date();
    const currentYear = today.getFullYear();
    let summary = { total: 0, used: 0, remaining: 0, isCurrentYear: false };

    let employeeConfig = null;
    try {
        const jsonData = await readJsonFile(LEAVE_DATA_FILE);
        employeeConfig = jsonData.find(emp => emp.MSNV === userId);
    } catch (fileError) {
        employeeConfig = null;
    }

    if (!employeeConfig) {
        summary.remaining = 0;
        summary.total = 0;
        summary.isCurrentYear = (selectedYear === currentYear);
        try {
            const totalUsedInYearResult = await connectionPool.request()
                .input('userid_param', sql.VarChar, userId)
                .input('year_param', sql.Int, year)
                .query`SELECT SUM(ISNULL(HHour, 0)) as TotalUsed FROM EMPHOLIDAY WHERE empid = @userid_param AND YEAR(DateUP) = @year_param AND JiaType = 'E'`;
            summary.used = totalUsedInYearResult.recordset[0]?.TotalUsed || 0;
        } catch (dbError) {
            console.error(`Lỗi khi truy vấn số phép đã dùng cho NV ${userId} (không có trong file config):`, dbError);
            summary.used = 0;
        }
        return summary;
    }

    if (selectedYear === currentYear) {
        summary.isCurrentYear = true;

        const configMonth = parseInt(employeeConfig['Month']) || 1;
        const carriedOverHours = parseFloat(employeeConfig['PHEP']) || 0;
        const currentMonth = today.getMonth() + 1; // 1-12
        let entitledThisYear = 0;
        let entitledMonthsCount = 0;
        
        if (configMonth === 12) {
             entitledMonthsCount = currentMonth;
        } else {
            if(currentMonth > configMonth) {
                entitledMonthsCount = currentMonth - configMonth;
            }
        }
        entitledThisYear = entitledMonthsCount * 8;
        
        const firstUsageMonthToConsider = (configMonth === 12) ? 1 : configMonth + 1;

        const usedLeaveSinceEntitlementResult = await connectionPool.request()
            .input('userid_param', sql.VarChar, userId)
            .input('year_param', sql.Int, year)
            .input('start_month_param', sql.Int, firstUsageMonthToConsider)
            .query`
                SELECT SUM(ISNULL(HHour, 0)) as TotalUsed
                FROM EMPHOLIDAY
                WHERE empid = @userid_param
                  AND YEAR(DateUP) = @year_param
                  AND MONTH(DateUP) >= @start_month_param
                  AND JiaType = 'E'
            `;
        const usedHoursSinceEntitlement = usedLeaveSinceEntitlementResult.recordset[0]?.TotalUsed || 0;

        const totalUsedInYearResult = await connectionPool.request()
            .input('userid_param', sql.VarChar, userId)
            .input('year_param', sql.Int, year)
            .query`SELECT SUM(ISNULL(HHour, 0)) as TotalUsed FROM EMPHOLIDAY WHERE empid = @userid_param AND YEAR(DateUP) = @year_param AND JiaType = 'E'`;
        const totalUsedForDisplay = totalUsedInYearResult.recordset[0]?.TotalUsed || 0;

        summary.total = carriedOverHours + entitledThisYear;
        summary.used = totalUsedForDisplay; 
        summary.remaining = summary.total - usedHoursSinceEntitlement;

    } else { 
        const totalUsedInYearResult = await connectionPool.request()
            .input('userid_param', sql.VarChar, userId)
            .input('year_param', sql.Int, year)
            .query`SELECT SUM(ISNULL(HHour, 0)) as TotalUsed FROM EMPHOLIDAY WHERE empid = @userid_param AND YEAR(DateUP) = @year_param AND JiaType = 'E'`;

        summary.used = totalUsedInYearResult.recordset[0]?.TotalUsed || 0;
        summary.total = 0;
        summary.remaining = 0;
    }

    return summary;
}
function calculateWorkDuration(indat) {
    if (!indat) return 'Không rõ';
    const startDate = new Date(indat);
    const endDate = new Date();
    let years = endDate.getFullYear() - startDate.getFullYear();
    let months = endDate.getMonth() - startDate.getMonth();
    let days = endDate.getDate() - startDate.getDate();
    if (days < 0) {
        months--;
        const prevMonthLastDay = new Date(endDate.getFullYear(), endDate.getMonth(), 0).getDate();
        days += prevMonthLastDay;
    }
    if (months < 0) {
        years--;
        months += 12;
    }
    return `${years} năm ${months} tháng ${days} ngày`;
}
function calculateLeaveHours(startTimeStr, endTimeStr) {
    if (!startTimeStr || !endTimeStr || startTimeStr.length < 4 || endTimeStr.length < 4) return 0;

    const startH = parseInt(startTimeStr.substring(0, 2));
    const startM = parseInt(startTimeStr.substring(2, 4));
    const endH = parseInt(endTimeStr.substring(0, 2));
    const endM = parseInt(endTimeStr.substring(2, 4));

    if (isNaN(startH) || isNaN(startM) || isNaN(endH) || isNaN(endM)) return 0;
    
    const start = new Date(1970, 0, 1, startH, startM);
    const end = new Date(1970, 0, 1, endH, endM);
    if (end <= start) return 0;

    const morningStart = new Date(1970, 0, 1, 8, 0);
    const morningEnd = new Date(1970, 0, 1, 12, 0);
    const afternoonStart = new Date(1970, 0, 1, 13, 0);
    const afternoonEnd = new Date(1970, 0, 1, 17, 0);
    let totalMs = 0;
    
    const morningOverlapStart = Math.max(start, morningStart);
    const morningOverlapEnd = Math.min(end, morningEnd);
    if (morningOverlapEnd > morningOverlapStart) totalMs += morningOverlapEnd - morningOverlapStart;

    const afternoonOverlapStart = Math.max(start, afternoonStart);
    const afternoonOverlapEnd = Math.min(end, afternoonEnd);
    if (afternoonOverlapEnd > afternoonOverlapStart) totalMs += afternoonOverlapEnd - afternoonOverlapStart;

    return Math.round((totalMs / 3600000) * 10) / 10;
}

const parseDate = (ddmmyyyy) => {
    if (!ddmmyyyy || ddmmyyyy.length !== 8) return null;
    const day = parseInt(ddmmyyyy.substring(0, 2));
    const month = parseInt(ddmmyyyy.substring(2, 4)) - 1; 
    const year = parseInt(ddmmyyyy.substring(4, 8));
    const date = new Date(Date.UTC(year, month, day));
    if (isNaN(date.getTime()) || date.getUTCFullYear() !== year || date.getUTCMonth() !== month || date.getUTCDate() !== day) {
        return null;
    }
    return date;
};

// --- 3. SỬ DỤNG MIDDLEWARE ---
app.use(cors(corsOptions));
app.use(express.json());
app.use(cookieParser()); 
const fileStoreOptions = {
    path: "./sessions", 
    ttl: 86400 * 30,    
    reapInterval: 86400 
};
app.use(session({
    name: 'app.sid', 
    store: new FileStore(fileStoreOptions),
    secret: process.env.SESSION_SECRET || 'a-very-strong-secret-key-for-session-is-important', 
    resave: false,
    saveUninitialized: false,
    cookie: {
        maxAge: 30 * 24 * 60 * 60 * 1000, 
        httpOnly: true,
        secure: process.env.NODE_ENV === 'production',
        sameSite: 'lax'
    }
}));
app.use(express.static('public')); 

const isAdmin = (req, res, next) => {
    if (req.session && req.session.user && req.session.user.isAdmin) {
        return next();
    }
    return res.status(401).json({ message: 'Chưa đăng nhập hoặc không có quyền. Vui lòng đăng nhập lại.' });
};

// --- API ENDPOINTS ---
app.post('/api/login', async (req, res) => {
    const { empid, password } = req.body; 
    if (!empid || !password) {
        return res.status(400).json({ message: 'Tên đăng nhập và Mật khẩu không được để trống.' });
    }
    
    try {
        await poolConnect;

        const adminResult = await pool.request()
            .input('admin_userid', sql.VarChar, empid)
            .input('admin_password', sql.VarChar, password)
            .query`SELECT MUSER, USERNAME FROM SYSUSER WHERE MUSER = @admin_userid AND PASSWORD = @admin_password`;
        
        if (adminResult.recordset.length > 0) {
            const adminUser = adminResult.recordset[0];
            const userSessionData = { id: adminUser.MUSER, name: adminUser.USERNAME || 'Administrator', isAdmin: true };
            req.session.user = userSessionData;
            
            req.session.save((err) => {
                if (err) {
                    console.error("Lỗi khi lưu session (admin):", err);
                    return res.status(500).json({ message: 'Lỗi server khi đăng nhập.' });
                }
                return res.status(200).json(userSessionData);
            });
            return;
        }

        let userResult = null;
        const customPasswords = await readJsonFile(USER_PASSWORDS_FILE);
        const customUserRecord = customPasswords.find(u => u.empid === empid);

        if (customUserRecord) {
            if (customUserRecord.password === password) {
                const employeeInfo = await pool.request().input('empid', sql.VarChar, empid).query`SELECT TOP 1 EMPID as id, EMPNAM_VN as name, indat FROM EMPFILE WHERE EMPID = @empid`;
                userResult = employeeInfo.recordset;
            }
        } else {
            const defaultPasswordResult = await pool.request()
                .input('empid_param', sql.VarChar, empid)
                .input('password_param', sql.VarChar, password)
                .query`
                    SELECT TOP 1 EMPID as id, EMPNAM_VN as name, indat 
                    FROM EMPFILE 
                    WHERE EMPID = @empid_param 
                    AND (RIGHT('00' + CAST(BDD AS VARCHAR(2)), 2) + RIGHT('00' + CAST(BMM AS VARCHAR(2)), 2) + CAST(BYY AS VARCHAR(4))) = @password_param
                `;
            userResult = defaultPasswordResult.recordset;
        }

        if (userResult && userResult.length > 0) {
            const user = userResult[0];
            const workDuration = calculateWorkDuration(user.indat);
            const userSessionData = { ...user, isAdmin: false, workDuration };
            req.session.user = userSessionData;

            req.session.save((err) => {
                if (err) {
                    console.error("Lỗi khi lưu session (nhân viên):", err);
                    return res.status(500).json({ message: 'Lỗi server khi đăng nhập.' });
                }
                res.status(200).json(userSessionData);
            });
        } else {
            res.status(401).json({ message: 'Mã nhân viên hoặc Mật khẩu không chính xác.' });
        }
    } catch (err) {
        console.error("Lỗi khi đăng nhập:", err);
        res.status(500).json({ message: 'Lỗi server khi đăng nhập.' });
    }
});
app.get('/api/check-session', (req, res) => {
    if (req.session && req.session.user) {
        res.status(200).json(req.session.user);
    } else {
        res.status(401).json({ message: 'Chưa đăng nhập.' });
    }
});

app.post('/api/logout', (req, res) => {
    req.session.destroy(err => {
        if (err) {
            return res.status(500).json({ message: 'Đăng xuất thất bại.' });
        }
        res.clearCookie('connect.sid');
        res.status(200).json({ message: 'Đăng xuất thành công.' });
    });
});

app.post('/api/webauthn/generate-registration-options', async (req, res) => {
    const user = req.session.user;
    if (!user) return res.status(401).json({ message: 'Bạn cần đăng nhập để thực hiện.' });
    
    const userCurrentAuthenticators = userAuthenticators.filter(auth => auth.userID === user.id);

    const options = await generateRegistrationOptions({
        rpName,
        rpID,
        userID: user.id,
        userName: user.name,
        excludeCredentials: userCurrentAuthenticators.map(auth => ({
            id: Buffer.from(auth.credentialID, 'base64url'),
            type: 'public-key',
            transports: auth.transports,
        })),
        authenticatorSelection: { userVerification: 'preferred', residentKey: 'required' },
    });
    
    req.session.challenge = options.challenge;
    res.status(200).json(options);
});

app.post('/api/webauthn/verify-registration', async (req, res) => {
    const { body } = req;
    const user = req.session.user;
    const challenge = req.session.challenge;
    
    if (!user || !challenge) return res.status(400).json({ message: 'Session không hợp lệ.' });
    
    try {
        const verification = await verifyRegistrationResponse({
            response: body,
            expectedChallenge: challenge,
            expectedOrigin: origin,
            expectedRPID: rpID,
            requireUserVerification: true,
        });

        const { verified, registrationInfo } = verification;
    
        if (verified && registrationInfo) {
            const { credentialPublicKey, credentialID, counter } = registrationInfo;
            
            userAuthenticators.push({
                userID: user.id,
                credentialID: Buffer.from(credentialID).toString('base64url'),
                credentialPublicKey: Buffer.from(credentialPublicKey).toString('base64url'),
                counter,
                transports: body.response.transports || [],
            });
            await writeAuthenticators(); 
    
            delete req.session.challenge;
            return res.status(200).json({ verified: true });
        }
    } catch (error) {
        console.error("Lỗi Verify Registration:", error);
        return res.status(400).json({ error: error.message });
    }
    return res.status(400).json({ verified: false, message: 'Xác thực đăng ký thất bại.' });
});

app.get('/api/webauthn/generate-authentication-options', async (req, res) => {
    const options = await generateAuthenticationOptions({
        rpID,
        allowCredentials: userAuthenticators.map(auth => ({
            id: Buffer.from(auth.credentialID, 'base64url'),
            type: 'public-key',
            transports: auth.transports,
        })),
        userVerification: 'preferred',
    });

    req.session.challenge = options.challenge;
    res.status(200).json(options);
});

app.post('/api/webauthn/verify-authentication', async (req, res) => {
    const { body } = req;
    const challenge = req.session.challenge;
    if (!challenge) return res.status(400).json({ message: 'Session không hợp lệ.' });

    const requestCredentialID = Buffer.from(body.rawId, 'base64url').toString('base64url');
    const authenticator = userAuthenticators.find(auth => auth.credentialID === requestCredentialID);

    if (!authenticator) return res.status(404).json({ message: 'Thiết bị chưa được đăng ký.' });

    try {
        const verification = await verifyAuthenticationResponse({
            response: body,
            expectedChallenge: challenge,
            expectedOrigin: origin,
            expectedRPID: rpID,
            authenticator: {
                credentialID: Buffer.from(authenticator.credentialID, 'base64url'),
                credentialPublicKey: Buffer.from(authenticator.credentialPublicKey, 'base64url'),
                counter: authenticator.counter,
                transports: authenticator.transports,
            },
            requireUserVerification: true,
        });

        const { verified, authenticationInfo } = verification;
    
        if (verified) {
            authenticator.counter = authenticationInfo.newCounter;
            await writeAuthenticators(); 

            await poolConnect;
            const userResult = await pool.request().input('empid', sql.NVarChar, authenticator.userID).query`SELECT TOP 1 EMPID as id, EMPNAM_VN as name, indat FROM EMPFILE WHERE EMPID = @empid`;
            
            if (userResult.recordset.length > 0) {
                const user = userResult.recordset[0];
                const workDuration = calculateWorkDuration(user.indat);
                const userSessionData = { ...user, isAdmin: false, workDuration };
                
                req.session.user = userSessionData;
                delete req.session.challenge;
                
                return res.status(200).json(userSessionData);
            }
        }
    } catch (error) {
        console.error("Lỗi Verify Authentication:", error);
        return res.status(400).json({ error: error.message });
    }
    
    return res.status(400).json({ verified: false, message: 'Xác thực đăng nhập thất bại.' });
});
app.post('/api/user/change-password', async (req, res) => {
    const { userId, oldPassword, newPassword } = req.body;
    if (!userId || !oldPassword || !newPassword) {
        return res.status(400).json({ message: 'Vui lòng nhập đầy đủ thông tin.' });
    }
    
    try {
        await poolConnect; 
        
        let isValidOldPassword = false;
        const customPasswords = await readJsonFile(USER_PASSWORDS_FILE);
        const customUser = customPasswords.find(u => u.empid === userId && u.password === oldPassword);
        
        if (customUser) {
            isValidOldPassword = true;
        } else {
            const defaultPasswordResult = await pool.request()
                .input('empid_param', sql.VarChar, userId)
                .input('password_param', sql.VarChar, oldPassword)
                .query`SELECT EMPID FROM EMPFILE WHERE EMPID = @empid_param AND (RIGHT('00' + CAST(BDD AS VARCHAR(2)), 2) + RIGHT('00' + CAST(BMM AS VARCHAR(2)), 2) + CAST(BYY AS VARCHAR(4))) = @password_param`;
            if (defaultPasswordResult.recordset.length > 0) {
                isValidOldPassword = true;
            }
        }

        if (!isValidOldPassword) {
            return res.status(401).json({ message: 'Mật khẩu cũ không chính xác.' });
        }

        let userFound = false;
        const updatedPasswords = customPasswords.map(u => {
            if (u.empid === userId) {
                userFound = true;
                return { ...u, password: newPassword };
            }
            return u;
        });

        if (!userFound) {
            updatedPasswords.push({ empid: userId, password: newPassword });
        }

        await writeJsonFile(USER_PASSWORDS_FILE, updatedPasswords);
        res.status(200).json({ message: 'Đổi mật khẩu thành công!' });

    } catch(err) {
        console.error("Lỗi khi đổi mật khẩu:", err);
        res.status(500).json({ message: 'Lỗi server khi đổi mật khẩu.' });
    }
});

app.post('/api/admin/upload-leave', isAdmin, upload.single('leaveFile'), async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ message: 'Không có file nào được upload.' });
        }

        const workbook = xlsx.readFile(req.file.path);
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const newDataRaw = xlsx.utils.sheet_to_json(sheet);
        const newData = newDataRaw.map(row => {
            const standardizedRow = {};
            let msnvValue = null;
            for (const key in row) {
                const trimmedKey = key.trim().toUpperCase();
                if (trimmedKey === 'MSNV') {
                    standardizedRow.MSNV = String(row[key]);
                    msnvValue = standardizedRow.MSNV;
                } else if (trimmedKey === 'MONTH') {
                    standardizedRow.Month = row[key];
                } else if (trimmedKey === 'PHEP') {
                    const phepValue = row[key];
                    standardizedRow.PHEP = (phepValue === '-' || phepValue === null || phepValue === undefined) ? 0 : parseFloat(phepValue) || 0;
                }
            }
            if (msnvValue) {
                return standardizedRow;
            }
            return null;
        }).filter(Boolean);

        if (newData.length === 0) {
             await fs.unlink(req.file.path);
             return res.status(400).json({ message: 'File Excel không có dữ liệu hợp lệ hoặc thiếu cột MSNV.' });
        }

        const existingData = await readJsonFile(LEAVE_DATA_FILE);
        
        const dataMap = new Map(existingData.map(item => [item.MSNV, item]));
        newData.forEach(newItem => {
            dataMap.set(newItem.MSNV, newItem);
        });
        const mergedData = Array.from(dataMap.values());

        await writeJsonFile(LEAVE_DATA_FILE, mergedData);
        
        await fs.unlink(req.file.path);

        res.status(200).json({ message: `Cập nhật thành công! ${newData.length} bản ghi đã được xử lý.` });

    } catch (error) {
        console.error("Lỗi khi upload file:", error);
        if (req.file && req.file.path) {
            try {
                await fs.unlink(req.file.path);
            } catch (unlinkError) {
                console.error("Lỗi khi xóa file tạm:", unlinkError);
            }
        }
        res.status(500).json({ message: 'Có lỗi xảy ra khi xử lý file.' });
    }
});
// ====================== BẮT ĐẦU THAY ĐỔI ======================
app.get('/api/admin/export-leave/:year', isAdmin, async (req, res) => {
    const { year } = req.params;
    const { groupId } = req.query; 

    try {
        await poolConnect;
        let employeeQuery = `
            SELECT
                f.EMPID,
                f.EMPNAM_VN,
                ISNULL(latest_sal.GROUPID, 'N/A') as GROUPID,
                ISNULL(g.SYS_VALUE, 'Chưa phân loại') as GroupName
            FROM
                EMPFILE f
            OUTER APPLY (
                SELECT TOP 1 s.GROUPID
                FROM EMPDSALARY s
                WHERE s.EMPID = f.EMPID
                ORDER BY s.YYMM DESC
            ) latest_sal
            LEFT JOIN BASICCODE g ON latest_sal.GROUPID = g.SYS_TYPE AND g.FUNC = 'GROUPID'
            WHERE (f.STATUS IS NULL OR f.STATUS != 'Q') AND f.OUTDAT IS NULL
        `;

        if (groupId && groupId !== 'ALL') {
            employeeQuery += ` AND latest_sal.GROUPID = @groupId`;
        }
        employeeQuery += ` ORDER BY GROUPID, f.EMPID`;

        const request = pool.request();
        if (groupId && groupId !== 'ALL') {
            request.input('groupId', sql.VarChar, groupId);
        }

        const employeesResult = await request.query(employeeQuery);
        const employees = employeesResult.recordset;

        const leaveDataForExport = [];
        for (const emp of employees) {
            const summary = await getLeaveSummary(pool, emp.EMPID, year);
            leaveDataForExport.push({
                'Mã Bộ Phận': emp.GROUPID,
                'Tên Bộ Phận': emp.GroupName,
                'Mã Nhân Viên': emp.EMPID,
                'Tên Nhân Viên': emp.EMPNAM_VN,
                'Số Giờ Phép Năm Còn Lại': summary.remaining
            });
        }

        const worksheet = xlsx.utils.json_to_sheet(leaveDataForExport);
        worksheet['!cols'] = [ { wch: 15 }, { wch: 25 }, { wch: 15 }, { wch: 30 }, { wch: 30 } ];
        const workbook = xlsx.utils.book_new();
        xlsx.utils.book_append_sheet(workbook, worksheet, `PhepNam${year}`);
        
        const buffer = xlsx.write(workbook, { bookType: 'xlsx', type: 'buffer' });
        res.setHeader('Content-Disposition', `attachment; filename=TongHopPhepNam_${year}${groupId && groupId !== 'ALL' ? '_' + groupId : ''}.xlsx`);
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.send(buffer);

    } catch (err) {
        console.error("Lỗi khi xuất file Excel phép năm:", err);
        res.status(500).json({ message: 'Lỗi server khi xuất file.' });
    }
});

app.get('/api/holidays/:userId/:year', async (req, res) => {
    const { userId, year } = req.params;
    try {
        await poolConnect;
        const employeeResult = await pool.request()
            .input('userid_param', sql.VarChar, userId)
            .query`SELECT EMPNAM_VN FROM EMPFILE WHERE EMPID = @userid_param`;

        if (employeeResult.recordset.length === 0) {
            return res.status(404).json({ message: 'Không tìm thấy nhân viên.' });
        }
        const employeeName = employeeResult.recordset[0].EMPNAM_VN;
        const summary = await getLeaveSummary(pool, userId, year);
        const holidayListResult = await pool.request()
            .input('userid_param', sql.VarChar, userId)
            .input('year_param', sql.Int, year)
            .query`SELECT DateUP, JiaType, HHour FROM EMPHOLIDAY WHERE empid = @userid_param AND YEAR(DateUP) = @year_param ORDER BY DateUP ASC`;

        const jiaTypeMap = { 'A': 'Việc riêng', 'B': 'Phép bệnh', 'E': 'Phép năm', 'D': 'Phép tang' };
        const formattedHolidayList = holidayListResult.recordset.map(row => {
            const jiaTypeCode = row.JiaType ? row.JiaType.trim() : '';
            return {
                date: row.DateUP,
                reason: jiaTypeMap[jiaTypeCode] || jiaTypeCode,
                hours: row.HHour || 0
            };
        });
        
        res.status(200).json({
            employeeName: employeeName,
            holidayList: formattedHolidayList,
            summary: summary
        });

    } catch (err) {
        console.error("Lỗi API Holidays:", err);
        res.status(500).json({ message: 'Lỗi server khi lấy dữ liệu ngày nghỉ.' });
    }
});
app.get('/api/admin/all-payrolls/:yearMonth', isAdmin, async (req, res) => {
    const { yearMonth } = req.params;
    const { groupId } = req.query;
    const yearMonthFormatted = yearMonth.replace('-', '');
    try {
        await poolConnect;
        let query = `
            SELECT 
                f.EMPID, 
                f.EMPNAM_VN, 
                s.REAL_TOTAL,
                s.GROUPID,
                g.SYS_VALUE as GroupName
            FROM EMPFILE f
            JOIN EMPDSALARY s ON f.EMPID = s.EMPID
            LEFT JOIN BASICCODE g ON s.GROUPID = g.SYS_TYPE AND g.FUNC = 'GROUPID'
            WHERE s.YYMM = @yymm_param
        `;

        if (groupId && groupId !== 'ALL') {
            query += ` AND s.GROUPID = @groupId`;
        }
        query += ` ORDER BY s.GROUPID, f.EMPID`;

        const request = pool.request();
        request.input('yymm_param', sql.VarChar, yearMonthFormatted);
        if (groupId && groupId !== 'ALL') {
            request.input('groupId', sql.VarChar, groupId);
        }

        const result = await request.query(query);
        const approvals = await readJsonFile(PAYROLL_APPROVAL_FILE);
        const isApproved = approvals.includes(yearMonth);
        res.status(200).json({ payrolls: result.recordset, isApproved });
    } catch (err) {
        console.error("Lỗi API All Payrolls:", err);
        res.status(500).json({ message: 'Lỗi server khi lấy dữ liệu lương.' });
    }
});

app.post('/api/admin/approve-payroll', isAdmin, async (req, res) => {
    const { yearMonth } = req.body;
    try {
        const approvals = await readJsonFile(PAYROLL_APPROVAL_FILE);
        if (!approvals.includes(yearMonth)) {
            approvals.push(yearMonth);
            await writeJsonFile(PAYROLL_APPROVAL_FILE, approvals);
        }
        res.status(200).json({ message: `Đã phê duyệt thành công lương tháng ${yearMonth}.` });

    } catch (err) {
        console.error("Lỗi API Approve Payroll:", err);
        res.status(500).json({ message: 'Lỗi server khi phê duyệt.' });
    }
});

// API MỚI: Xuất file bảng lương theo bộ phận
app.get('/api/admin/export-payrolls/:yearMonth', isAdmin, async (req, res) => {
    const { yearMonth } = req.params;
    const { groupId } = req.query;
    const yearMonthFormatted = yearMonth.replace('-', '');

    try {
        await poolConnect;

        let query = `
            SELECT 
                s.GROUPID as 'Mã Bộ Phận',
                g.SYS_VALUE as 'Tên Bộ Phận',
                f.EMPID as 'Mã Nhân Viên',
                f.EMPNAM_VN as 'Tên Nhân Viên',
                s.REAL_TOTAL as 'Lương Thực Lãnh'
            FROM EMPFILE f
            JOIN EMPDSALARY s ON f.EMPID = s.EMPID
            LEFT JOIN BASICCODE g ON s.GROUPID = g.SYS_TYPE AND g.FUNC = 'GROUPID'
            WHERE s.YYMM = @yymm_param
        `;

        if (groupId && groupId !== 'ALL') {
            query += ` AND s.GROUPID = @groupId`;
        }
        
        query += ` ORDER BY s.GROUPID, f.EMPID`;
        
        const request = pool.request();
        request.input('yymm_param', sql.VarChar, yearMonthFormatted);
        if (groupId && groupId !== 'ALL') {
            request.input('groupId', sql.VarChar, groupId);
        }

        const result = await request.query(query);

        const worksheet = xlsx.utils.json_to_sheet(result.recordset);
        worksheet['!cols'] = [ { wch: 15 }, { wch: 25 }, { wch: 15 }, { wch: 30 }, { wch: 20 } ];
        const workbook = xlsx.utils.book_new();
        xlsx.utils.book_append_sheet(workbook, worksheet, `BangLuong_${yearMonth}`);
        
        const buffer = xlsx.write(workbook, { bookType: 'xlsx', type: 'buffer' });
        res.setHeader('Content-Disposition', `attachment; filename=BangLuong_${yearMonth}${groupId && groupId !== 'ALL' ? '_' + groupId : ''}.xlsx`);
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.send(buffer);

    } catch (err) {
        console.error("Lỗi khi xuất file bảng lương:", err);
        res.status(500).json({ message: 'Lỗi server khi xuất file bảng lương.' });
    }
});
// ====================== KẾT THÚC THAY ĐỔI ======================
app.get('/api/payroll/:userId/:yearMonth', async (req, res) => {
    const { userId, yearMonth } = req.params;
    const approvalStartDate = new Date('2025-03-01');
    const selectedDate = new Date(`${yearMonth}-01`);
    const yearMonthFormatted = yearMonth.replace('-', '');

    if (selectedDate >= approvalStartDate) {
        const approvals = await readJsonFile(PAYROLL_APPROVAL_FILE);
        if (!approvals.includes(yearMonth)) {
            return res.status(200).json({ approved: false, message: 'Phiếu lương cho tháng này chưa được phê duyệt.' });
        }
    }
    try {
        await poolConnect;
        
        const result = await pool.request()
            .input('userid_param', sql.VarChar, userId)
            .input('yymm_param', sql.VarChar, yearMonthFormatted)
            .query`
                SELECT TOP 1
                    f.EMPID, f.EMPNAM_VN, 
                    s.*,
                    job_code.SYS_VALUE as ChucVu, 
                    group_code.SYS_VALUE as DonVi,
                    d.BHXH5, d.BHYT1, d.BHTN1
                FROM 
                    EMPFILE f
                OUTER APPLY (
                    SELECT TOP 1 BB, CV, PHU, NN, KT, MT ,TTKH, TNKH, QC, GT, KTAXM, MONEY_H, REAL_TOTAL, workdays, JOB, GROUPID, H1, H1M, H2, H2M, H3, H3M, B3, B3M, B4, B4M, B5, B5M, JX, BZKM, QITA, SOLE
                    FROM EMPDSALARY
                    WHERE EMPID = f.EMPID AND YYMM = @yymm_param
                ) s
                OUTER APPLY (
                    SELECT TOP 1 BHXH5, BHYT1, BHTN1
                    FROM EMPBHGT
                    WHERE EMPID = f.EMPID AND YYMM = @yymm_param
                ) d
                OUTER APPLY (
                    SELECT TOP 1 SYS_VALUE FROM BASICCODE WHERE SYS_TYPE = s.JOB
                ) job_code
                OUTER APPLY (
                    SELECT TOP 1 SYS_VALUE FROM BASICCODE WHERE SYS_TYPE = s.GROUPID
                ) group_code
                WHERE f.EMPID = @userid_param
            `;

        if (result.recordset.length > 0) {
            const data = result.recordset[0];
            
            const payrollDetails = {
                employeeInfo: {
                    soThe: data.EMPID, hoTen: data.EMPNAM_VN, chucVu: data.ChucVu || data.JOB || 'N/A',
                    donVi: data.DonVi || data.GROUPID || 'N/A',  nam: yearMonth.substring(0, 4), thang: yearMonth.substring(5, 7)
                },
                earnings: [
                    { label: 'LƯƠNG CB (BB)', value: data.BB || 0 }, { label: 'CHỨC VỤ (CV)', value: data.CV || 0 },
                    { label: 'ĐIỆN THOẠI', value: data.PHU || 0 }, { label: 'XĂNG XE', value: data.NN || 0 },
                    { label: 'KỸ THUẬT', value: data.KT || 0 }, { label: 'MÔI TRƯỜNG', value: data.MT || 0 },
                    { label: 'NHÀ Ở', value: data.TTKH || 0 }, { label: 'CHUYÊN CẦN', value: data.QC || 0 }
                ],
                deductions: [
                    { label: 'Trừ phép thường', value: data.BZKM || 0 }, { label: 'BHXH', value: data.BHXH5 || 0 },
                    { label: 'BHYT', value: data.BHYT1 || 0 }, { label: 'BHTN', value: data.BHTN1 || 0 },
                    { label: 'Phí công đoàn', value: data.GT || 0 }, { label: 'Trừ tiền khác', value: data.QITA || 0 },
                    { label: 'Thuế TN Cá Nhân', value: data.KTAXM || 0 }
                ],
                overtimeAndBonus: [
                    { label: 'Phụ cấp 0.5', hours: data.B4 || 0, amount: data.B4M || 0 },
                    { label: 'Phụ cấp 0.3', hours: data.B5 || 0, amount: data.B5M || 0 },
                    { label: 'Phụ Cấp loại A', hours: data.H1 || 0, amount: data.H1M || 0 },
                    { label: 'Phụ cấp loại B', hours: data.H2 || 0, amount: data.H2M || 0 },
                    { label: 'Phụ cấp loại C', hours: data.H3 || 0, amount: data.H3M || 0 },
                    { label: 'Phụ cấp loại D', hours: data.B3 || 0, amount: data.B3M || 0 },
                    { label: 'Phụ cấp loại Z', hours: '-' , amount: data.JX || 0 },
                    { label: 'Thu nhập khác' , hours: '-' , amount: data.TNKH || 0 }
                ],
                summary: {
                    tinhLuongMoiGio: data.MONEY_H || 0, tongSoNgayLam: data.workdays || 0,
                    soLe: data.SOLE || 0, luongThucLanh: data.REAL_TOTAL || 0
                }
            };
            res.status(200).json(payrollDetails);
        } else {
            res.status(200).json(null);
        }
    } catch (err) {
        console.error("Lỗi khi lấy dữ liệu lương của nhân viên:", err);
        res.status(500).json({ message: 'Lỗi server khi lấy dữ liệu lương.' });
    }
});

app.get('/api/timesheet/:userId/:yearMonth', async (req, res) => {
    const { userId, yearMonth } = req.params; 
    const yearMonthFormatted = yearMonth.replace('-', '');
    const year = parseInt(yearMonth.substring(0, 4));
    const month = parseInt(yearMonth.substring(5, 7));

    try {
        await poolConnect;
        
        const result = await pool.request()
            .input('userid_param', sql.VarChar, userId)
            .input('yymm_param', sql.VarChar, yearMonthFormatted)
            .input('year_param', sql.Int, year)
            .input('month_param', sql.Int, month)
            .query`
                WITH WorkDays AS (
                    SELECT workdat, timeup, timedown, TOTH, H1, H2, H3, B3, B4
                    FROM EMPWORK
                    WHERE EmpID = @userid_param AND LEFT(workdat, 6) = @yymm_param
                ),
                HolidayDays AS (
                    SELECT DateUP, JiaType, HHour
                    FROM EMPHOLIDAY
                    WHERE empid = @userid_param AND YEAR(DateUP) = @year_param AND MONTH(DateUP) = @month_param
                )
                SELECT
                    COALESCE(w.workdat, FORMAT(h.DateUP, 'yyyyMMdd')) as workdat,
                    w.timeup,
                    w.timedown,
                    w.TOTH,
                    w.H1, w.H2, w.H3, w.B3, w.B4,
                    h.JiaType,
                    h.HHour as leaveHours
                FROM WorkDays w
                FULL OUTER JOIN HolidayDays h ON CAST(w.workdat AS DATE) = CAST(h.DateUP AS DATE)
                ORDER BY workdat ASC;
            `;
        
        const jiaTypeMap = { 'A': 'Việc riêng', 'B': 'Phép bệnh', 'E': 'Phép năm', 'D': 'Phép tang', 'F': 'Nghỉ thai sản', 'G': 'Nghỉ công tác', 'H': 'Nghỉ C.Thường', 'I': 'Đi đường', 'K': 'Không lương' };
        
        const formattedData = result.recordset.map(row => {
            if (!row.workdat) return null; 
            
            const year = row.workdat.substring(0, 4);
            const month = row.workdat.substring(4, 6);
            const day = row.workdat.substring(6, 8);
            const formattedDate = `${year}-${month}-${day}`;
            const formatTime = (t) => (!t || t.trim() === '000000' || t.trim() === '0') ? null : `${t.padStart(6, '0').substring(0, 2)}:${t.padStart(6, '0').substring(2, 4)}:${t.padStart(6, '0').substring(4, 6)}`;
            const jiaTypeCode = row.JiaType ? row.JiaType.trim() : '';

            let status = 'Nghỉ'; 
            const hasCheckIn = row.timeup && row.timeup.trim() !== '000000' && row.timeup.trim() !== '0';
            const hasLeave = row.leaveHours > 0;

            if (hasCheckIn) {
                status = hasLeave ? 'Đi làm & Nghỉ phép' : 'Đi làm';
            } else if (hasLeave) {
                status = 'Nghỉ phép';
            }
            return { 
                date: formattedDate, 
                checkIn: formatTime(row.timeup), 
                checkOut: formatTime(row.timedown), 
                hoursWorked: row.TOTH, 
                h1: row.H1, h2: row.H2, h3: row.H3,
                b3: row.B3, b4: row.B4,
                status: status,
                leaveHours: row.leaveHours || 0,
                leaveType: jiaTypeMap[jiaTypeCode] || ''
            };
        }).filter(Boolean); 
        
        res.status(200).json(formattedData);
    } catch (err) {
        console.error("Lỗi khi lấy dữ liệu chấm công:", err);
        res.status(500).json({ message: 'Lỗi server khi lấy dữ liệu chấm công.' });
    }
});
app.post('/api/admin/submit-leave', isAdmin, async (req, res) => {
    const { userId, startDate, endDate, leaveType, startTime, endTime, reason } = req.body;
    if (!userId || !startDate || !leaveType || !startTime || !endTime) {
        return res.status(400).json({ message: 'Vui lòng cung cấp đầy đủ thông tin.' });
    }
    
    const leaveTypeMap = {
        'E': 'Phép năm', 'A': 'Việc riêng', 'B': 'Phép Bệnh', 'C': 'Nghỉ kết hôn',
        'D': 'Phép Tang', 'F': 'Nghỉ thai sản', 'G': 'Nghỉ công tác', 'H': 'Nghỉ C.Thường',
        'I': 'Đi đường', 'K': 'Không lương'
    };
    const PUBLIC_HOLIDAYS_2025 = [
        '2025-01-01', '2025-01-28', '2025-01-29', '2025-01-30', '2025-01-31', 
        '2025-02-01', '2025-02-02', '2025-02-03', '2025-04-08', '2025-04-30', 
        '2025-05-01', '2025-05-02', '2025-09-01', '2025-09-02',
    ];
    const transaction = new sql.Transaction(pool);
    try {
        await transaction.begin();

        const sDate = parseDate(startDate);
        const eDate = parseDate(endDate || startDate);
        if (!sDate || !eDate || eDate < sDate) {
            await transaction.rollback();
            return res.status(400).json({ message: 'Định dạng ngày không hợp lệ hoặc ngày kết thúc nhỏ hơn ngày bắt đầu.' });
        }

        let totalHoursForPeriod = 0;
        let recordsToInsert = [];
        
        let currentDate = new Date(sDate.getTime());
        while(currentDate <= eDate) {
            const dayOfWeek = currentDate.getUTCDay(); 
            const dateString = currentDate.toISOString().split('T')[0];

            if (dayOfWeek === 0 || PUBLIC_HOLIDAYS_2025.includes(dateString)) {
                currentDate.setUTCDate(currentDate.getUTCDate() + 1);
                continue; 
            }
            
            let hoursThisDay = 0;
            let timeUp = startTime;
            let timeDown = endTime;

            const isSameDayPeriod = sDate.getTime() === eDate.getTime();
            const isFirstDay = currentDate.getTime() === sDate.getTime();
            const isLastDay = currentDate.getTime() === eDate.getTime();

            if (isSameDayPeriod) { hoursThisDay = calculateLeaveHours(startTime, endTime); } 
            else if (isFirstDay) { timeDown = '1700'; hoursThisDay = calculateLeaveHours(startTime, timeDown); } 
            else if (isLastDay) { timeUp = '0800'; hoursThisDay = calculateLeaveHours(timeUp, endTime); } 
            else { hoursThisDay = 8; }
            
            if (hoursThisDay > 0) {
                totalHoursForPeriod += hoursThisDay;
                recordsToInsert.push({
                    date: new Date(currentDate.getTime()),
                    timeUp: `${timeUp.substring(0,2)}:${timeUp.substring(2,4)}`,
                    timeDown: `${timeDown.substring(0,2)}:${timeDown.substring(2,4)}`,
                    hours: hoursThisDay,
                });
            }
            currentDate.setUTCDate(currentDate.getUTCDate() + 1);
        }

        if (totalHoursForPeriod <= 0) {
            await transaction.rollback();
            return res.status(400).json({ message: 'Không có giờ nghỉ nào được tính trong khoảng thời gian đã chọn (có thể do chỉ chọn ngày nghỉ/lễ).' });
        }
        
        const requestYear = sDate.getUTCFullYear();
        if (leaveType === 'E') {
            const summary = await getLeaveSummary(pool, userId, requestYear);
            if (totalHoursForPeriod > summary.remaining) {
                await transaction.rollback();
                return res.status(400).json({ 
                    message: `Không đủ phép. Tổng số giờ yêu cầu (${totalHoursForPeriod}) lớn hơn số giờ còn lại (${summary.remaining}).` 
                });
            }
        }
        
        for (const record of recordsToInsert) {
            const request = new sql.Request(transaction);
            const memoContent = reason || leaveTypeMap[leaveType] || leaveType;
            await request
                .input('empid', sql.VarChar, userId)
                .input('JiaType', sql.VarChar, leaveType)
                .input('DateUP', sql.DateTime, record.date)
                .input('TimeUP', sql.VarChar, record.timeUp)
                .input('DateDown', sql.DateTime, record.date)
                .input('TimeDown', sql.VarChar, record.timeDown)
                .input('HHour', sql.Decimal(10, 1), record.hours)
                .input('memo', sql.NVarChar, memoContent)
                .input('muser', sql.VarChar, req.session.user?.id || 'ADMIN_APP')
                .query(`
                    INSERT INTO dbo.EmpHoliday (empid, JiaType, DateUP, TimeUP, DateDown, TimeDown, HHour, memo, muser)
                    VALUES (@empid, @JiaType, @DateUP, @TimeUP, @DateDown, @TimeDown, @HHour, @memo, @muser)
                `);
        }
        await transaction.commit();
        res.status(201).json({ message: `Đã cập nhật thành công ${recordsToInsert.length} đơn phép vào Database.` });
    } catch (err) {
        try { await transaction.rollback(); } catch (rbErr) { console.error("Lỗi khi rollback transaction:", rbErr); }
        console.error("Lỗi khi admin nhập phép:", err);
        const dbError = err.originalError ? err.originalError.info.message : err.message;
        res.status(500).json({ message: `Lỗi server khi xử lý yêu cầu: ${dbError}` });
    }
});
app.get('/api/groups', isAdmin, async (req, res) => {
    try {
        await poolConnect;
        const result = await pool.request()
            .query`SELECT SYS_TYPE as groupId, SYS_VALUE as groupName FROM BASICCODE WHERE FUNC = 'GROUPID' ORDER BY SYS_TYPE`;
        res.status(200).json(result.recordset);
    } catch (err) {
        console.error("Lỗi khi lấy danh sách bộ phận:", err);
        res.status(500).json({ message: 'Lỗi server khi lấy danh sách bộ phận.' });
    }
});

app.get('/api/admin/export-timesheet/:yearMonth', isAdmin, async (req, res) => {
    const { yearMonth } = req.params;
    const { groupId } = req.query; 
    const yearMonthFormatted = yearMonth.replace('-', '');
    const year = parseInt(yearMonth.substring(0, 4));
    const month = parseInt(yearMonth.substring(5, 7));

    try {
        await poolConnect;

        let query = `
            WITH WorkData AS (
                SELECT
                    EmpID,
                    SUM(ISNULL(TOTH, 0)) as TotalHours,
                    SUM(ISNULL(H1, 0)) as TotalH1,
                    SUM(ISNULL(H2, 0)) as TotalH2,
                    SUM(ISNULL(H3, 0)) as TotalH3,
                    SUM(ISNULL(B3, 0)) as TotalB3,
                    SUM(ISNULL(B4, 0)) as TotalB4
                FROM
                    EMPWORK
                WHERE
                    LEFT(workdat, 6) = @yearMonthFormatted
                GROUP BY
                    EmpID
            ),
            EmployeeDepartment AS (
                SELECT
                    f.EMPID,
                    f.EMPNAM_VN,
                    latest_sal.GROUPID
                FROM
                    EMPFILE f
                OUTER APPLY (
                    SELECT TOP 1 s.GROUPID
                    FROM EMPDSALARY s
                    WHERE s.EMPID = f.EMPID
                    ORDER BY s.YYMM DESC
                ) latest_sal
                WHERE
                    f.OUTDAT IS NULL
            ),
            LeaveData AS (
                SELECT
                    empid,
                    SUM(CASE WHEN JiaType = 'E' THEN ISNULL(HHour, 0) ELSE 0 END) as PhepNam,
                    SUM(CASE WHEN JiaType = 'A' THEN ISNULL(HHour, 0) ELSE 0 END) as ViecRieng,
                    SUM(CASE WHEN JiaType = 'B' THEN ISNULL(HHour, 0) ELSE 0 END) as PhepBenh,
                    SUM(CASE WHEN JiaType = 'D' THEN ISNULL(HHour, 0) ELSE 0 END) as PhepTang
                FROM
                    EMPHOLIDAY
                WHERE
                    YEAR(DateUP) = @year AND MONTH(DateUP) = @month
                GROUP BY
                    empid
            )
            SELECT
                ed.GROUPID as 'Mã Bộ Phận',
                g.SYS_VALUE as 'Tên Bộ Phận',
                ed.EMPID as 'Mã Nhân Viên',
                ed.EMPNAM_VN as 'Tên Nhân Viên',
                ISNULL(w.TotalHours, 0) as 'Tổng Giờ Làm',
                ISNULL(w.TotalH1, 0) as 'Tăng Ca 1.5',
                ISNULL(w.TotalH2, 0) as 'Tăng Ca 2.0',
                ISNULL(w.TotalH3, 0) as 'Tăng Ca 3.0',
                ISNULL(w.TotalB3, 0) as 'Tăng Ca Đêm',
                ISNULL(w.TotalB4, 0) as 'Phụ Cấp 0.5',
                ISNULL(l.PhepNam, 0) as 'Phép Năm (giờ)',
                ISNULL(l.ViecRieng, 0) as 'Việc Riêng (giờ)',
                ISNULL(l.PhepBenh, 0) as 'Phép Bệnh (giờ)',
                ISNULL(l.PhepTang, 0) as 'Phép Tang (giờ)'
            FROM
                EmployeeDepartment ed
            LEFT JOIN WorkData w ON ed.EMPID = w.EmpID
            LEFT JOIN LeaveData l ON ed.EMPID = l.empid
            LEFT JOIN BASICCODE g ON ed.GROUPID = g.SYS_TYPE AND g.FUNC = 'GROUPID'
        `;
        
        if (groupId && groupId !== 'ALL') {
            query += ` WHERE ed.GROUPID = @groupId`;
        }
        
        query += ` ORDER BY ed.GROUPID, ed.EMPID`;
        
        const request = pool.request();
        request.input('yearMonthFormatted', sql.VarChar, yearMonthFormatted);
        request.input('year', sql.Int, year);
        request.input('month', sql.Int, month);

        if (groupId && groupId !== 'ALL') {
            request.input('groupId', sql.VarChar, groupId);
        }

        const result = await request.query(query);

        const worksheet = xlsx.utils.json_to_sheet(result.recordset);
        worksheet['!cols'] = [ 
            { wch: 15 }, { wch: 25 }, { wch: 15 }, { wch: 30 }, 
            { wch: 15 }, { wch: 15 }, { wch: 15 }, { wch: 15 },
            { wch: 15 }, { wch: 15 }, { wch: 18 }, { wch: 18 },
            { wch: 18 }, { wch: 18 }
        ];
        const workbook = xlsx.utils.book_new();
        xlsx.utils.book_append_sheet(workbook, worksheet, `ChamCong_${yearMonth}`);
        
        const buffer = xlsx.write(workbook, { bookType: 'xlsx', type: 'buffer' });
        res.setHeader('Content-Disposition', `attachment; filename=BaoCaoChamCong_${yearMonth}${groupId && groupId !== 'ALL' ? '_' + groupId : ''}.xlsx`);
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.send(buffer);

    } catch (err) {
        console.error("Lỗi khi xuất file chấm công:", err);
        res.status(500).json({ message: 'Lỗi server khi xuất file chấm công.' });
    }
});

// --- 5. PHỤC VỤ FILE TĨNH VÀ ROUTING ---
app.use(express.static(path.join(__dirname, 'public')));
app.use(express.static(path.join(__dirname, 'build')));
app.get('*', (req, res) => {
    res.sendFile(path.join(__dirname, 'build', 'index.html'));
});

// --- 6. KHỞI CHẠY SERVER ---
const startServer = async () => {
    try {
        await fs.mkdir(SESSIONS_DIR, { recursive: true });
        console.log(`Thư mục session đã sẵn sàng tại: ${SESSIONS_DIR}`);

        app.listen(port, () => {
            console.log(`Backend server đang chạy tại http://localhost:${port}`);
        });
    } catch (error) {
        console.error('Không thể khởi động server:', error);
    }
};

startServer();
