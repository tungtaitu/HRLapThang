/*
 * File: controllers/auth.controller.js
 * Mô tả: Chứa logic xử lý cho các route xác thực.
 */

const sql = require('mssql');
const { pool, poolConnect } = require('../config/db');
const { jsonDataState, writeJsonAndUpdateState, USER_PASSWORDS_FILE } = require('../services/json.service');
const { calculateWorkDuration } = require('../utils/calculate');

exports.login = async (req, res) => {
    const { empid, password } = req.body;
    if (!empid || !password) {
        return res.status(400).json({ message: 'Tên đăng nhập và Mật khẩu không được để trống.' });
    }

    try {
        await poolConnect;

        // Kiểm tra tài khoản admin
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

        // Kiểm tra tài khoản nhân viên
        let userResult = null;
        const customPasswords = jsonDataState.userPasswords;
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
};

exports.checkSession = (req, res) => {
    if (req.session && req.session.user) {
        res.status(200).json(req.session.user);
    } else {
        res.status(401).json({ message: 'Chưa đăng nhập.' });
    }
};

exports.logout = (req, res) => {
    req.session.destroy(err => {
        if (err) {
            return res.status(500).json({ message: 'Đăng xuất thất bại.' });
        }
        res.clearCookie('app.sid', { path: '/' });
        res.status(200).json({ message: 'Đăng xuất thành công.' });
    });
};

exports.changePassword = async (req, res) => {
    const { userId, oldPassword, newPassword } = req.body;
    if (!userId || !oldPassword || !newPassword) {
        return res.status(400).json({ message: 'Vui lòng nhập đầy đủ thông tin.' });
    }

    try {
        await poolConnect;

        let isValidOldPassword = false;
        const customPasswords = jsonDataState.userPasswords;
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

        await writeJsonAndUpdateState(USER_PASSWORDS_FILE, updatedPasswords);
        res.status(200).json({ message: 'Đổi mật khẩu thành công!' });

    } catch (err) {
        console.error("Lỗi khi đổi mật khẩu:", err);
        res.status(500).json({ message: 'Lỗi server khi đổi mật khẩu.' });
    }
};
