/*
 * File: middleware/auth.middleware.js
 * Mô tả: Chứa các middleware để kiểm tra xác thực và quyền truy cập.
 */

const isAdmin = (req, res, next) => {
    if (req.session && req.session.user && req.session.user.isAdmin) {
        return next();
    }
    return res.status(401).json({ message: 'Chưa đăng nhập hoặc không có quyền. Vui lòng đăng nhập lại.' });
};

const isAuthenticated = (req, res, next) => {
    if (req.session && req.session.user) {
        return next();
    }
    return res.status(401).json({ message: 'Chưa đăng nhập.' });
};

module.exports = {
    isAdmin,
    isAuthenticated
};
