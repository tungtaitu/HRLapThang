/*
 * File: routes/user.routes.js
 * Mô tả: Định tuyến cho các chức năng của nhân viên.
 */

const express = require('express');
const router = express.Router();
const userController = require('../controllers/user.controller');
const { isAuthenticated } = require('../middleware/auth.middleware');

// Tất cả các route trong file này đều yêu cầu người dùng phải đăng nhập
router.use(isAuthenticated);

router.get('/holidays/:userId/:year', userController.getHolidays);
router.get('/payroll/:userId/:yearMonth', userController.getPayroll);
router.get('/timesheet/:userId/:yearMonth', userController.getTimesheet);
router.get('/luong-t13/:userId/:year', userController.getLuongT13);

module.exports = router;
