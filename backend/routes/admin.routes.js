/*
 * File: routes/admin.routes.js
 * Mô tả: Định tuyến cho các chức năng của quản trị viên.
 */

const express = require('express');
const router = express.Router();
const multer = require('multer');
const adminController = require('../controllers/admin.controller');
const { isAdmin } = require('../middleware/auth.middleware');

const upload = multer({ dest: 'uploads/' });

// Tất cả các route trong file này đều yêu cầu quyền admin
router.use(isAdmin);

router.post('/reset-password', adminController.resetPassword);
router.get('/employee-info/:empid', adminController.getEmployeeInfo);
router.post('/upload-leave', upload.single('leaveFile'), adminController.uploadLeaveFile);
router.get('/export-leave/:year', adminController.exportLeaveFile);
router.get('/all-payrolls/:yearMonth', adminController.getAllPayrolls);
router.post('/approve-payroll', adminController.approvePayroll);
router.get('/export-payrolls/:yearMonth', adminController.exportPayrolls);
router.get('/groups', adminController.getGroups);
router.get('/export-timesheet/:yearMonth', adminController.exportTimesheet);
router.post('/upload-luong-t13', upload.single('luongT13File'), adminController.uploadLuongT13);
router.get('/basic-code/:func', adminController.getBasicCodeOptions);
module.exports = router;
