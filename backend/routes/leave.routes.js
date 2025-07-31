/*
 * File: routes/leave.routes.js
 * Mô tả: Định tuyến cho các chức năng quản lý phép năm của admin.
 */
const express = require('express');
const router = express.Router();
const leaveController = require('../controllers/leave.controller');
const { isAdmin } = require('../middleware/auth.middleware');

// Tất cả các route trong file này đều yêu cầu quyền admin
router.use(isAdmin);

// Route để admin nhập phép thủ công
router.post('/submit-leave', leaveController.submitLeave);

// Route để lấy danh sách chi tiết các ngày phép đã nhập
router.get('/leave-entries/:userId/:year', leaveController.getLeaveEntries);

// Route để xóa một ngày phép cụ thể
router.delete('/leave-entry/:id', leaveController.deleteLeaveEntry);

// --- ROUTE MỚI ĐỂ SỬA ---
router.put('/leave-entry/:id', leaveController.updateLeaveEntry);

module.exports = router;
