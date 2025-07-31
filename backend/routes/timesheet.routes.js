/*
 * File: routes/timesheet.routes.js
 * Mô tả: Định tuyến cho các chức năng chấm công của admin.
 */
const express = require('express');
const router = express.Router();
const timesheetController = require('../controllers/timesheet.controller');
const { isAdmin } = require('../middleware/auth.middleware');

router.use(isAdmin);

// --- CÁC ROUTE CỤ THỂ PHẢI ĐƯỢC ĐẶT LÊN TRƯỚC ---

// Route để lấy GROUPID của tất cả nhân viên
router.get('/groupids', timesheetController.getAllEmployeeGroupIDs);

// Route để lấy bảng tổng hợp chấm công theo tháng
// GET /api/admin/timesheet/summary/202507?groupId=A033
router.get('/summary/:yymm', timesheetController.getMonthlyTimesheetSummary);


// --- CÁC ROUTE CHUNG CHUNG ĐẶT Ở DƯỚI ---

// Route để upload file chấm công
router.post('/upload', timesheetController.uploadTimesheet);

// Route để tra cứu chấm công chi tiết theo mã nhân viên và tháng
router.get('/:empid/:yymm', timesheetController.getTimesheetForEmployee);

// Route để cập nhật một dòng chấm công bằng autoid
router.put('/:autoid', timesheetController.updateTimesheetEntry);


module.exports = router;
