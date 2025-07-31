/*
 * File: routes/employee.routes.js
 * Mô tả: Định tuyến cho các chức năng quản lý nhân viên của admin.
 * Cập nhật: Thêm route mới để lấy chi tiết tất cả nhân viên.
 */
const express = require('express');
const router = express.Router();
const employeeController = require('../controllers/employee.controller');
const { isAdmin } = require('../middleware/auth.middleware');

router.use(isAdmin);


// GET /api/admin/employee -> Lấy danh sách tất cả nhân viên
router.get('/', employeeController.getAllEmployees);

// POST /api/admin/employee -> Thêm nhân viên mới
router.post('/', employeeController.addEmployee);

// GET /api/admin/employee/:empid/info -> Lấy thông tin chi tiết
router.get('/:empid/info', employeeController.getEmployeeInfo);

// PUT /api/admin/employee/:empid -> Cập nhật thông tin
router.put('/:empid', employeeController.updateEmployee);

// DELETE /api/admin/employee/:empid -> Xóa mềm (thôi việc)
router.put('/:empid/resign', employeeController.deleteEmployee);

// POST /api/admin/employee/reset-password -> Reset mật khẩu
router.post('/reset-password', employeeController.resetPassword);

// GET /api/admin/employee/next-id --> Lấy mã nhân viên tiếp theo
router.get('/next-id', employeeController.getNextEmployeeId);


module.exports = router;
