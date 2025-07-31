/*
 * File: routes/index.js
 * Mô tả: File router tổng, gộp tất cả các router con lại.
 */
const express = require('express');
const router = express.Router();

const authRoutes = require('./auth.routes');
const adminRoutes = require('./admin.routes');
const userRoutes = require('./user.routes');
const leaveRoutes = require('./leave.routes');
const employeeRoutes = require('./employee.routes');
const timesheetRoutes = require('./timesheet.routes'); // <-- IMPORT MỚI

// Gắn các router con vào router chính
router.use(authRoutes);
router.use('/admin', adminRoutes);
router.use('/admin', leaveRoutes);
router.use('/admin/employee', employeeRoutes);
router.use('/admin/timesheet', timesheetRoutes); // <-- SỬ DỤNG ROUTE MỚI
router.use(userRoutes);

module.exports = router;
