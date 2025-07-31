/*
 * File: routes/auth.routes.js
 * Mô tả: Định tuyến cho các chức năng xác thực người dùng.
 */

const express = require('express');
const router = express.Router();
const authController = require('../controllers/auth.controller');
const { isAuthenticated } = require('../middleware/auth.middleware');

router.post('/login', authController.login);
router.post('/logout', authController.logout);
router.get('/check-session', authController.checkSession);
router.post('/user/change-password', isAuthenticated, authController.changePassword);

module.exports = router;
