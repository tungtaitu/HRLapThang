/*
 * File: controllers/leave.controller.js
 * Mô tả: Chứa logic xử lý cho các chức năng quản lý phép năm của admin.
 * Phiên bản đúng, sử dụng cột 'autoid' từ CSDL.
 */
const sql = require('mssql');
const { pool, poolConnect } = require('../config/db');
const { getLeaveSummary } = require('../services/leave.service');
const { calculateLeaveHours } = require('../utils/calculate');
const { parseDate } = require('../utils/helpers');
const { PUBLIC_HOLIDAYS } = require('../constants/holidays');

// Chức năng nhập phép
exports.submitLeave = async (req, res) => {
    const { userId, startDate, endDate, leaveType, startTime, endTime, reason } = req.body;
    if (!userId || !startDate || !leaveType || !startTime || !endTime) {
        return res.status(400).json({ message: 'Vui lòng cung cấp đầy đủ thông tin.' });
    }
    const leaveTypeMap = { 'E': 'Phép năm', 'A': 'Việc riêng', 'B': 'Phép Bệnh', 'C': 'Nghỉ kết hôn', 'D': 'Phép Tang', 'F': 'Nghỉ thai sản', 'G': 'Nghỉ công tác', 'H': 'Nghỉ C.Thường', 'I': 'Đi đường', 'K': 'Không lương' };
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
            if (dayOfWeek === 0 || PUBLIC_HOLIDAYS[dateString]) {
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
                recordsToInsert.push({ date: new Date(currentDate.getTime()), timeUp: `${timeUp.substring(0,2)}:${timeUp.substring(2,4)}`, timeDown: `${timeDown.substring(0,2)}:${timeDown.substring(2,4)}`, hours: hoursThisDay });
            }
            currentDate.setUTCDate(currentDate.getUTCDate() + 1);
        }
        if (totalHoursForPeriod <= 0) {
            await transaction.rollback();
            return res.status(400).json({ message: 'Không có giờ nghỉ nào được tính trong khoảng thời gian đã chọn.' });
        }
        const requestYear = sDate.getUTCFullYear();
        if (leaveType === 'E') {
            const summary = await getLeaveSummary(pool, userId, requestYear);
            if (totalHoursForPeriod > summary.remaining) {
                await transaction.rollback();
                return res.status(400).json({ message: `Không đủ phép. Yêu cầu (${totalHoursForPeriod}) > Còn lại (${summary.remaining}).` });
            }
        }
        for (const record of recordsToInsert) {
            const request = new sql.Request(transaction);
            const memoContent = reason || leaveTypeMap[leaveType] || leaveType;
            await request.input('empid', sql.VarChar, userId).input('JiaType', sql.VarChar, leaveType).input('DateUP', sql.DateTime, record.date).input('TimeUP', sql.VarChar, record.timeUp).input('DateDown', sql.DateTime, record.date).input('TimeDown', sql.VarChar, record.timeDown).input('HHour', sql.Decimal(10, 1), record.hours).input('memo', sql.NVarChar, memoContent).input('muser', sql.VarChar, req.session.user?.id || 'ADMIN_APP')
                .query(`INSERT INTO dbo.EmpHoliday (empid, JiaType, DateUP, TimeUP, DateDown, TimeDown, HHour, memo, muser) VALUES (@empid, @JiaType, @DateUP, @TimeUP, @DateDown, @TimeDown, @HHour, @memo, @muser)`);
        }
        await transaction.commit();
        res.status(201).json({ message: `Đã cập nhật thành công ${recordsToInsert.length} đơn phép.` });
    } catch (err) {
        try { await transaction.rollback(); } catch (rbErr) { console.error("Lỗi khi rollback transaction:", rbErr); }
        console.error("Lỗi khi admin nhập phép:", err);
        const dbError = err.originalError ? err.originalError.info.message : err.message;
        res.status(500).json({ message: `Lỗi server: ${dbError}` });
    }
};

// Chức năng lấy danh sách phép
exports.getLeaveEntries = async (req, res) => {
    const { userId, year } = req.params;
    try {
        await poolConnect;
        const result = await pool.request().input('empid', sql.VarChar, userId).input('year', sql.Int, year)
            // Lấy thêm TimeDown để phục vụ cho việc sửa
            .query(`
                SELECT 
                    autoid as id, DateUP, TimeUP, TimeDown, JiaType, HHour, memo 
                FROM EMPHOLIDAY 
                WHERE empid = @empid AND YEAR(DateUP) = @year 
                ORDER BY DateUP DESC
            `);
        
        const jiaTypeMap = { 'E': 'Phép năm', 'A': 'Việc riêng', 'B': 'Phép Bệnh', 'C': 'Nghỉ kết hôn', 'D': 'Phép Tang', 'F': 'Nghỉ thai sản', 'G': 'Nghỉ công tác', 'H': 'Nghỉ C.Thường', 'I': 'Đi đường', 'K': 'Không lương' };
        const formattedData = result.recordset.map(row => ({ ...row, JiaTypeName: jiaTypeMap[row.JiaType.trim()] || row.JiaType.trim() }));
        res.status(200).json(formattedData);
    } catch (err) {
        console.error("Lỗi khi lấy danh sách ngày phép:", err);
        res.status(500).json({ message: 'Lỗi server khi lấy dữ liệu.' });
    }
};

// --- HÀM MỚI: SỬA MỘT NGÀY PHÉP ---
exports.updateLeaveEntry = async (req, res) => {
    const { id } = req.params;
    const { DateUP, TimeUP, TimeDown, HHour, JiaType, memo } = req.body;
    const adminUserId = req.session.user?.id || 'ADMIN_APP';

    try {
        await poolConnect;
        const request = pool.request().input('id', sql.Int, id);

        // Tính toán lại số giờ nghỉ nếu thời gian thay đổi
        const updatedHHour = (TimeUP && TimeDown) ? calculateLeaveHours(TimeUP.replace(':', ''), TimeDown.replace(':', '')) : HHour;

        const result = await request
            .input('DateUP', sql.DateTime, new Date(DateUP))
            .input('TimeUP', sql.VarChar, TimeUP)
            .input('TimeDown', sql.VarChar, TimeDown)
            .input('HHour', sql.Decimal(10, 1), updatedHHour)
            .input('JiaType', sql.VarChar, JiaType)
            .input('memo', sql.NVarChar, memo)
            .query(`
                UPDATE EMPHOLIDAY 
                SET DateUP = @DateUP, TimeUP = @TimeUP, TimeDown = @TimeDown, 
                    HHour = @HHour, JiaType = @JiaType, memo = @memo
                WHERE autoid = @id
            `);
        
        if (result.rowsAffected[0] > 0) {
            console.log(`Admin [${adminUserId}] đã cập nhật thành công bản ghi phép có ID: ${id}`);
            res.status(200).json({ message: 'Đã cập nhật ngày phép thành công.' });
        } else {
            res.status(404).json({ message: 'Không tìm thấy ngày phép để cập nhật.' });
        }
    } catch (err) {
        console.error(`Lỗi khi cập nhật bản ghi phép ID ${id}:`, err);
        res.status(500).json({ message: 'Lỗi server khi cập nhật dữ liệu.' });
    }
};


// Chức năng xóa phép
exports.deleteLeaveEntry = async (req, res) => {
    const { id } = req.params;
    const adminUserId = req.session.user?.id || 'ADMIN_APP';
    try {
        await poolConnect;
        const result = await pool.request()
            .input('id', sql.Int, id)
            .query('DELETE FROM EMPHOLIDAY WHERE autoid = @id');
        if (result.rowsAffected[0] > 0) {
            console.log(`Admin [${adminUserId}] đã xóa thành công bản ghi phép có ID: ${id}`);
            res.status(200).json({ message: 'Đã xóa ngày phép thành công.' });
        } else {
            res.status(404).json({ message: 'Không tìm thấy ngày phép để xóa.' });
        }
    } catch (err) {
        console.error(`Lỗi khi xóa bản ghi phép ID ${id}:`, err);
        res.status(500).json({ message: 'Lỗi server khi xóa dữ liệu.' });
    }
};
