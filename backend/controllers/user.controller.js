/*
 * File: controllers/user.controller.js
 * Mô tả: Chứa logic xử lý cho các route của nhân viên.
 * Sử dụng jsonDataState đã được đồng bộ hóa tự động.
 */
const sql = require('mssql');
const { pool, poolConnect } = require('../config/db');
const { jsonDataState } = require('../services/json.service'); 
const { getLeaveSummary } = require('../services/leave.service');

exports.getHolidays = async (req, res) => {
    const { userId, year } = req.params;
    try {
        await poolConnect;
        const employeeResult = await pool.request()
            .input('userid_param', sql.VarChar, userId)
            .query`SELECT EMPNAM_VN FROM EMPFILE WHERE EMPID = @userid_param`;

        if (employeeResult.recordset.length === 0) {
            return res.status(404).json({ message: 'Không tìm thấy nhân viên.' });
        }
        const employeeName = employeeResult.recordset[0].EMPNAM_VN;
        const summary = await getLeaveSummary(pool, userId, year);

        const holidayListResult = await pool.request()
            .input('userid_param', sql.VarChar, userId)
            .input('year_param', sql.Int, year)
            .query`SELECT DateUP, JiaType, HHour, memo FROM EMPHOLIDAY WHERE empid = @userid_param AND YEAR(DateUP) = @year_param ORDER BY DateUP ASC`;

        const jiaTypeMap = { 'A': 'Việc riêng', 'B': 'Phép bệnh', 'E': 'Phép năm','C': 'Nghỉ kết hôn', 'D': 'Phép tang','G': 'Nghỉ công tác', 'H': 'Nghỉ C.Thường',
        'I': 'Đi đường', 'K': 'Không lương' };

        const formattedHolidayList = holidayListResult.recordset.map(row => {
            const jiaTypeCode = row.JiaType ? row.JiaType.trim() : '';
            return {
                date: row.DateUP,
                reason: jiaTypeMap[jiaTypeCode] || jiaTypeCode,
                hours: row.HHour || 0,
                memo: row.memo || null
            };
        });
        
        const employeeConfig = jsonDataState.leaveData.find(emp => emp.MSNV === userId);
        const isForeigner = employeeConfig?.isForeigner || false;

        res.status(200).json({
            employeeName: employeeName,
            holidayList: formattedHolidayList,
            summary: summary,
            isForeigner: isForeigner
        });

    } catch (err) {
        console.error("Lỗi API Holidays:", err);
        res.status(500).json({ message: 'Lỗi server khi lấy dữ liệu ngày nghỉ.' });
    }
};


exports.getPayroll = async (req, res) => {
    const { userId, yearMonth } = req.params; // yearMonth có dạng 'YYYY-MM'
    const yearMonthFormatted = yearMonth.replace('-', '');

    const sessionUser = req.session.user;
    const isAdmin = sessionUser && sessionUser.isAdmin;

    // Logic này giờ đã đáng tin cậy vì jsonDataState luôn được cập nhật
    const approvals = jsonDataState.payrollApprovals;
    const isMonthApproved = approvals.includes(yearMonth);
    
    // Chỉ kiểm tra phê duyệt nếu người dùng không phải là admin
    if (!isAdmin) {
        const approvalStartDate = new Date('2025-03-01');
        const selectedPayrollDate = new Date(yearMonth + '-01');

        if (selectedPayrollDate >= approvalStartDate && !isMonthApproved) {
            return res.status(200).json({ 
                approved: false,
                message: 'Phiếu lương cho tháng này chưa được phê duyệt.' 
            });
        }
    }

    try {
        await poolConnect;
        const result = await pool.request()
            .input('userid_param', sql.VarChar, userId)
            .input('yymm_param', sql.VarChar, yearMonthFormatted)
            .query`
                SELECT TOP 1
                    f.EMPID, f.EMPNAM_VN,
                    s.*,
                    job_code.SYS_VALUE as ChucVu,
                    group_code.SYS_VALUE as DonVi,
                    d.BHXH5, d.BHYT1, d.BHTN1
                FROM
                    EMPFILE f
                OUTER APPLY (
                    SELECT TOP 1 BB, CV, PHU, NN, KT, MT ,TTKH, TNKH, QC, GT, KTAXM, MONEY_H, REAL_TOTAL, workdays, JOB, GROUPID, H1, H1M, H2, H2M, H3, H3M, B3, B3M, B4, B4M, B5, B5M, JX, BZKM, QITA, SOLE
                    FROM EMPDSALARY
                    WHERE EMPID = f.EMPID AND YYMM = @yymm_param
                ) s
                OUTER APPLY (
                    SELECT TOP 1 BHXH5, BHYT1, BHTN1
                    FROM EMPBHGT
                    WHERE EMPID = f.EMPID AND YYMM = @yymm_param
                ) d
                OUTER APPLY (
                    SELECT TOP 1 SYS_VALUE FROM BASICCODE WHERE SYS_TYPE = s.JOB
                ) job_code
                OUTER APPLY (
                    SELECT TOP 1 SYS_VALUE FROM BASICCODE WHERE SYS_TYPE = s.GROUPID
                ) group_code
                WHERE f.EMPID = @userid_param
            `;

        if (result.recordset.length > 0) {
            const data = result.recordset[0];

            const payrollDetails = {
                employeeInfo: {
                    soThe: data.EMPID, hoTen: data.EMPNAM_VN, chucVu: data.ChucVu || data.JOB || 'N/A',
                    donVi: data.DonVi || data.GROUPID || 'N/A',  nam: yearMonth.substring(0, 4), thang: yearMonth.substring(5, 7)
                },
                earnings: [
                    { label: 'LƯƠNG CB (BB)', value: data.BB || 0 }, { label: 'CHỨC VỤ (CV)', value: data.CV || 0 },
                    { label: 'ĐIỆN THOẠI', value: data.PHU || 0 }, { label: 'XĂNG XE', value: data.NN || 0 },
                    { label: 'KỸ THUẬT', value: data.KT || 0 }, { label: 'MÔI TRƯỜNG', value: data.MT || 0 },
                    { label: 'NHÀ Ở', value: data.TTKH || 0 }, { label: 'CHUYÊN CẦN', value: data.QC || 0 }
                ],
                deductions: [
                    { label: 'Trừ phép thường', value: data.BZKM || 0 }, { label: 'BHXH', value: data.BHXH5 || 0 },
                    { label: 'BHYT', value: data.BHYT1 || 0 }, { label: 'BHTN', value: data.BHTN1 || 0 },
                    { label: 'Phí công đoàn', value: data.GT || 0 }, { label: 'Trừ tiền khác', value: data.QITA || 0 },
                    { label: 'Thuế TN Cá Nhân', value: data.KTAXM || 0 }
                ],
                overtimeAndBonus: [
                    { label: 'Phụ cấp 0.5', hours: data.B4 || 0, amount: data.B4M || 0 },
                    { label: 'Phụ cấp 0.3', hours: data.B5 || 0, amount: data.B5M || 0 },
                    { label: 'Phụ Cấp loại A', hours: data.H1 || 0, amount: data.H1M || 0 },
                    { label: 'Phụ cấp loại B', hours: data.H2 || 0, amount: data.H2M || 0 },
                    { label: 'Phụ cấp loại C', hours: data.H3 || 0, amount: data.H3M || 0 },
                    { label: 'Phụ cấp loại D', hours: data.B3 || 0, amount: data.B3M || 0 },
                    { label: 'Phụ cấp loại Z', hours: '-' , amount: data.JX || 0 },
                    { label: 'Thu nhập khác' , hours: '-' , amount: data.TNKH || 0 }
                ],
                summary: {
                    tinhLuongMoiGio: data.MONEY_H || 0, tongSoNgayLam: data.workdays || 0,
                    soLe: data.SOLE || 0, luongThucLanh: data.REAL_TOTAL || 0
                },
                approved: true
            };
            res.status(200).json(payrollDetails);
        } else {
            res.status(200).json(null);
        }
    } catch (err) {
        console.error("Lỗi khi lấy dữ liệu lương của nhân viên:", err);
        res.status(500).json({ message: 'Lỗi server khi lấy dữ liệu lương.' });
    }
};

exports.getTimesheet = async (req, res) => {
    const { userId, yearMonth } = req.params;
    const yearMonthFormatted = yearMonth.replace('-', '');
    const year = parseInt(yearMonth.substring(0, 4));
    const month = parseInt(yearMonth.substring(5, 7));

    try {
        await poolConnect;

        const result = await pool.request()
            .input('userid_param', sql.VarChar, userId)
            .input('yymm_param', sql.VarChar, yearMonthFormatted)
            .input('year_param', sql.Int, year)
            .input('month_param', sql.Int, month)
            .query`
                WITH WorkDays AS (
                    SELECT workdat, timeup, timedown, TOTH, H1, H2, H3, B3, B4
                    FROM EMPWORK
                    WHERE EmpID = @userid_param AND LEFT(workdat, 6) = @yymm_param
                ),
                HolidayDays AS (
                    SELECT DateUP, JiaType, HHour
                    FROM EMPHOLIDAY
                    WHERE empid = @userid_param AND YEAR(DateUP) = @year_param AND MONTH(DateUP) = @month_param
                )
                SELECT
                    COALESCE(w.workdat, FORMAT(h.DateUP, 'yyyyMMdd')) as workdat,
                    w.timeup,
                    w.timedown,
                    w.TOTH,
                    w.H1, w.H2, w.H3, w.B3, w.B4,
                    h.JiaType,
                    h.HHour as leaveHours
                FROM WorkDays w
                FULL OUTER JOIN HolidayDays h ON CAST(w.workdat AS DATE) = CAST(h.DateUP AS DATE)
                ORDER BY workdat ASC;
            `;

        const jiaTypeMap = { 'A': 'Việc riêng', 'B': 'Phép bệnh', 'E': 'Phép năm', 'C': 'Nghỉ kết hôn', 'D': 'Phép tang', 'F': 'Nghỉ thai sản', 'G': 'Nghỉ công tác', 'H': 'Nghỉ C.Thường',
        'I': 'Đi đường', 'K': 'Không lương' };

        const formattedData = result.recordset.map(row => {
            if (!row.workdat) return null;

            const year = row.workdat.substring(0, 4);
            const month = row.workdat.substring(4, 6);
            const day = row.workdat.substring(6, 8);
            const formattedDate = `${year}-${month}-${day}`;
            const formatTime = (t) => (!t || t.trim() === '000000' || t.trim() === '0') ? null : `${t.padStart(6, '0').substring(0, 2)}:${t.padStart(6, '0').substring(2, 4)}:${t.padStart(6, '0').substring(4, 6)}`;
            const jiaTypeCode = row.JiaType ? row.JiaType.trim() : '';

            let status = 'Nghỉ';
            const hasCheckIn = row.timeup && row.timeup.trim() !== '000000' && row.timeup.trim() !== '0';
            const hasLeave = row.leaveHours > 0;

            if (hasCheckIn) {
                status = hasLeave ? 'Đi làm & Nghỉ phép' : 'Đi làm';
            } else if (hasLeave) {
                status = 'Nghỉ phép';
            }
            return {
                date: formattedDate,
                checkIn: formatTime(row.timeup),
                checkOut: formatTime(row.timedown),
                hoursWorked: row.TOTH,
                h1: row.H1, h2: row.H2, h3: row.H3,
                b3: row.B3, b4: row.B4,
                status: status,
                leaveHours: row.leaveHours || 0,
                leaveType: jiaTypeMap[jiaTypeCode] || ''
            };
        }).filter(Boolean);

        res.status(200).json(formattedData);
    } catch (err) {
        console.error("Lỗi khi lấy dữ liệu chấm công:", err);
        res.status(500).json({ message: 'Lỗi server khi lấy dữ liệu chấm công.' });
    }
};

exports.getLuongT13 = async (req, res) => {
    const { userId, year } = req.params;
    try {
        const allData = jsonDataState.luongT13Data;
        const yearData = allData[year];
        if (!yearData) return res.status(200).json(null);
        const userData = yearData.find(emp => emp.MSNV === userId);
        res.status(200).json(userData || null);
    } catch (error) {
        console.error(`Lỗi khi lấy dữ liệu lương T13 cho ${userId}/${year}:`, error);
        res.status(500).json({ message: 'Lỗi server khi lấy dữ liệu.' });
    }
};
