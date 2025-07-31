/*
 * File: controllers/timesheet.controller.js
 * Mô tả: Chứa logic xử lý cho các chức năng chấm công.
 * Cập nhật: Sửa lỗi tại API getTimesheetForEmployee để đảm bảo 'autoid' luôn là một số duy nhất.
 */
const sql = require('mssql');
const { pool, poolConnect } = require('../config/db');
const { PUBLIC_HOLIDAYS } = require('../constants/holidays');

// --- CẤU HÌNH CÁC CA LÀM VIỆC ---
const SHIFT_CONFIG = {
    'HC': { name: 'Hành chính', startTime: '08:00', endTime: '17:00', breakMinutes: 60, standardHours: 8 },
    'CA1': { name: 'Ca 1', startTime: '06:00', endTime: '14:00', breakMinutes: 0, standardHours: 7.5 },
    'CA2': { name: 'Ca 2', startTime: '14:00', endTime: '22:00', breakMinutes: 0, standardHours: 7.5 },
    'CADEM': { name: 'Ca Đêm', startTime: '22:00', endTime: '06:00', breakMinutes: 0, standardHours: 7.5 }
};

// --- HÀM TÍNH TOÁN GIỜ LÀM VÀ TĂNG CA THEO CA ---
const calculateWorkHours = (checkIn, checkOut, workDate, shiftType) => {
    const results = { TOTH: 0, H1: 0, H2: 0, H3: 0, B3: 0, normalHours: 0 };
    if (!checkIn || !checkOut || !shiftType || !SHIFT_CONFIG[shiftType]) return results;
    const start = new Date(checkIn);
    const end = new Date(checkOut);
    if (end <= start) return results;
    const shift = SHIFT_CONFIG[shiftType];
    let totalMinutes = (end - start) / (1000 * 60);
    totalMinutes -= shift.breakMinutes;
    results.TOTH = parseFloat((totalMinutes / 60).toFixed(2));
    let nightMinutes = 0;
    const nightStartBoundary = new Date(start.toDateString() + ' 22:00:00');
    const nightEndBoundary = new Date(start.toDateString() + ' 06:00:00');
    if (shiftType === 'CADEM') { nightEndBoundary.setDate(nightEndBoundary.getDate() + 1); }
    for (let time = new Date(start); time < end; time.setMinutes(time.getMinutes() + 1)) {
        const isNightHour = (time >= nightStartBoundary) || (time < nightEndBoundary && time.getDate() === nightEndBoundary.getDate());
        if (isNightHour) { nightMinutes++; }
    }
    results.B3 = parseFloat((nightMinutes / 60).toFixed(2));
    if (results.TOTH > shift.standardHours) {
        results.normalHours = shift.standardHours;
        const overtimeHours = results.TOTH - shift.standardHours;
        const dayOfWeek = workDate.getDay();
        const dateString = workDate.toISOString().split('T')[0];
        if (PUBLIC_HOLIDAYS[dateString]) { results.H3 = overtimeHours; } 
        else if (dayOfWeek === 0) { results.H2 = overtimeHours; } 
        else { results.H1 = overtimeHours; }
    } else {
        results.normalHours = results.TOTH;
    }
    return results;
};

// API Endpoint để tính toán thử
exports.calculateHoursPreview = async (req, res) => {
    const { checkInTime, checkOutTime, workDate, shiftType } = req.body;
    if (!checkInTime || !checkOutTime || !workDate || !shiftType) {
        return res.status(400).json({ message: "Thiếu thông tin giờ vào, giờ ra, ngày làm việc hoặc ca." });
    }
    
    let checkIn = new Date(`${workDate}T${checkInTime}`);
    let checkOut = new Date(`${workDate}T${checkOutTime}`);
    
    // Xử lý trường hợp qua ngày (đặc biệt cho ca đêm)
    if (shiftType === 'CADEM' && checkOut < checkIn) {
        checkOut.setDate(checkOut.getDate() + 1);
    }

    const results = calculateWorkHours(checkIn, checkOut, new Date(workDate), shiftType);
    res.status(200).json(results);
};

// Chức năng Lấy GROUPID VÀ INDAT và OUTDAT của tất cả nhân viên
exports.getAllEmployeeGroupIDs = async (req, res) => {
    try {
        await poolConnect;
        const result = await pool.request()
            .query(`
                SELECT 
                    f.EMPID,
                    f.EMPNAM_VN,
                    f.INDAT,
                    f.OUTDAT,       -- <-- THÊM CỘT OUTDAT BỊ THIẾU
                    b.b_groupid AS GROUPID,
                    b.b_whsno AS WHSNO
                FROM EMPFILE f
                LEFT JOIN EMPFILEB b ON f.EMPID = b.EMPID
                WHERE f.STATUS IS NULL OR f.STATUS != 'Q'
            `);
        
        res.status(200).json(result.recordset);

    } catch (err) {
        console.error("Lỗi khi lấy danh sách chi tiết nhân viên:", err);
        res.status(500).json({ message: 'Lỗi server khi lấy dữ liệu nhân viên.' });
    }
};

// Chức năng upload chấm công đã được tối ưu
exports.uploadTimesheet = async (req, res) => {
    const records = req.body;
    if (!Array.isArray(records) || records.length === 0) {
        return res.status(400).json({ message: "Dữ liệu upload không hợp lệ." });
    }

    const table = new sql.Table('#TempTimesheet'); 
    table.create = true; 

    table.columns.add('empid', sql.VarChar(10), { nullable: false });
    table.columns.add('workdat', sql.VarChar(10), { nullable: false });
    table.columns.add('timeup', sql.VarChar(6), { nullable: true });
    table.columns.add('timedown', sql.VarChar(6), { nullable: true });
    table.columns.add('TOTH', sql.Decimal(18, 1), { nullable: true });
    table.columns.add('FORGET', sql.Int, { nullable: true });
    table.columns.add('Latefor', sql.Int, { nullable: true });
    table.columns.add('Kzhour', sql.Decimal(18, 1), { nullable: true });
    table.columns.add('H1', sql.Real, { nullable: true });
    table.columns.add('H2', sql.Real, { nullable: true });
    table.columns.add('H3', sql.Real, { nullable: true });
    table.columns.add('B3', sql.Real, { nullable: true });
    table.columns.add('B4', sql.Real, { nullable: true });
    table.columns.add('B5', sql.Real, { nullable: true });
    table.columns.add('BC', sql.VarChar(10), { nullable: true });
    table.columns.add('yymm', sql.VarChar(6), { nullable: true });
    table.columns.add('flag', sql.VarChar(10), { nullable: true });
    table.columns.add('muser', sql.VarChar(10), { nullable: true });
    table.columns.add('EMPWHSNO', sql.VarChar(10), { nullable: true });
    table.columns.add('indat', sql.DateTime, { nullable: true });
    table.columns.add('outdat', sql.DateTime, { nullable: true });
    table.columns.add('memo', sql.NVarChar(100), { nullable: true });

    for (const record of records) {
        table.rows.add(
            record.empid, record.workdat, record.timeup || null, record.timedown || null,
            record.TOTH || 0, record.FORGET || 0, record.Latefor || 0, record.Kzhour || 0,
            record.H1 || 0, record.H2 || 0, record.H3 || 0, record.B3 || 0, record.B4 || 0, record.B5 || 0,
            record.BC || null, record.yymm || null, record.flag || 'Auto', record.muser || 'Step2',
            record.EMPWHSNO || 'LT', record.indat ? new Date(record.indat) : null, record.outdat ? new Date(record.outdat) : null,
            record.memo || null
        );
    }

    const transaction = new sql.Transaction(pool);
    try {
        await poolConnect;
        await transaction.begin();

        await transaction.request().query(`
            CREATE TABLE #TempTimesheet (
                empid VARCHAR(10) NOT NULL, workdat VARCHAR(10) NOT NULL,
                timeup VARCHAR(6), timedown VARCHAR(6), TOTH DECIMAL(18, 1),
                FORGET INT, Latefor INT, Kzhour DECIMAL(18, 1),
                H1 REAL, H2 REAL, H3 REAL, B3 REAL, B4 REAL, B5 REAL,
                BC VARCHAR(10), yymm VARCHAR(6), flag VARCHAR(10), muser VARCHAR(10),
                EMPWHSNO VARCHAR(10), indat DATETIME, outdat DATETIME, memo NVARCHAR(100),
                PRIMARY KEY (empid, workdat)
            );
        `);

        const result = await transaction.request().bulk(table);

        await transaction.request().query(`
            MERGE EMPWORK AS target
            USING #TempTimesheet AS source ON (target.empid = source.empid AND target.workdat = source.workdat)
            WHEN MATCHED THEN
                UPDATE SET
                    target.timeup = source.timeup, target.timedown = source.timedown, target.TOTH = source.TOTH,
                    target.FORGET = source.FORGET, target.Latefor = source.Latefor, target.Kzhour = source.Kzhour,
                    target.H1 = source.H1, target.H2 = source.H2, target.H3 = source.H3,
                    target.B3 = source.B3, target.B4 = source.B4, target.B5 = source.B5,
                    target.BC = source.BC, target.yymm = source.yymm, target.flag = source.flag,
                    target.muser = source.muser, target.MDTM = GETDATE(), target.EMPWHSNO = source.EMPWHSNO,
                    target.indat = source.indat, target.outdat = source.outdat, target.memo = source.memo
            WHEN NOT MATCHED BY TARGET THEN
                INSERT (
                    empid, workdat, timeup, timedown, TOTH, FORGET, Latefor, Kzhour,
                    H1, H2, H3, B3, B4, B5, BC, yymm, flag, muser, MDTM, EMPWHSNO,
                    indat, outdat, memo, JIAA, JIAB, JIAC, JIAD, JIAE, JIAF, JIAG, JIAH
                ) VALUES (
                    source.empid, source.workdat, source.timeup, source.timedown, source.TOTH, source.FORGET, source.Latefor, source.Kzhour,
                    source.H1, source.H2, source.H3, source.B3, source.B4, source.B5, source.BC, source.yymm, source.flag,
                    source.muser, GETDATE(), source.EMPWHSNO, source.indat, source.outdat, source.memo,
                    0, 0, 0, 0, 0, 0, 0, 0
                );
            DROP TABLE #TempTimesheet;
        `);

        await transaction.commit();
        res.status(200).json({
            message: `Hoàn tất xử lý. ${result.rowsAffected} dòng đã được ghi/cập nhật thành công.`,
            successCount: result.rowsAffected,
        });

    } catch (err) {
        console.error("Lỗi khi bulk upload chấm công:", err);
        if (transaction.active) {
            await transaction.rollback();
        }
        res.status(500).json({ message: 'Lỗi server khi upload chấm công.', error: err.message });
    }
};


// Chức năng tra cứu chấm công theo nhân viên và tháng
exports.getTimesheetForEmployee = async (req, res) => {
    const { empid, yymm } = req.params;
    try {
        await poolConnect;
        const result = await pool.request()
            .input('empid', sql.VarChar, empid)
            .input('yymm', sql.VarChar, yymm)
            .query(`SELECT * FROM EMPWORK WHERE empid = @empid AND yymm = @yymm ORDER BY workdat ASC`);
        
        // SỬA LỖI: Xử lý và làm sạch dữ liệu trước khi gửi về frontend
        // Đảm bảo 'autoid' luôn là một giá trị số duy nhất.
        const sanitizedData = result.recordset.map(row => {
            let cleanRow = { ...row };
            let id = cleanRow.autoid || cleanRow.AUTOID;

            if (Array.isArray(id)) {
                console.warn(`[Backend Sanitize] Cảnh báo: 'autoid' là một mảng cho empid=${empid}, workdat=${cleanRow.workdat}. Lấy giá trị đầu tiên: ${id[0]}`);
                cleanRow.autoid = id[0];
            }
            return cleanRow;
        });
        
        res.status(200).json(sanitizedData);
    } catch (err) {
        console.error("Lỗi khi tra cứu chấm công:", err);
        res.status(500).json({ message: 'Lỗi server khi tra cứu chấm công.' });
    }
};

// Chức năng cập nhật một dòng chấm công
exports.updateTimesheetEntry = async (req, res) => {
    const { autoid } = req.params;
    const { timeup, timedown, TOTH, Kzhour, H1, H2, H3, B3, B4, B5, BC, FORGET, Latefor } = req.body;

    // Thêm kiểm tra an toàn cho autoid
    if (!autoid || isNaN(parseInt(autoid))) {
        return res.status(400).json({ message: "ID bản ghi (autoid) không hợp lệ."});
    }

    try {
        await poolConnect;
        await pool.request()
            .input('autoid', sql.Int, parseInt(autoid)) // Chuyển đổi sang Int để chắc chắn
            .input('timeup', sql.VarChar, timeup)
            .input('timedown', sql.VarChar, timedown)
            .input('TOTH', sql.Decimal(18, 1), TOTH)
            .input('Kzhour', sql.Decimal(18, 1), Kzhour)
            .input('H1', sql.Real, H1)
            .input('H2', sql.Real, H2)
            .input('H3', sql.Real, H3)
            .input('B3', sql.Real, B3)
            .input('B4', sql.Real, B4)
            .input('B5', sql.Real, B5)
            .input('BC', sql.VarChar(10), BC)
            .input('FORGET', sql.Int, FORGET)
            .input('Latefor', sql.Int, Latefor)
            .query(`
                UPDATE EMPWORK SET
                    timeup = @timeup, timedown = @timedown, TOTH = @TOTH, Kzhour = @Kzhour,
                    H1 = @H1, H2 = @H2, H3 = @H3, B3 = @B3, B4 = @B4, B5 = @B5,
                    BC = @BC, FORGET = @FORGET, Latefor = @Latefor,
                    muser = 'Auto', MDTM = GETDATE()
                WHERE autoid = @autoid
            `);
        res.status(200).json({ message: 'Cập nhật thành công.' });
    } catch (err) {
        console.error("Lỗi khi cập nhật chấm công:", err);
        res.status(500).json({ message: 'Lỗi server khi cập nhật chấm công.' });
    }
};

// Chức năng lấy bảng tổng hợp chấm công theo tháng (có lọc theo bộ phận)
exports.getMonthlyTimesheetSummary = async (req, res) => {
    const { yymm } = req.params;
    const { groupId = 'ALL' } = req.query; 

    if (!yymm || yymm.length !== 6) {
        return res.status(400).json({ message: 'Định dạng tháng không hợp lệ. Yêu cầu: YYYYMM' });
    }

    try {
        await poolConnect;
        const request = pool.request();
        request.input('yymm', sql.VarChar, yymm);

        // SỬA ĐỔI: Thêm điều kiện lọc nhân viên đã nghỉ việc (e.OUTDAT)
        let query = `
            WITH TimesheetSummary AS (
                SELECT 
                    t.empid,
                    e.EMPNAM_VN AS empName,
                    b.b_groupid,
                    SUM(ISNULL(t.Kzhour, 0)) AS total_Kzhour,
                    SUM(ISNULL(t.TOTH, 0)) AS total_TOTH,
                    SUM(ISNULL(t.H1, 0)) AS total_H1,
                    SUM(ISNULL(t.H2, 0)) AS total_H2,
                    SUM(ISNULL(t.H3, 0)) AS total_H3,
                    SUM(ISNULL(t.B3, 0)) AS total_B3,
                    SUM(ISNULL(t.B4, 0)) AS total_B4,
                    SUM(ISNULL(t.B5, 0)) AS total_B5
                FROM EMPWORK t
                LEFT JOIN EMPFILE e ON t.empid = e.EMPID
                LEFT JOIN EMPFILEB b ON t.empid = b.EMPID
                WHERE t.yymm = @yymm 
                  AND (e.OUTDAT IS NULL OR e.OUTDAT >= CONVERT(date, @yymm + '01', 112))
                GROUP BY t.empid, e.EMPNAM_VN, b.b_groupid
            ),
            LeaveSummary AS (
                SELECT
                    EMPID,
                    SUM(ISNULL(HHour, 0)) AS total_leave_hours,
                    STUFF(
                        (SELECT DISTINCT ', ' + 
                            CASE RTRIM(JiaType)
                                WHEN 'A' THEN N'Việc riêng'
                                WHEN 'B' THEN N'Phép Bệnh'
                                WHEN 'C' THEN N'Nghỉ kết hôn'
                                WHEN 'D' THEN N'Phép Tang'
                                WHEN 'E' THEN N'Phép năm'
                                WHEN 'F' THEN N'Nghỉ thai sản'
                                WHEN 'G' THEN N'Nghỉ công tác'
                                WHEN 'H' THEN N'Nghỉ C.Thường'
                                WHEN 'I' THEN N'Đi đường'
                                WHEN 'K' THEN N'Không lương'
                                ELSE RTRIM(JiaType)
                            END
                         FROM EMPHOLIDAY
                         WHERE EMPID = h.EMPID AND FORMAT(DateUP, 'yyyyMM') = @yymm
                         FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 2, ''
                    ) AS leave_types
                FROM EMPHOLIDAY h
                WHERE FORMAT(DateUP, 'yyyyMM') = @yymm
                GROUP BY EMPID
            )
            SELECT
                ts.empid,
                ts.empName,
                ts.total_Kzhour,
                ts.total_TOTH,
                ts.total_H1,
                ts.total_H2,
                ts.total_H3,
                ts.total_B3,
                ts.total_B4,
                ts.total_B5,
                ISNULL(ls.total_leave_hours, 0) AS total_leave_hours,
                ls.leave_types
            FROM TimesheetSummary ts
            LEFT JOIN LeaveSummary ls ON ts.empid = ls.EMPID
        `;

        if (groupId && groupId !== 'ALL') {
            query += ` WHERE ts.b_groupid = @groupId`;
            request.input('groupId', sql.VarChar, groupId);
        }

        query += ` ORDER BY ts.empid`;
        
        const result = await request.query(query);
        res.status(200).json(result.recordset);

    } catch (err) {
        console.error("Lỗi khi lấy bảng tổng hợp chấm công:", err);
        res.status(500).json({ message: 'Lỗi server khi lấy bảng tổng hợp chấm công.', error: err.message });
    }
};