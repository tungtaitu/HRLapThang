/*
 * File: controllers/admin.controller.js
 * Mô tả: Chứa logic xử lý cho các route của admin.
 */
const sql = require('mssql');
const fs = require('fs').promises;
const xlsx = require('xlsx');
const { pool, poolConnect } = require('../config/db');
const { jsonDataState, writeJsonAndUpdateState, LEAVE_DATA_FILE, PAYROLL_APPROVAL_FILE, USER_PASSWORDS_FILE, LUONG_T13_DATA_FILE } = require('../services/json.service');
const { getLeaveSummary } = require('../services/leave.service');
const { calculateLeaveHours } = require('../utils/calculate');
const { parseDate } = require('../utils/helpers');
const { PUBLIC_HOLIDAYS } = require('../constants/holidays');

exports.resetPassword = async (req, res) => {
    const { empid } = req.body;
    if (!empid) {
        return res.status(400).json({ message: 'Vui lòng cung cấp Mã số nhân viên.' });
    }

    try {
        const currentPasswords = jsonDataState.userPasswords;
        const userExistsInFile = currentPasswords.some(user => user.empid === empid);

        if (!userExistsInFile) {
            return res.status(200).json({ message: `Nhân viên ${empid} đang dùng mật khẩu mặc định. Không cần reset.` });
        }

        const updatedPasswords = currentPasswords.filter(user => user.empid !== empid);
        await writeJsonAndUpdateState(USER_PASSWORDS_FILE, updatedPasswords);

        res.status(200).json({ message: `Đã reset mật khẩu cho nhân viên ${empid} về mặc định (ngày sinh).` });

    } catch (err) {
        console.error(`Lỗi khi reset mật khẩu cho ${empid}:`, err);
        res.status(500).json({ message: 'Lỗi server khi reset mật khẩu.' });
    }
};

exports.getEmployeeInfo = async (req, res) => {
    const { empid } = req.params;
    if (!empid) {
        return res.status(400).json({ message: 'Vui lòng cung cấp Mã số nhân viên.' });
    }

    try {
        await poolConnect;
        const result = await pool.request()
            .input('empid', sql.VarChar, empid)
            .query(`
                SELECT TOP 1
                    f.EMPNAM_VN as name,
                    g.SYS_VALUE as department
                FROM
                    EMPFILE f
                LEFT JOIN (
                    SELECT TOP 1 EMPID, GROUPID
                    FROM EMPDSALARY
                    WHERE EMPID = @empid
                    ORDER BY YYMM DESC
                ) latest_sal ON f.EMPID = latest_sal.EMPID
                LEFT JOIN BASICCODE g ON latest_sal.GROUPID = g.SYS_TYPE AND g.FUNC = 'GROUPID'
                WHERE f.EMPID = @empid
            `);

        if (result.recordset.length > 0) {
            res.status(200).json(result.recordset[0]);
        } else {
            res.status(404).json({ message: 'Không tìm thấy nhân viên.' });
        }
    } catch (err) {
        console.error(`Lỗi khi lấy thông tin nhân viên ${empid}:`, err);
        res.status(500).json({ message: 'Lỗi server khi lấy thông tin nhân viên.' });
    }
};

exports.uploadLeaveFile = async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ message: 'Không có file nào được upload.' });
        }

        const workbook = xlsx.readFile(req.file.path);
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const newDataRaw = xlsx.utils.sheet_to_json(sheet);
        
        const newData = newDataRaw.map(row => {
            const standardizedRow = { isForeigner: false };
            let msnvValue = null;
            for (const key in row) {
                const trimmedKey = key.trim().toUpperCase();
                if (trimmedKey === 'MSNV') {
                    standardizedRow.MSNV = String(row[key]);
                    msnvValue = standardizedRow.MSNV;
                } else if (trimmedKey === 'MONTH') {
                    standardizedRow.Month = row[key];
                } else if (trimmedKey === 'PHEP') {
                    const phepValue = row[key];
                    standardizedRow.PHEP = (phepValue === '-' || phepValue === null || phepValue === undefined) ? 0 : parseFloat(phepValue) || 0;
                } else if (trimmedKey === 'NUOCNGOAI' || trimmedKey === 'FOREIGNER') {
                    const foreignerValue = String(row[key]).toLowerCase();
                    standardizedRow.isForeigner = ['x', 'yes', 'co', 'true', '1'].includes(foreignerValue);
                }
            }
            if (msnvValue) {
                return standardizedRow;
            }
            return null;
        }).filter(Boolean);

        if (newData.length === 0) {
             await fs.unlink(req.file.path);
             return res.status(400).json({ message: 'File Excel không có dữ liệu hợp lệ hoặc thiếu cột MSNV.' });
        }

        const existingData = jsonDataState.leaveData;

        const dataMap = new Map(existingData.map(item => [item.MSNV, item]));
        newData.forEach(newItem => {
            dataMap.set(newItem.MSNV, newItem);
        });
        const mergedData = Array.from(dataMap.values());

        await writeJsonAndUpdateState(LEAVE_DATA_FILE, mergedData);

        await fs.unlink(req.file.path);

        res.status(200).json({ message: `Cập nhật thành công! ${newData.length} bản ghi đã được xử lý.` });

    } catch (error) {
        console.error("Lỗi khi upload file:", error);
        if (req.file && req.file.path) {
            try {
                await fs.unlink(req.file.path);
            } catch (unlinkError) {
                console.error("Lỗi khi xóa file tạm:", unlinkError);
            }
        }
        res.status(500).json({ message: 'Có lỗi xảy ra khi xử lý file.' });
    }
};

exports.exportLeaveFile = async (req, res) => {
    const { year } = req.params;
    const { groupId } = req.query;

    try {
        await poolConnect;
        let employeeQuery = `
            SELECT
                f.EMPID,
                f.EMPNAM_VN,
                ISNULL(latest_sal.GROUPID, 'N/A') as GROUPID,
                ISNULL(g.SYS_VALUE, 'Chưa phân loại') as GroupName
            FROM
                EMPFILE f
            OUTER APPLY (
                SELECT TOP 1 s.GROUPID
                FROM EMPDSALARY s
                WHERE s.EMPID = f.EMPID
                ORDER BY s.YYMM DESC
            ) latest_sal
            LEFT JOIN BASICCODE g ON latest_sal.GROUPID = g.SYS_TYPE AND g.FUNC = 'GROUPID'
            WHERE (f.STATUS IS NULL OR f.STATUS != 'Q') AND f.OUTDAT IS NULL
        `;

        if (groupId && groupId !== 'ALL') {
            employeeQuery += ` AND latest_sal.GROUPID = @groupId`;
        }
        employeeQuery += ` ORDER BY GROUPID, f.EMPID`;

        const request = pool.request();
        if (groupId && groupId !== 'ALL') {
            request.input('groupId', sql.VarChar, groupId);
        }

        const employeesResult = await request.query(employeeQuery);
        const employees = employeesResult.recordset;

        const leaveDataForExport = [];
        for (const emp of employees) {
            const summary = await getLeaveSummary(pool, emp.EMPID, year);
            const employeeConfig = jsonDataState.leaveData.find(e => e.MSNV === emp.EMPID);
            const isForeigner = employeeConfig?.isForeigner || false;

            leaveDataForExport.push({
                'Mã Bộ Phận': emp.GROUPID,
                'Tên Bộ Phận': emp.GroupName,
                'Mã Nhân Viên': emp.EMPID,
                'Tên Nhân Viên': emp.EMPNAM_VN,
                'Đối Tượng': isForeigner ? 'Nước ngoài' : 'Trong nước',
                'Số Giờ Phép Năm Còn Lại': summary.remaining
            });
        }

        const worksheet = xlsx.utils.json_to_sheet(leaveDataForExport);
        worksheet['!cols'] = [ { wch: 15 }, { wch: 25 }, { wch: 15 }, { wch: 30 }, { wch: 15 }, { wch: 30 } ];
        const workbook = xlsx.utils.book_new();
        xlsx.utils.book_append_sheet(workbook, worksheet, `PhepNam${year}`);

        const buffer = xlsx.write(workbook, { bookType: 'xlsx', type: 'buffer' });
        res.setHeader('Content-Disposition', `attachment; filename=TongHopPhepNam_${year}${groupId && groupId !== 'ALL' ? '_' + groupId : ''}.xlsx`);
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.send(buffer);

    } catch (err) {
        console.error("Lỗi khi xuất file Excel phép năm:", err);
        res.status(500).json({ message: 'Lỗi server khi xuất file.' });
    }
};

exports.getAllPayrolls = async (req, res) => {
    const { yearMonth } = req.params;
    const { groupId } = req.query;
    const yearMonthFormatted = yearMonth.replace('-', '');
    try {
        await poolConnect;
        let query = `
            SELECT
                f.EMPID,
                f.EMPNAM_VN,
                s.REAL_TOTAL,
                s.GROUPID,
                g.SYS_VALUE as GroupName
            FROM EMPFILE f
            JOIN EMPDSALARY s ON f.EMPID = s.EMPID
            LEFT JOIN BASICCODE g ON s.GROUPID = g.SYS_TYPE AND g.FUNC = 'GROUPID'
            WHERE s.YYMM = @yymm_param
        `;

        if (groupId && groupId !== 'ALL') {
            query += ` AND s.GROUPID = @groupId`;
        }
        query += ` ORDER BY s.GROUPID, f.EMPID`;

        const request = pool.request();
        request.input('yymm_param', sql.VarChar, yearMonthFormatted);
        if (groupId && groupId !== 'ALL') {
            request.input('groupId', sql.VarChar, groupId);
        }

        const result = await request.query(query);
        const approvals = jsonDataState.payrollApprovals;
        const isApproved = approvals.includes(yearMonth);
        res.status(200).json({ payrolls: result.recordset, isApproved });
    } catch (err) {
        console.error("Lỗi API All Payrolls:", err);
        res.status(500).json({ message: 'Lỗi server khi lấy dữ liệu lương.' });
    }
};

exports.approvePayroll = async (req, res) => {
    const { yearMonth } = req.body;
    try {
        const approvals = [...jsonDataState.payrollApprovals];
        if (!approvals.includes(yearMonth)) {
            approvals.push(yearMonth);
            await writeJsonAndUpdateState(PAYROLL_APPROVAL_FILE, approvals);
        }
        res.status(200).json({ message: `Đã phê duyệt thành công lương tháng ${yearMonth}.` });

    } catch (err) {
        console.error("Lỗi API Approve Payroll:", err);
        res.status(500).json({ message: 'Lỗi server khi phê duyệt.' });
    }
};

exports.exportPayrolls = async (req, res) => {
    const { yearMonth } = req.params;
    const { groupId } = req.query;
    const yearMonthFormatted = yearMonth.replace('-', '');

    try {
        await poolConnect;

        let query = `
            SELECT
                s.GROUPID as 'Mã Bộ Phận',
                g.SYS_VALUE as 'Tên Bộ Phận',
                f.EMPID as 'Mã Nhân Viên',
                f.EMPNAM_VN as 'Tên Nhân Viên',
                s.REAL_TOTAL as 'Lương Thực Lãnh'
            FROM EMPFILE f
            JOIN EMPDSALARY s ON f.EMPID = s.EMPID
            LEFT JOIN BASICCODE g ON s.GROUPID = g.SYS_TYPE AND g.FUNC = 'GROUPID'
            WHERE s.YYMM = @yymm_param
        `;

        if (groupId && groupId !== 'ALL') {
            query += ` AND s.GROUPID = @groupId`;
        }

        query += ` ORDER BY s.GROUPID, f.EMPID`;

        const request = pool.request();
        request.input('yymm_param', sql.VarChar, yearMonthFormatted);
        if (groupId && groupId !== 'ALL') {
            request.input('groupId', sql.VarChar, groupId);
        }

        const result = await request.query(query);

        const worksheet = xlsx.utils.json_to_sheet(result.recordset);
        worksheet['!cols'] = [ { wch: 15 }, { wch: 25 }, { wch: 15 }, { wch: 30 }, { wch: 20 } ];
        const workbook = xlsx.utils.book_new();
        xlsx.utils.book_append_sheet(workbook, worksheet, `BangLuong_${yearMonth}`);

        const buffer = xlsx.write(workbook, { bookType: 'xlsx', type: 'buffer' });
        res.setHeader('Content-Disposition', `attachment; filename=BangLuong_${yearMonth}${groupId && groupId !== 'ALL' ? '_' + groupId : ''}.xlsx`);
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.send(buffer);

    } catch (err) {
        console.error("Lỗi khi xuất file bảng lương:", err);
        res.status(500).json({ message: 'Lỗi server khi xuất file bảng lương.' });
    }
};

exports.getGroups = async (req, res) => {
    try {
        await poolConnect;
        const result = await pool.request()
            .query`SELECT SYS_TYPE as groupId, SYS_VALUE as groupName FROM BASICCODE WHERE FUNC = 'GROUPID' ORDER BY SYS_TYPE`;
        res.status(200).json(result.recordset);
    } catch (err) {
        console.error("Lỗi khi lấy danh sách bộ phận:", err);
        res.status(500).json({ message: 'Lỗi server khi lấy danh sách bộ phận.' });
    }
};

exports.exportTimesheet = async (req, res) => {
    const { yearMonth } = req.params;
    const { groupId } = req.query;
    const yearMonthFormatted = yearMonth.replace('-', '');
    const year = parseInt(yearMonth.substring(0, 4));
    const month = parseInt(yearMonth.substring(5, 7));

    try {
        await poolConnect;

        let query = `
            WITH PayableDays AS (
                SELECT
                    EmpID,
                    CAST(workdat AS DATE) as PayableDate
                FROM
                    EMPWORK
                WHERE
                    LEFT(workdat, 6) = @yearMonthFormatted
                    AND (DATEDIFF(day, '17530101', CAST(workdat AS DATE)) % 7) <> 6
                    AND timeup IS NOT NULL AND timeup <> '000000' AND timeup <> '0'
                UNION
                SELECT
                    empid,
                    CAST(DateUP AS DATE) as PayableDate
                FROM
                    EMPHOLIDAY
                WHERE
                    YEAR(DateUP) = @year
                    AND MONTH(DateUP) = @month
                    AND JiaType = 'E'
            ),
            AggregatedWorkData AS (
                SELECT
                    EmpID,
                    COUNT(PayableDate) as WorkDays
                FROM PayableDays
                GROUP BY EmpID
            ),
            AggregatedHours AS (
                SELECT
                    EmpID,
                    SUM(ISNULL(TOTH, 0)) as TotalHours,
                    SUM(ISNULL(H1, 0)) as TotalH1,
                    SUM(ISNULL(H2, 0)) as TotalH2,
                    SUM(ISNULL(H3, 0)) as TotalH3,
                    SUM(ISNULL(B3, 0)) as TotalB3,
                    SUM(ISNULL(B4, 0)) as TotalB4
                FROM
                    EMPWORK
                WHERE
                    LEFT(workdat, 6) = @yearMonthFormatted
                GROUP BY
                    EmpID
            ),
            EmployeeDepartment AS (
                SELECT
                    f.EMPID,
                    f.EMPNAM_VN,
                    latest_sal.GROUPID
                FROM
                    EMPFILE f
                LEFT JOIN (
                    SELECT EMPID, MAX(YYMM) as MaxYYMM
                    FROM EMPDSALARY
                    GROUP BY EMPID
                ) latest_sal_yymm ON f.EMPID = latest_sal_yymm.EMPID
                LEFT JOIN EMPDSALARY latest_sal ON latest_sal_yymm.EMPID = latest_sal.EMPID AND latest_sal_yymm.MaxYYMM = latest_sal.YYMM
                WHERE (f.STATUS IS NULL OR f.STATUS != 'Q') AND f.OUTDAT IS NULL
            ),
            LeaveData AS (
                SELECT
                    empid,
                    SUM(CASE WHEN JiaType = 'E' THEN ISNULL(HHour, 0) ELSE 0 END) as PhepNam,
                    SUM(CASE WHEN JiaType = 'A' THEN ISNULL(HHour, 0) ELSE 0 END) as ViecRieng,
                    SUM(CASE WHEN JiaType = 'B' THEN ISNULL(HHour, 0) ELSE 0 END) as PhepBenh,
                    SUM(CASE WHEN JiaType = 'D' THEN ISNULL(HHour, 0) ELSE 0 END) as PhepTang
                FROM
                    EMPHOLIDAY
                WHERE
                    YEAR(DateUP) = @year AND MONTH(DateUP) = @month
                GROUP BY
                    empid
            )
            SELECT
                ed.GROUPID as 'Mã Bộ Phận',
                g.SYS_VALUE as 'Tên Bộ Phận',
                ed.EMPID as 'Mã Nhân Viên',
                ed.EMPNAM_VN as 'Tên Nhân Viên',
                ISNULL(awd.WorkDays, 0) as 'Số Ngày Làm Việc',
                ISNULL(ah.TotalHours, 0) as 'Tổng Giờ Làm',
                ISNULL(ah.TotalH1, 0) as 'Tăng Ca 1.5',
                ISNULL(ah.TotalH2, 0) as 'Tăng Ca 2.0',
                ISNULL(ah.TotalH3, 0) as 'Tăng Ca 3.0',
                ISNULL(ah.TotalB3, 0) as 'Tăng Ca Đêm',
                ISNULL(ah.TotalB4, 0) as 'Phụ Cấp 0.5',
                ISNULL(l.PhepNam, 0) as 'Phép Năm (giờ)',
                ISNULL(l.ViecRieng, 0) as 'Việc Riêng (giờ)',
                ISNULL(l.PhepBenh, 0) as 'Phép Bệnh (giờ)',
                ISNULL(l.PhepTang, 0) as 'Phép Tang (giờ)'
            FROM
                EmployeeDepartment ed
            LEFT JOIN AggregatedWorkData awd ON ed.EMPID = awd.EmpID
            LEFT JOIN AggregatedHours ah ON ed.EMPID = ah.EmpID
            LEFT JOIN LeaveData l ON ed.EMPID = l.empid
            LEFT JOIN BASICCODE g ON ed.GROUPID = g.SYS_TYPE AND g.FUNC = 'GROUPID'
        `;

        if (groupId && groupId !== 'ALL') {
            query += ` WHERE ed.GROUPID = @groupId`;
        }

        query += ` ORDER BY ed.GROUPID, ed.EMPID`;

        const request = pool.request();
        request.input('yearMonthFormatted', sql.VarChar, yearMonthFormatted);
        request.input('year', sql.Int, year);
        request.input('month', sql.Int, month);

        if (groupId && groupId !== 'ALL') {
            request.input('groupId', sql.VarChar, groupId);
        }

        const result = await request.query(query);

        const worksheet = xlsx.utils.json_to_sheet(result.recordset);
        worksheet['!cols'] = [
            { wch: 15 }, { wch: 25 }, { wch: 15 }, { wch: 30 },
            { wch: 18 }, { wch: 15 }, { wch: 15 }, { wch: 15 },
            { wch: 15 }, { wch: 15 }, { wch: 15 }, { wch: 18 },
            { wch: 18 }, { wch: 18 }, { wch: 18 }
        ];
        const workbook = xlsx.utils.book_new();
        xlsx.utils.book_append_sheet(workbook, worksheet, `ChamCong_${yearMonth}`);

        const buffer = xlsx.write(workbook, { bookType: 'xlsx', type: 'buffer' });
        res.setHeader('Content-Disposition', `attachment; filename=BaoCaoChamCong_${yearMonth}${groupId && groupId !== 'ALL' ? '_' + groupId : ''}.xlsx`);
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.send(buffer);

    } catch (err) {
        console.error("Lỗi khi xuất file chấm công:", err);
        res.status(500).json({ message: 'Lỗi server khi xuất file chấm công.' });
    }
};

exports.uploadLuongT13 = async (req, res) => {
    const { year } = req.body;
    if (!req.file || !year) return res.status(400).json({ message: 'Vui lòng cung cấp file Excel và năm áp dụng.' });

    try {
        const workbook = xlsx.readFile(req.file.path);
        let jsonData = [];
        let sheetFound = false;
        for (const sheetName of workbook.SheetNames) {
            const sheet = workbook.Sheets[sheetName];
            if (!sheet) continue;
            const rawData = xlsx.utils.sheet_to_json(sheet, { header: 1, defval: '' });
            let headerRowIndex = rawData.findIndex(row => Array.isArray(row) && row.some(cell => typeof cell === 'string' && cell.trim().toLowerCase() === 'msnv'));
            if (headerRowIndex !== -1) {
                const header = rawData[headerRowIndex].map(h => typeof h === 'string' ? h.trim().toLowerCase() : '');
                const dataRows = rawData.slice(headerRowIndex + 1);
                jsonData = dataRows.map(row => {
                    const obj = {};
                    header.forEach((key, i) => { if (key) obj[key] = row[i]; });
                    return obj;
                });
                sheetFound = true;
                break;
            }
        }
        if (!sheetFound) {
            await fs.unlink(req.file.path);
            return res.status(400).json({ message: "Không tìm thấy sheet nào trong file Excel chứa cột 'msnv'." });
        }

        const processedData = jsonData.map(row => {
            const getNumericValue = (key) => {
                const value = row[key];
                if (value === undefined || value === null || value === '') return 0;
                return parseFloat(String(value).replace(/,/g, '')) || 0;
            };
            const getStringValue = (key) => (row[key] === undefined || row[key] === null) ? '' : String(row[key]);

            const employeeData = {
                MSNV: getStringValue('msnv'), HoTen: getStringValue('ho_ten'), NgayVaoLam: getStringValue('ngay_vao_lam'),
                LuongCoBan: getNumericValue('luong_cb'), PhuCapChucVu: getNumericValue('pc_chuc_vu'),
                PhuCapKyThuat: getNumericValue('pc_ky_thuat'), PhuCapMoiTruong: getNumericValue('pc_moi_truong'),
                PhuCapXangXe: getNumericValue('pc_xang_xe'), PhuCapDienThoai: getNumericValue('pc_dien_thoai'),
                PhuCapNhaO: getNumericValue('pc_nha_o'), ChuyenCan: getNumericValue('chuyen_can'),
                ThuNhapKhac: getNumericValue('thu_nhap_khac'), TienThuongThang13: getNumericValue('tien_thuong_t13'),
                TienPhepNam: getNumericValue('tien_phep_nam'), TienTruKhac: getNumericValue('tien_tru_khac'),
                HeSoThuong: getNumericValue('he_so_thuong'), SoTiengPhepNamConLai: getNumericValue('so_tieng_phep_nam'),
                SoNgayCongThucTe: getNumericValue('so_ngay_cong'), SoNgayNghiKhongLuong: getNumericValue('so_ngay_nghi_kl'),
                SoLanBiBienBan: getNumericValue('so_lan_bien_ban'), ChucVu: ''
            };

            const chucVuValue = getStringValue('pc_chuc_vu');
            if (isNaN(parseFloat(chucVuValue))) {
                employeeData.ChucVu = chucVuValue;
                employeeData.PhuCapChucVu = 0;
            }

            employeeData.TienThuongThang13 = Math.round(employeeData.TienThuongThang13);
            employeeData.TienPhepNam = Math.round(employeeData.TienPhepNam);
            employeeData.TienTruKhac = Math.round(employeeData.TienTruKhac);
            employeeData.TongLuong = Math.round(employeeData.LuongCoBan + employeeData.PhuCapChucVu + employeeData.PhuCapKyThuat + employeeData.PhuCapMoiTruong + employeeData.PhuCapXangXe + employeeData.PhuCapDienThoai + employeeData.PhuCapNhaO + employeeData.ChuyenCan + employeeData.ThuNhapKhac);
            employeeData.TongCong = Math.round(employeeData.TienThuongThang13 + employeeData.TienPhepNam);
            employeeData.ThucLanh = Math.round(employeeData.TongCong - employeeData.TienTruKhac);

            if (!employeeData.MSNV) return null;
            return employeeData;
        }).filter(Boolean);

        if (processedData.length === 0) return res.status(400).json({ message: "Không có dữ liệu nhân viên hợp lệ trong file." });

        const allData = JSON.parse(JSON.stringify(jsonDataState.luongT13Data));
        allData[year] = processedData;
        await writeJsonAndUpdateState(LUONG_T13_DATA_FILE, allData);
        await fs.unlink(req.file.path);
        res.status(200).json({ message: `Cập nhật thành công lương tháng 13 cho năm ${year} với ${processedData.length} nhân viên.` });
    } catch (error) {
        console.error("Lỗi khi upload file lương T13:", error);
        if (req.file?.path) { try { await fs.unlink(req.file.path); } catch (e) {} }
        res.status(500).json({ message: 'Có lỗi xảy ra phía server khi xử lý file.' });
    }
};
// --- HÀM MỚI: LẤY DỮ LIỆU CHO DROPDOWN ---
exports.getBasicCodeOptions = async (req, res) => {
    const { func } = req.params;
    if (!func) {
        return res.status(400).json({ message: 'Function code is required.' });
    }
    try {
        await poolConnect;
        const result = await pool.request()
            .input('func', sql.VarChar, func)
            .query(`SELECT SYS_TYPE, SYS_VALUE FROM BASICCODE WHERE FUNC = @func ORDER BY SYS_TYPE`);
        res.status(200).json(result.recordset);
    } catch (err) {
        console.error(`Lỗi khi lấy BASICCODE cho func ${func}:`, err);
        res.status(500).json({ message: 'Lỗi server khi lấy dữ liệu tùy chọn.' });
    }
};