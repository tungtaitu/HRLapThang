/*
 * File: controllers/employee.controller.js
 * Mô tả: Chứa logic xử lý cho các chức năng quản lý nhân viên.
 */
const sql = require('mssql');
const { pool, poolConnect } = require('../config/db');
const { jsonDataState, writeJsonAndUpdateState, USER_PASSWORDS_FILE } = require('../services/json.service');

// Hàm chuyển đổi ngày tháng từ DD/MM/YYYY sang định dạng Date object hoặc null
const parseDateString = (dateStr) => {
    if (!dateStr) return null;
    const parts = dateStr.split('/');
    if (parts.length === 3) {
        const [day, month, year] = parts.map(Number);
        // Date constructor nhận tháng từ 0-11
        const date = new Date(year, month - 1, day);
        // Kiểm tra tính hợp lệ của ngày
        if (!isNaN(date.getTime()) && date.getDate() === day && date.getMonth() === month - 1 && date.getFullYear() === year) {
            return date;
        }
    }
    return null; // Trả về null nếu không hợp lệ
};

// Hàm chuyển đổi giá trị thành chuỗi rỗng nếu là null/undefined, ngược lại giữ nguyên
const toEmptyStringIfNull = (value) => {
    return value === null || value === undefined ? '' : value;
};

// Chức năng Lấy danh sách tất cả nhân viên
exports.getAllEmployees = async (req, res) => {
    try {
        await poolConnect;
        const result = await pool.request()
            .query(`
                SELECT EMPID, EMPNAM_VN, INDAT, OUTDAT, STATUS, PHONE, EMAIL 
                FROM EMPFILE 
                WHERE OUTDAT IS NULL OR STATUS != 'Q'
                ORDER BY EMPID ASC
            `);
        res.status(200).json(result.recordset);
    } catch (err) {
        console.error("Lỗi khi lấy danh sách nhân viên:", err);
        res.status(500).json({ message: 'Lỗi server khi lấy danh sách nhân viên.' });
    }
};

// Chức năng Lấy thông tin chi tiết một nhân viên
exports.getEmployeeInfo = async (req, res) => {
    const { empid } = req.params;
    if (!empid) {
        return res.status(400).json({ message: 'Vui lòng cung cấp Mã số nhân viên.' });
    }
    try {
        await poolConnect;
        // Lấy thông tin từ EMPFILE
        const empFileResult = await pool.request()
            .input('empid', sql.VarChar, empid)
            .query(`SELECT TOP 1 * FROM EMPFILE WHERE EMPID = @empid`);

        if (empFileResult.recordset.length === 0) {
            return res.status(404).json({ message: 'Không tìm thấy nhân viên.' });
        }

        const employeeInfo = empFileResult.recordset[0];

        // Chuyển đổi định dạng ngày tháng sang DD/MM/YYYY cho frontend
        employeeInfo.INDAT = employeeInfo.INDAT ? new Date(employeeInfo.INDAT).toLocaleDateString('en-GB') : '';
        employeeInfo.OUTDAT = employeeInfo.OUTDAT ? new Date(employeeInfo.OUTDAT).toLocaleDateString('en-GB') : '';
        
        // Xử lý ngày sinh: ghép BYY, BMM, BDD thành chuỗi DD/MM/YYYY
        employeeInfo.birthDate = (employeeInfo.BDD && employeeInfo.BMM && employeeInfo.BYY) 
            ? `${String(employeeInfo.BDD).padStart(2, '0')}/${String(employeeInfo.BMM).padStart(2, '0')}/${employeeInfo.BYY}`
            : '';

        // Lấy thông tin từ EMPFILEB
        const empFileBResult = await pool.request()
            .input('empid', sql.VarChar, empid)
            .query(`SELECT TOP 1 * FROM EMPFILEB WHERE EMPID = @empid`);
        
        if (empFileBResult.recordset.length > 0) {
            const empFileBInfo = empFileBResult.recordset[0];
            // Gộp thông tin từ EMPFILEB vào employeeInfo, ánh xạ các trường
            employeeInfo.WHSNO = toEmptyStringIfNull(empFileBInfo.b_whsno);
            employeeInfo.GROUPID = toEmptyStringIfNull(empFileBInfo.b_groupid);
            employeeInfo.ZUNO = toEmptyStringIfNull(empFileBInfo.b_zuno);
            employeeInfo.JOB = toEmptyStringIfNull(empFileBInfo.b_job);
            
            // Ánh xạ các trường khác từ EMPFILEB
            employeeInfo.WKD_No = toEmptyStringIfNull(empFileBInfo.WKD_No);
            employeeInfo.WKD_dueDate = empFileBInfo.WKD_dueDate ? new Date(empFileBInfo.WKD_dueDate).toLocaleDateString('en-GB') : '';
            employeeInfo.experience = toEmptyStringIfNull(empFileBInfo.experience);
            employeeInfo.urgent_person = toEmptyStringIfNull(empFileBInfo.urgent_person);
            employeeInfo.releation = toEmptyStringIfNull(empFileBInfo.releation);
            employeeInfo.urgent_addr = toEmptyStringIfNull(empFileBInfo.urgent_addr);
            employeeInfo.urgent_tel = toEmptyStringIfNull(empFileBInfo.urgent_tel);
            employeeInfo.urgent_mobile = toEmptyStringIfNull(empFileBInfo.urgent_mobile);
            employeeInfo.bh_person = toEmptyStringIfNull(empFileBInfo.bh_person);
            employeeInfo.bh_personID = toEmptyStringIfNull(empFileBInfo.bh_personID);
            employeeInfo.b_shift = toEmptyStringIfNull(empFileBInfo.b_shift);
            employeeInfo.soBH = toEmptyStringIfNull(empFileBInfo.soBH);

        } else {
            // Đảm bảo các trường này tồn tại ngay cả khi không có EMPFILEB record
            employeeInfo.WHSNO = '';
            employeeInfo.GROUPID = '';
            employeeInfo.ZUNO = '';
            employeeInfo.JOB = '';
            employeeInfo.WKD_No = '';
            employeeInfo.WKD_dueDate = '';
            employeeInfo.experience = '';
            employeeInfo.urgent_person = '';
            employeeInfo.releation = '';
            employeeInfo.urgent_addr = '';
            employeeInfo.urgent_tel = '';
            employeeInfo.urgent_mobile = '';
            employeeInfo.bh_person = '';
            employeeInfo.bh_personID = '';
            employeeInfo.b_shift = '';
            employeeInfo.soBH = '';
        }

        // Ánh xạ PASSPORTNO từ DB thành idCardIssueDate cho frontend (PASSPORTNO là NVARCHAR)
        employeeInfo.idCardIssueDate = toEmptyStringIfNull(employeeInfo.PASSPORTNO);
        // Đảm bảo PIssueDate luôn là rỗng theo yêu cầu mới
        employeeInfo.PIssueDate = ''; 

        // Đảm bảo các trường khác của EMPFILE cũng là chuỗi rỗng nếu là null
        employeeInfo.empType = toEmptyStringIfNull(employeeInfo.empType);
        employeeInfo.TX = employeeInfo.TX === null ? 0.0 : employeeInfo.TX; // TX là decimal, mặc định 0.0
        employeeInfo.BHDAT = toEmptyStringIfNull(employeeInfo.BHDAT);
        employeeInfo.GTDAT = toEmptyStringIfNull(employeeInfo.GTDAT);
        employeeInfo.MEMO = toEmptyStringIfNull(employeeInfo.MEMO);
        employeeInfo.photos = null; // photos là image, không xử lý trực tiếp, để null
        employeeInfo.AGES = toEmptyStringIfNull(employeeInfo.AGES); // Lấy giá trị AGE từ DB
        employeeInfo.PDUEDATE = toEmptyStringIfNull(employeeInfo.PDUEDATE);
        employeeInfo.VDUEDATE = toEmptyStringIfNull(employeeInfo.VDUEDATE);
        employeeInfo.StudyJob = toEmptyStringIfNull(employeeInfo.StudyJob);
        employeeInfo.Grps = toEmptyStringIfNull(employeeInfo.Grps);

        employeeInfo.EMPNAM_CN = toEmptyStringIfNull(employeeInfo.EMPNAM_CN);
        employeeInfo.SEX = toEmptyStringIfNull(employeeInfo.SEX);
        employeeInfo.PERSONID = toEmptyStringIfNull(employeeInfo.PERSONID);
        employeeInfo.HOMEADDR = toEmptyStringIfNull(employeeInfo.HOMEADDR);
        employeeInfo.PHONE = toEmptyStringIfNull(employeeInfo.PHONE);
        employeeInfo.MOBILEPHONE = toEmptyStringIfNull(employeeInfo.MOBILEPHONE);
        employeeInfo.EMAIL = toEmptyStringIfNull(employeeInfo.EMAIL);
        employeeInfo.MARRYED = toEmptyStringIfNull(employeeInfo.MARRYED);
        employeeInfo.SCHOOL = toEmptyStringIfNull(employeeInfo.SCHOOL);
        employeeInfo.COUNTRY = toEmptyStringIfNull(employeeInfo.COUNTRY);
        employeeInfo.BANKID = toEmptyStringIfNull(employeeInfo.BANKID);
        employeeInfo.taxCode = toEmptyStringIfNull(employeeInfo.taxCode);
        employeeInfo.VISANO = toEmptyStringIfNull(employeeInfo.VISANO);


        res.status(200).json(employeeInfo);
    } catch (err) {
        console.error(`Lỗi khi lấy thông tin nhân viên ${empid}:`, err);
        res.status(500).json({ message: 'Lỗi server khi lấy thông tin nhân viên.' });
    }
};

// Chức năng Lấy mã nhân viên tiếp theo
exports.getNextEmployeeId = async (req, res) => {
    try {
        await poolConnect;
        const result = await pool.request()
            .query("SELECT MAX(EMPID) as lastId FROM EMPFILE WHERE EMPID LIKE 'LT[0-9][0-9][0-9][0-9]'");
        let nextId = 'LT0001';
        const lastId = result.recordset[0].lastId;
        if (lastId) {
            const lastNumber = parseInt(lastId.substring(2), 10);
            const nextNumberStr = String(lastNumber + 1).padStart(4, '0');
            nextId = `LT${nextNumberStr}`;
        }
        res.status(200).json({ nextId });
    } catch (err) {
        console.error("Lỗi khi lấy mã nhân viên tiếp theo:", err);
        res.status(500).json({ message: 'Lỗi server khi tạo mã nhân viên.' });
    }
};

// Chức năng Thêm Nhân viên mới
exports.addEmployee = async (req, res) => {
    const {
        EMPID, EMPNAM_VN, SEX,
        BYY, BMM, BDD, 
        INDAT, OUTDAT,
        PERSONID, PIssueDate,
        PASSPORTNO,
        HOMEADDR, PHONE, MOBILEPHONE, EMAIL,
        MARRYED, SCHOOL, COUNTRY,
        BANKID, taxCode, VISANO,
        JOB, GROUPID, WHSNO, ZUNO,
        // Các trường mới từ EMPFILEB
        WKD_No, WKD_dueDate, experience, urgent_person, releation,
        urgent_addr, urgent_tel, urgent_mobile, bh_person, bh_personID, b_shift, soBH,
        // Các trường bổ sung từ EMPFILE
        empType, TX, BHDAT, GTDAT, MEMO, AGES, PDUEDATE, VDUEDATE, StudyJob, Grps
    } = req.body;

    if (!EMPID || !EMPNAM_VN || !INDAT) {
        return res.status(400).json({ message: 'Mã nhân viên, Tên và Ngày vào làm là bắt buộc.' });
    }

    const transaction = new sql.Transaction(pool);
    try {
        await transaction.begin();
        
        const checkRequest = new sql.Request(transaction);
        const checkExist = await checkRequest.input('EMPID', sql.VarChar, EMPID).query('SELECT EMPID FROM EMPFILE WHERE EMPID = @EMPID');

        if (checkExist.recordset.length > 0) {
            await transaction.rollback();
            return res.status(409).json({ message: `Mã nhân viên ${EMPID} đã tồn tại.` });
        }

        // Parse ngày tháng từ DD/MM/YYYY sang Date object
        const parsedINDAT = parseDateString(INDAT);
        const parsedOUTDAT = parseDateString(OUTDAT);
        
        // Insert vào EMPFILE
        const insertEmpFileRequest = new sql.Request(transaction);
        await insertEmpFileRequest
            .input('EMPID', sql.VarChar, EMPID)
            .input('empType', sql.VarChar, toEmptyStringIfNull(empType))
            .input('INDAT', sql.DateTime, parsedINDAT)
            .input('TX', sql.Decimal(10,1), TX || 0.0)
            .input('EMPNAM_CN', sql.NVarChar, '') // Luôn gửi chuỗi rỗng
            .input('EMPNAM_VN', sql.NVarChar, EMPNAM_VN)
            .input('BHDAT', sql.VarChar, toEmptyStringIfNull(BHDAT))
            .input('GTDAT', sql.VarChar, toEmptyStringIfNull(GTDAT))
            .input('OUTDAT', sql.DateTime, parsedOUTDAT)
            .input('PASSPORTNO', sql.NVarChar, toEmptyStringIfNull(PASSPORTNO))
            .input('COUNTRY', sql.VarChar, toEmptyStringIfNull(COUNTRY))
            .input('PHONE', sql.VarChar, toEmptyStringIfNull(PHONE))
            .input('MOBILEPHONE', sql.VarChar, toEmptyStringIfNull(MOBILEPHONE))
            .input('EMAIL', sql.VarChar, toEmptyStringIfNull(EMAIL))
            .input('HOMEADDR', sql.NVarChar, toEmptyStringIfNull(HOMEADDR))
            .input('PERSONID', sql.VarChar, toEmptyStringIfNull(PERSONID))
            .input('MEMO', sql.NVarChar, toEmptyStringIfNull(MEMO))
            .input('STATUS', sql.VarChar, toEmptyStringIfNull(null))
            .input('photos', sql.VarBinary, null)
            .input('BYY', sql.VarChar(4), toEmptyStringIfNull(BYY))
            .input('BMM', sql.VarChar(2), toEmptyStringIfNull(BMM))
            .input('BDD', sql.VarChar(2), toEmptyStringIfNull(BDD))
            .input('AGES', sql.VarChar(2), toEmptyStringIfNull(AGES))
            .input('SEX', sql.VarChar, toEmptyStringIfNull(SEX))
            .input('MARRYED', sql.VarChar, toEmptyStringIfNull(MARRYED))
            .input('SCHOOL', sql.NVarChar, toEmptyStringIfNull(SCHOOL))
            .input('PDUEDATE', sql.VarChar, toEmptyStringIfNull(PDUEDATE))
            .input('PIssueDate', sql.VarChar, toEmptyStringIfNull(PIssueDate))
            .input('VDUEDATE', sql.VarChar, toEmptyStringIfNull(VDUEDATE))
            .input('VISANO', sql.NVarChar, toEmptyStringIfNull(VISANO))
            .input('BANKID', sql.VarChar, toEmptyStringIfNull(BANKID))
            .input('StudyJob', sql.NVarChar, toEmptyStringIfNull(StudyJob))
            .input('Grps', sql.VarChar, toEmptyStringIfNull(Grps))
            .input('taxCode', sql.VarChar, toEmptyStringIfNull(taxCode))
            .input('muser', sql.VarChar, req.session.user?.id || 'ADMIN_APP')
            .query(`
                INSERT INTO EMPFILE (
                    EMPID, empType, INDAT, TX, EMPNAM_CN, EMPNAM_VN, BHDAT, GTDAT, OUTDAT,
                    PASSPORTNO, COUNTRY, PHONE, MOBILEPHONE, EMAIL, HOMEADDR, PERSONID, MEMO,
                    STATUS, photos, BYY, BMM, BDD, AGES, SEX, MARRYED, SCHOOL, PDUEDATE,
                    PIssueDate, VDUEDATE, VISANO, BANKID, StudyJob, Grps, taxCode, muser, mdtm
                ) VALUES (
                    @EMPID, @empType, @INDAT, @TX, @EMPNAM_CN, @EMPNAM_VN, @BHDAT, @GTDAT, @OUTDAT,
                    @PASSPORTNO, @COUNTRY, @PHONE, @MOBILEPHONE, @EMAIL, @HOMEADDR, @PERSONID, @MEMO,
                    @STATUS, @photos, @BYY, @BMM, @BDD, @AGES, @SEX, @MARRYED, @SCHOOL, @PDUEDATE,
                    @PIssueDate, @VDUEDATE, @VISANO, @BANKID, @StudyJob, @Grps, @taxCode, @muser, GETDATE()
                )
            `);
        
        // Parse WKD_dueDate
        const parsedWKD_dueDate = parseDateString(WKD_dueDate);

        // Insert vào EMPFILEB
        const insertEmpFileBRequest = new sql.Request(transaction);
        await insertEmpFileBRequest
            .input('EMPID', sql.VarChar, EMPID)
            .input('b_whsno', sql.VarChar, toEmptyStringIfNull(WHSNO))
            .input('b_groupid', sql.VarChar, toEmptyStringIfNull(GROUPID))
            .input('b_zuno', sql.VarChar, toEmptyStringIfNull(ZUNO))
            .input('b_job', sql.VarChar, toEmptyStringIfNull(JOB))
            .input('WKD_No', sql.VarChar, toEmptyStringIfNull(WKD_No))
            .input('WKD_dueDate', sql.VarChar, toEmptyStringIfNull(WKD_dueDate))
            .input('experience', sql.NVarChar, toEmptyStringIfNull(experience))
            .input('urgent_person', sql.VarChar, toEmptyStringIfNull(urgent_person))
            .input('releation', sql.VarChar, toEmptyStringIfNull(releation))
            .input('urgent_addr', sql.VarChar, toEmptyStringIfNull(urgent_addr))
            .input('urgent_tel', sql.VarChar, toEmptyStringIfNull(urgent_tel))
            .input('urgent_mobile', sql.VarChar, toEmptyStringIfNull(MOBILEPHONE))
            .input('bh_person', sql.VarChar, toEmptyStringIfNull(bh_person))
            .input('bh_personID', sql.VarChar, toEmptyStringIfNull(bh_personID))
            .input('b_shift', sql.VarChar, toEmptyStringIfNull(b_shift))
            .input('soBH', sql.NVarChar, toEmptyStringIfNull(soBH))
            .input('muser', sql.VarChar, req.session.user?.id || 'ADMIN_APP')
            .query(`
                INSERT INTO EMPFILEB (
                    EMPID, WKD_No, WKD_dueDate, experience, urgent_person, releation,
                    urgent_addr, urgent_tel, urgent_mobile, bh_person, bh_personID,
                    b_whsno, b_groupid, b_zuno, b_shift, soBH, b_job, muser, mdtm
                ) VALUES (
                    @EMPID, @WKD_No, @WKD_dueDate, @experience, @urgent_person, @releation,
                    @urgent_addr, @urgent_tel, @urgent_mobile, @bh_person, @bh_personID,
                    @b_whsno, @b_groupid, @b_zuno, @b_shift, @soBH, @b_job, @muser, GETDATE()
                )
            `);

        await transaction.commit();
        res.status(201).json({ message: `Đã thêm thành công nhân viên ${EMPNAM_VN}` });

    } catch (err) {
        console.error("Lỗi khi thêm nhân viên mới:", err);
        if (transaction.active) await transaction.rollback();
        res.status(500).json({ message: 'Lỗi server khi thêm nhân viên.' });
    }
};

// Chức năng Sửa thông tin nhân viên
exports.updateEmployee = async (req, res) => {
    const { empid } = req.params;
    const {
        EMPNAM_VN, SEX, BYY, BMM, BDD, INDAT, OUTDAT,
        PERSONID, PIssueDate,
        PASSPORTNO,
        HOMEADDR, PHONE, MOBILEPHONE, EMAIL,
        MARRYED, SCHOOL, COUNTRY, 
        BANKID, taxCode, VISANO,
        GROUPID, JOB, WHSNO, ZUNO,
        // Các trường mới từ EMPFILEB
        WKD_No, WKD_dueDate, experience, urgent_person, releation,
        urgent_addr, urgent_tel, urgent_mobile, bh_person, bh_personID, b_shift, soBH,
        // Các trường bổ sung từ EMPFILE
        empType, TX, BHDAT, GTDAT, MEMO, AGES, PDUEDATE, VDUEDATE, StudyJob, Grps
    } = req.body;

    const transaction = new sql.Transaction(pool);
    try {
        await transaction.begin();

        // Parse ngày tháng từ DD/MM/YYYY sang Date object
        const parsedINDAT = parseDateString(INDAT);
        const parsedOUTDAT = parseDateString(OUTDAT);
        
        // Cập nhật EMPFILE
        const updateEmpFileRequest = new sql.Request(transaction);
        await updateEmpFileRequest
            .input('EMPID', sql.VarChar, empid)
            .input('empType', sql.VarChar, toEmptyStringIfNull(empType))
            .input('INDAT', sql.DateTime, parsedINDAT)
            .input('TX', sql.Decimal(10,1), TX || 0.0)
            .input('EMPNAM_CN', sql.NVarChar, '') // Luôn gửi chuỗi rỗng
            .input('EMPNAM_VN', sql.NVarChar, EMPNAM_VN)
            .input('BHDAT', sql.VarChar, toEmptyStringIfNull(BHDAT))
            .input('GTDAT', sql.VarChar, toEmptyStringIfNull(GTDAT))
            .input('OUTDAT', sql.DateTime, parsedOUTDAT)
            .input('PASSPORTNO', sql.NVarChar, toEmptyStringIfNull(PASSPORTNO))
            .input('COUNTRY', sql.VarChar, toEmptyStringIfNull(COUNTRY)) 
            .input('PHONE', sql.VarChar, toEmptyStringIfNull(PHONE))
            .input('MOBILEPHONE', sql.VarChar, toEmptyStringIfNull(MOBILEPHONE))
            .input('EMAIL', sql.VarChar, toEmptyStringIfNull(EMAIL))
            .input('HOMEADDR', sql.NVarChar, toEmptyStringIfNull(HOMEADDR))
            .input('PERSONID', sql.VarChar, toEmptyStringIfNull(PERSONID))
            .input('MEMO', sql.NVarChar, toEmptyStringIfNull(MEMO))
            .input('STATUS', sql.VarChar, toEmptyStringIfNull(null))
            .input('photos', sql.VarBinary, null)
            .input('BYY', sql.VarChar(4), toEmptyStringIfNull(BYY))
            .input('BMM', sql.VarChar(2), toEmptyStringIfNull(BMM))
            .input('BDD', sql.VarChar(2), toEmptyStringIfNull(BDD))
            .input('AGES', sql.VarChar(2), toEmptyStringIfNull(AGES))
            .input('SEX', sql.VarChar, toEmptyStringIfNull(SEX))
            .input('MARRYED', sql.VarChar, toEmptyStringIfNull(MARRYED))
            .input('SCHOOL', sql.NVarChar, toEmptyStringIfNull(SCHOOL))
            .input('PDUEDATE', sql.VarChar, toEmptyStringIfNull(PDUEDATE))
            .input('PIssueDate', sql.VarChar, toEmptyStringIfNull(PIssueDate))
            .input('VDUEDATE', sql.VarChar, toEmptyStringIfNull(VDUEDATE))
            .input('VISANO', sql.NVarChar, toEmptyStringIfNull(VISANO))
            .input('BANKID', sql.VarChar, toEmptyStringIfNull(BANKID))
            .input('StudyJob', sql.NVarChar, toEmptyStringIfNull(StudyJob))
            .input('Grps', sql.VarChar, toEmptyStringIfNull(Grps))
            .input('taxCode', sql.VarChar, toEmptyStringIfNull(taxCode))
            .input('muser_S', sql.VarChar, req.session.user?.id || 'ADMIN_APP')
            .query(`
                UPDATE EMPFILE SET
                    empType = @empType, INDAT = @INDAT, TX = @TX, EMPNAM_CN = @EMPNAM_CN, EMPNAM_VN = @EMPNAM_VN, 
                    BHDAT = @BHDAT, GTDAT = @GTDAT, OUTDAT = @OUTDAT, PASSPORTNO = @PASSPORTNO, COUNTRY = @COUNTRY, 
                    PHONE = @PHONE, MOBILEPHONE = @MOBILEPHONE, EMAIL = @EMAIL, HOMEADDR = @HOMEADDR, PERSONID = @PERSONID, 
                    MEMO = @MEMO, STATUS = @STATUS, photos = @photos, BYY = @BYY, BMM = @BMM, BDD = @BDD, AGES = @AGES, 
                    SEX = @SEX, MARRYED = @MARRYED, SCHOOL = @SCHOOL, PDUEDATE = @PDUEDATE, PIssueDate = @PIssueDate, 
                    VDUEDATE = @VDUEDATE, VISANO = @VISANO, BANKID = @BANKID, StudyJob = @StudyJob, Grps = @Grps, 
                    taxCode = @taxCode, mdtm_S = GETDATE(), muser_S = @muser_S
                WHERE EMPID = @EMPID
            `);
        
        // Cập nhật EMPFILEB - kiểm tra xem có bản ghi tồn tại không, nếu không thì insert
        const checkEmpFileB = await new sql.Request(transaction)
            .input('EMPID', sql.VarChar, empid)
            .query(`SELECT EMPID FROM EMPFILEB WHERE EMPID = @EMPID`);

        const empFileBRequest = new sql.Request(transaction);
        if (checkEmpFileB.recordset.length > 0) {
            // Update EMPFILEB
            await empFileBRequest
                .input('EMPID', sql.VarChar, empid)
                .input('WKD_No', sql.VarChar, toEmptyStringIfNull(WKD_No))
                .input('WKD_dueDate', sql.VarChar, toEmptyStringIfNull(WKD_dueDate)) 
                .input('experience', sql.NVarChar, toEmptyStringIfNull(experience))
                .input('urgent_person', sql.VarChar, toEmptyStringIfNull(urgent_person))
                .input('releation', sql.VarChar, toEmptyStringIfNull(releation))
                .input('urgent_addr', sql.VarChar, toEmptyStringIfNull(urgent_addr))
                .input('urgent_tel', sql.VarChar, toEmptyStringIfNull(urgent_tel))
                .input('urgent_mobile', sql.VarChar, toEmptyStringIfNull(urgent_mobile))
                .input('bh_person', sql.VarChar, toEmptyStringIfNull(bh_person))
                .input('bh_personID', sql.VarChar, toEmptyStringIfNull(bh_personID))
                .input('b_whsno', sql.VarChar, toEmptyStringIfNull(WHSNO))
                .input('b_groupid', sql.VarChar, toEmptyStringIfNull(GROUPID))
                .input('b_zuno', sql.VarChar, toEmptyStringIfNull(ZUNO))
                .input('b_shift', sql.VarChar, toEmptyStringIfNull(b_shift))
                .input('soBH', sql.NVarChar, toEmptyStringIfNull(soBH))
                .input('b_job', sql.VarChar, toEmptyStringIfNull(JOB))
                .input('muser_S', sql.VarChar, req.session.user?.id || 'ADMIN_APP')
                .query(`
                    UPDATE EMPFILEB SET
                        WKD_No = @WKD_No, WKD_dueDate = @WKD_dueDate, experience = @experience,
                        urgent_person = @urgent_person, releation = @releation, urgent_addr = @urgent_addr,
                        urgent_tel = @urgent_tel, urgent_mobile = @urgent_mobile,
                        bh_person = @bh_person, bh_personID = @bh_personID,
                        b_whsno = @b_whsno, b_groupid = @b_groupid, b_zuno = @b_zuno, b_shift = @b_shift, soBH = @soBH, b_job = @b_job,
                        mdtm = GETDATE(), muser = @muser_S
                    WHERE EMPID = @EMPID
                `);
        } else {
            // Insert EMPFILEB nếu chưa có
            await empFileBRequest
                .input('EMPID', sql.VarChar, empid)
                .input('WKD_No', sql.VarChar, toEmptyStringIfNull(WKD_No))
                .input('WKD_dueDate', sql.VarChar, toEmptyStringIfNull(WKD_dueDate)) 
                .input('experience', sql.NVarChar, toEmptyStringIfNull(experience))
                .input('urgent_person', sql.VarChar, toEmptyStringIfNull(urgent_person))
                .input('releation', sql.VarChar, toEmptyTiStringIfNull(releation))
                .input('urgent_addr', sql.VarChar, toEmptyStringIfNull(urgent_addr))
                .input('urgent_tel', sql.VarChar, toEmptyStringIfNull(urgent_tel))
                .input('urgent_mobile', sql.VarChar, toEmptyStringIfNull(urgent_mobile))
                .input('bh_person', sql.VarChar, toEmptyStringIfNull(bh_person))
                .input('bh_personID', sql.VarChar, toEmptyStringIfNull(bh_personID))
                .input('b_whsno', sql.VarChar, toEmptyStringIfNull(WHSNO))
                .input('b_groupid', sql.VarChar, toEmptyStringIfNull(GROUPID))
                .input('b_zuno', sql.VarChar, toEmptyStringIfNull(ZUNO))
                .input('b_job', sql.VarChar, toEmptyStringIfNull(JOB))
                .input('muser', sql.VarChar, req.session.user?.id || 'ADMIN_APP')
                .query(`
                    INSERT INTO EMPFILEB (
                        EMPID, WKD_No, WKD_dueDate, experience, urgent_person, releation,
                        urgent_addr, urgent_tel, urgent_mobile, bh_person, bh_personID,
                        b_whsno, b_groupid, b_zuno, b_shift, soBH, b_job, muser, mdtm
                    ) VALUES (
                        @EMPID, @WKD_No, @WKD_dueDate, @experience, @urgent_person, @releation,
                        @urgent_addr, @urgent_tel, @urgent_mobile, @bh_person, @bh_personID,
                        @b_whsno, @b_groupid, @b_zuno, @b_shift, @soBH, @b_job, @muser, GETDATE()
                    )
                `);
        }

        await transaction.commit();
        res.status(200).json({ message: `Cập nhật thành công thông tin cho nhân viên ${empid}` });
    } catch (err) {
        console.error(`Lỗi khi cập nhật nhân viên ${empid}:`, err);
        if (transaction.active) await transaction.rollback();
        res.status(500).json({ message: 'Lỗi server khi cập nhật thông tin.' });
    }
};

// Chức năng Xóa mềm (Thôi việc)
exports.deleteEmployee = async (req, res) => {
    const { empid } = req.params;
    const { outDate } = req.body; // Lấy ngày thôi việc từ body

    if (!outDate) {
        return res.status(400).json({ message: 'Vui lòng cung cấp ngày thôi việc.' });
    }

    try {
        await poolConnect;
        await pool.request()
            .input('EMPID', sql.VarChar, empid)
            .input('OUTDAT', sql.Date, outDate) // Sử dụng ngày được cung cấp
            .query(`UPDATE EMPFILE SET STATUS = 'Q', OUTDAT = @OUTDAT WHERE EMPID = @EMPID`);
        res.status(200).json({ message: `Đã cập nhật trạng thái thôi việc cho nhân viên ${empid}` });
    } catch (err) {
        console.error(`Lỗi khi cập nhật thôi việc cho nhân viên ${empid}:`, err);
        res.status(500).json({ message: 'Lỗi server khi cập nhật trạng thái.' });
    }
};
// Chức năng Reset mật khẩu
exports.resetPassword = async (req, res) => {
    const { empid } = req.body;
    if (!empid) {
        return res.status(400).json({ message: 'Vui lòng cung cấp Mã số nhân viên.' });
    }
    try {
        const currentPasswords = jsonDataState.userPasswords;
        const updatedPasswords = currentPasswords.filter(user => user.empid !== empid);
        await writeJsonAndUpdateState(USER_PASSWORDS_FILE, updatedPasswords);
        res.status(200).json({ message: `Đã reset mật khẩu cho nhân viên ${empid} về mặc định.` });
    } catch (err) {
        console.error(`Lỗi khi reset mật khẩu cho ${empid}:`, err);
        res.status(500).json({ message: 'Lỗi server khi reset mật khẩu.' });
    }
};