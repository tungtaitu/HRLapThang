/*
 * File: services/json.service.js
 * Mô tả: Quản lý việc đọc, ghi và lưu trữ trạng thái của các file JSON.
 * Giải pháp cuối cùng: Sửa lỗi tham chiếu module bằng cách thay đổi thuộc tính của object.
 */
const fs = require('fs').promises;
const path = require('path');

const DATA_DIR = path.join(__dirname, '..', 'data');
const LEAVE_DATA_FILE = path.join(DATA_DIR, 'leave_data.json');
const PAYROLL_APPROVAL_FILE = path.join(DATA_DIR, 'payroll_approvals.json');
const USER_PASSWORDS_FILE = path.join(DATA_DIR, 'user_passwords.json');
const LUONG_T13_DATA_FILE = path.join(DATA_DIR, 'luong_thang_13.json');

let jsonDataState = {
    leaveData: [],
    payrollApprovals: [],
    userPasswords: [],
    luongT13Data: {}
};

const readFileAndParse = async (filePath, defaultValue) => {
    try {
        const fileContent = await fs.readFile(filePath, 'utf-8');
        if (!fileContent) return defaultValue;
        return JSON.parse(fileContent);
    } catch (error) {
        if (error.code !== 'ENOENT') {
            console.error(`!!! LỖI khi đọc file ${path.basename(filePath)}:`, error.message);
        }
        return defaultValue;
    }
};

const loadAllJsonData = async () => {
    try {
        await fs.mkdir(DATA_DIR, { recursive: true });
        const [leave, approvals, passwords, luongT13] = await Promise.all([
            readFileAndParse(LEAVE_DATA_FILE, []),
            readFileAndParse(PAYROLL_APPROVAL_FILE, []),
            readFileAndParse(USER_PASSWORDS_FILE, []),
            readFileAndParse(LUONG_T13_DATA_FILE, {}),
        ]);
        
        // --- SỬA LỖI QUAN TRỌNG ---
        // Thay vì gán lại toàn bộ object, chúng ta cập nhật các thuộc tính của nó.
        // Điều này đảm bảo các module khác đang giữ tham chiếu đến object này sẽ thấy được thay đổi.
        jsonDataState.leaveData = leave;
        jsonDataState.payrollApprovals = approvals;
        jsonDataState.userPasswords = passwords;
        jsonDataState.luongT13Data = luongT13;

        console.log('>>> Dữ liệu JSON đã được tải vào bộ nhớ thành công.');
    } catch (error) {
        console.error('!!! Lỗi nghiêm trọng khi tải dữ liệu JSON:', error);
    }
};

// Ghi dữ liệu vào file VÀ cập nhật trực tiếp vào bộ nhớ
const writeJsonAndUpdateState = async (filePath, data) => {
    await fs.writeFile(filePath, JSON.stringify(data, null, 2));
    
    // Cập nhật ngay lập tức vào bộ nhớ (state) để đảm bảo tính nhất quán
    if (filePath === LEAVE_DATA_FILE) jsonDataState.leaveData = data;
    else if (filePath === PAYROLL_APPROVAL_FILE) jsonDataState.payrollApprovals = data;
    else if (filePath === USER_PASSWORDS_FILE) jsonDataState.userPasswords = data;
    else if (filePath === LUONG_T13_DATA_FILE) jsonDataState.luongT13Data = data;
    console.log(`>>> [State] Đã cập nhật bộ nhớ trực tiếp cho file: ${path.basename(filePath)}`);
};

// Hàm này chỉ dành cho watcher sử dụng
const reloadJsonFile = async (filePath, options = {}) => {
    try {
        if (options.deleted) {
            if (filePath === LEAVE_DATA_FILE) jsonDataState.leaveData = [];
            else if (filePath === PAYROLL_APPROVAL_FILE) jsonDataState.payrollApprovals = [];
            else if (filePath === USER_PASSWORDS_FILE) jsonDataState.userPasswords = [];
            else if (filePath === LUONG_T13_DATA_FILE) jsonDataState.luongT13Data = {};
            return;
        }
        const jsonData = await readFileAndParse(filePath, null);
        if (jsonData !== null) {
            if (filePath === LEAVE_DATA_FILE) jsonDataState.leaveData = jsonData;
            else if (filePath === PAYROLL_APPROVAL_FILE) jsonDataState.payrollApprovals = jsonData;
            else if (filePath === USER_PASSWORDS_FILE) jsonDataState.userPasswords = jsonData;
            else if (filePath === LUONG_T13_DATA_FILE) jsonDataState.luongT13Data = jsonData;
            console.log(`>>> [Watcher] Đã đồng bộ bộ nhớ từ file: ${path.basename(filePath)}`);
        }
    } catch (error) {
        console.error(`!!! [Watcher] Không thể đồng bộ bộ nhớ cho file ${path.basename(filePath)} do có lỗi.`);
    }
};

module.exports = {
    jsonDataState,
    loadAllJsonData,
    writeJsonAndUpdateState,
    reloadJsonFile,
    DATA_DIR,
    LEAVE_DATA_FILE,
    PAYROLL_APPROVAL_FILE,
    USER_PASSWORDS_FILE,
    LUONG_T13_DATA_FILE
};
