/*
 * File: api/index.js
 * Mô tả: Tập trung tất cả các hàm gọi API của frontend.
 */

// Xác định URL của backend server
const API_URL = process.env.NODE_ENV === 'production'
    ? '' // Ở production, API được gọi trên cùng domain
    : 'http://localhost:5000'; // Ở development, chỉ định rõ backend

// --- Hàm fetch tùy chỉnh để tự động gửi cookie và xử lý lỗi ---
const customFetch = async (url, options = {}) => {
    options.credentials = 'include'; // Luôn gửi kèm cookie
    
    if (options.body && typeof options.body !== 'string' && !(options.body instanceof FormData)) {
        options.body = JSON.stringify(options.body);
    }
    
    if (!options.headers) {
        options.headers = {};
    }
    if ((options.method === 'POST' || options.method === 'PUT') && !(options.body instanceof FormData)) {
        options.headers['Content-Type'] = 'application/json';
    }

    const finalUrl = `${API_URL}${url}`;
    const response = await fetch(finalUrl, options);

    if (!response.ok) {
        let errorMessage = `Lỗi ${response.status}: ${response.statusText}`;
        try {
            const errorData = await response.json();
            if (errorData && errorData.message) {
                errorMessage = errorData.message;
            }
        } catch (jsonError) {
            // Bỏ qua nếu không parse được JSON, dùng lỗi mặc định
        }
        throw new Error(errorMessage);
    }
    
    const contentType = response.headers.get("content-type");
    if (contentType && contentType.indexOf("application/json") !== -1) {
        return response.json();
    }
    
    return { success: true, message: 'Thao tác thành công' };
};

// Auth APIs
export const apiLogin = (empid, password) => customFetch(`/api/login`, { method: 'POST', body: { empid, password } });
export const apiLogout = () => customFetch(`/api/logout`, { method: 'POST' });
export const apiCheckSession = () => customFetch(`/api/check-session`);
export const apiChangePassword = (userId, oldPassword, newPassword) => customFetch(`/api/user/change-password`, { method: 'POST', body: { userId, oldPassword, newPassword } });

// User APIs
export const apiFetchTimesheet = (userId, yearMonth) => customFetch(`/api/timesheet/${userId}/${yearMonth}`);
export const apiFetchPayroll = (userId, yearMonth) => customFetch(`/api/payroll/${userId}/${yearMonth}`);
export const apiFetchHolidays = (userId, year) => customFetch(`/api/holidays/${userId}/${year}`);
export const apiFetchLuongT13 = (userId, year) => customFetch(`/api/luong-t13/${userId}/${year}`);

// Admin - Employee Management APIs
export const apiGetAllEmployees = () => customFetch('/api/admin/employee');
export const apiAddEmployee = (employeeData) => customFetch('/api/admin/employee', { method: 'POST', body: employeeData });
export const apiGetEmployeeInfo = (empid) => customFetch(`/api/admin/employee/${empid}/info`);
export const apiUpdateEmployee = (empid, data) => customFetch(`/api/admin/employee/${empid}`, { method: 'PUT', body: data });
export const apiDeleteEmployee = (empid, outDate) => customFetch(`/api/admin/employee/${empid}/resign`, {
    method: 'PUT',
    body: { outDate }
});
export const apiAdminResetPassword = (empid) => customFetch('/api/admin/employee/reset-password', { method: 'POST', body: { empid } });
export const apiGetNextEmployeeId = () => customFetch('/api/admin/employee/next-id');

// Admin - Leave Management APIs
export const apiAdminSubmitLeave = (leaveData) => customFetch(`/api/admin/submit-leave`, { method: 'POST', body: leaveData });
export const apiGetLeaveEntries = (userId, year) => customFetch(`/api/admin/leave-entries/${userId}/${year}`);
export const apiDeleteLeaveEntry = (id) => customFetch(`/api/admin/leave-entry/${id}`, { method: 'DELETE' });
export const apiUpdateLeaveEntry = (id, updateData) => customFetch(`/api/admin/leave-entry/${id}`, { method: 'PUT', body: updateData });

// Admin - Other APIs
export const apiFetchAllPayrolls = (yearMonth, groupId = 'ALL') => customFetch(`/api/admin/all-payrolls/${yearMonth}?groupId=${groupId}`);
export const apiApprovePayroll = (yearMonth) => customFetch(`/api/admin/approve-payroll`, { method: 'POST', body: { yearMonth } });
export const apiGetGroups = () => customFetch(`/api/admin/groups`);
export const apiGetBasicCodeOptions = (func) => customFetch(`/api/admin/basic-code/${func}`);
export const apiUploadLeaveFile = async (file) => {
    const formData = new FormData();
    formData.append('leaveFile', file);
    return customFetch(`/api/admin/upload-leave`, { method: 'POST', body: formData });
};
export const apiUploadLuongT13 = (file, year) => {
    const formData = new FormData();
    formData.append('luongT13File', file);
    formData.append('year', year);
    return customFetch(`/api/admin/upload-luong-t13`, { method: 'POST', body: formData });
};
// Admin Export APIs
export const apiExportLeaveFile = async (year, groupId = 'ALL') => {
    try {
        const response = await fetch(`${API_URL}/api/admin/export-leave/${year}?groupId=${groupId}`, { credentials: 'include' });
        if (!response.ok) {
             const errorData = await response.json().catch(() => ({ message: 'Xuất file thất bại.' }));
             throw new Error(errorData.message);
        }
        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `TongHopPhepNam_${year}${groupId && groupId !== 'ALL' ? '_' + groupId : ''}.xlsx`;
        document.body.appendChild(a);
        a.click();
        a.remove();
        window.URL.revokeObjectURL(url);
    } catch (error) {
        console.error("Lỗi khi xuất file Excel:", error);
        alert("Lỗi: " + error.message);
    }
};

export const apiExportPayrolls = async (yearMonth, groupId) => {
    try {
        const response = await fetch(`${API_URL}/api/admin/export-payrolls/${yearMonth}?groupId=${groupId}`, {
            credentials: 'include'
        });

        if (!response.ok) {
            const errorData = await response.json().catch(() => ({ message: `Lỗi ${response.status}: Không thể xuất file.` }));
            throw new Error(errorData.message || 'Có lỗi xảy ra.');
        }

        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `BangLuong_${yearMonth}${groupId && groupId !== 'ALL' ? '_' + groupId : ''}.xlsx`;
        document.body.appendChild(a);
        a.click();
        a.remove();
        window.URL.revokeObjectURL(url);
    } catch (error) {
        console.error("Lỗi khi xuất file bảng lương:", error);
        alert("Lỗi khi xuất file: " + error.message);
    }
};

export const apiExportTimesheet = async (yearMonth, groupId) => {
    try {
        const response = await fetch(`${API_URL}/api/admin/export-timesheet/${yearMonth}?groupId=${groupId}`, {
            credentials: 'include'
        });
        if (!response.ok) {
            const errorData = await response.json().catch(() => {
                return { message: `Lỗi ${response.status}: Không thể xuất file.` };
            });
            throw new Error(errorData.message || 'Có lỗi xảy ra.');
        }

        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `BaoCaoChamCong_${yearMonth}${groupId && groupId !== 'ALL' ? '_' + groupId : ''}.xlsx`;
        document.body.appendChild(a);
        a.click();
        a.remove();
        window.URL.revokeObjectURL(url);

    } catch (error) {
        console.error(">>> ĐÃ XẢY RA LỖI KHI XUẤT FILE CHẤM CÔNG:", error);
        alert("Lỗi khi xuất file: " + error.message);
        if (error.message.includes('Chưa đăng nhập') || error.message.includes('401')) {
            window.location.reload();
        }
    }
};

// === TIMESHEET APIs ===

// Lấy dữ liệu chấm công tổng hợp theo tháng và bộ phận (có lọc)
export const apiGetMonthlyTimesheetSummary = (yymm, groupId = 'ALL') => customFetch(`/api/admin/timesheet/summary/${yymm}?groupId=${groupId}`);

// Lấy dữ liệu chấm công chi tiết của một nhân viên trong tháng
export const apiGetTimesheetForEmployee = (empid, yymm) => customFetch(`/api/admin/timesheet/${empid}/${yymm}`);

// Cập nhật một dòng chấm công
export const apiUpdateTimesheetEntry = (autoid, data) => customFetch(`/api/admin/timesheet/${autoid}`, {
    method: 'PUT',
    body: data
});

// Các API cho chức năng upload
export const apiUploadTimesheet = (data) => customFetch('/api/admin/timesheet/upload', {
    method: 'POST',
    body: data
});
export const apiGetAllEmployeesDetails = () => customFetch('/api/admin/timesheet/groupids');
