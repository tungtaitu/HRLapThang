/*
================================================================================
File: src/api.js
Mô tả: Tập trung tất cả các hàm gọi API đến backend.
================================================================================
*/
const API_URL = process.env.NODE_ENV === 'production'
    ? '' // Ở production, API được gọi trên cùng domain
    : 'http://localhost:5000'; // Ở development, chỉ định rõ backend

// --- Hàm fetch tùy chỉnh để tự động gửi cookie ---
const customFetch = async (url, options = {}) => {
    options.credentials = 'include';
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
            // Do nothing, use default error message
        }
        throw new Error(errorMessage);
    }
    const contentType = response.headers.get("content-type");
    if (contentType && contentType.indexOf("application/json") !== -1) {
        return response.json();
    }
    return { success: true, message: 'Thao tác thành công' };
};

// --- Các hàm gọi API ---
export const apiLogin = (empid, password) => customFetch(`/api/login`, { method: 'POST', body: { empid, password } });
export const apiChangePassword = (userId, oldPassword, newPassword) => customFetch(`/api/user/change-password`, { method: 'POST', body: { userId, oldPassword, newPassword } });
export const apiAdminSubmitLeave = (leaveData) => customFetch(`/api/admin/submit-leave`, { method: 'POST', body: leaveData });
export const apiFetchTimesheet = (userId, yearMonth) => customFetch(`/api/timesheet/${userId}/${yearMonth}`);
export const apiFetchPayroll = (userId, yearMonth) => customFetch(`/api/payroll/${userId}/${yearMonth}`);
export const apiFetchHolidays = (userId, year) => customFetch(`/api/holidays/${userId}/${year}`);
export const apiFetchAllPayrolls = (yearMonth, groupId = 'ALL') => customFetch(`/api/admin/all-payrolls/${yearMonth}?groupId=${groupId}`);
export const apiApprovePayroll = (yearMonth) => customFetch(`/api/admin/approve-payroll`, { method: 'POST', body: { yearMonth } });
export const apiUploadLeaveFile = async (file) => {
    const formData = new FormData();
    formData.append('leaveFile', file);
    return customFetch(`/api/admin/upload-leave`, { method: 'POST', body: formData });
};
export const apiAdminResetPassword = (empid) => customFetch('/api/admin/reset-password', { method: 'POST', body: { empid } });
export const apiFetchEmployeeInfo = (empid) => customFetch(`/api/admin/employee-info/${empid}`);

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

export const apiCheckSession = () => customFetch(`/api/check-session`);
export const apiLogout = () => customFetch(`/api/logout`, { method: 'POST' });
export const apiGetGroups = () => customFetch(`/api/groups`);

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
export const apiUploadLuongT13 = (file, year) => {
    const formData = new FormData();
    formData.append('luongT13File', file);
    formData.append('year', year);
    return customFetch(`/api/admin/upload-luong-t13`, { method: 'POST', body: formData });
};
export const apiFetchLuongT13 = (userId, year) => customFetch(`/api/luong-t13/${userId}/${year}`);


/*
================================================================================
File: src/components/LoginForm.js
Mô tả: Component hiển thị form đăng nhập.
================================================================================
*/
import React, { useState } from 'react';

export default function LoginForm({ onLogin, error, isLoading }) {
    const [empid, setEmpid] = useState('');
    const [password, setPassword] = useState('');
    const [showPassword, setShowPassword] = useState(false);

    const handleSubmit = (e) => {
        e.preventDefault();
        onLogin(empid, password);
    };

    return (
        <div className="flex items-center justify-center min-h-screen bg-gray-100 px-4">
            <div className="w-full max-w-md p-8 space-y-4 bg-white rounded-lg shadow-md">
                <div className="flex flex-col items-center justify-center mb-6">
                    <img src="/logo.png" alt="Logo Công ty Lập Thắng" className="h-40 mb-4" />
                    <h1 className="text-xl font-bold text-center  text-indigo-600">CÔNG TY TNHH LẬP THẮNG</h1>
                    <h2 className="text-3xl font-bold text-center text-gray-700">Hệ Thống Nhân Sự</h2>
                </div>
                <h2 className="text-2xl font-bold text-center text-gray-800">Đăng nhập</h2>

                <form className="space-y-6" onSubmit={handleSubmit}>
                    <div>
                        <label htmlFor="empid" className="text-sm font-medium text-gray-700">Tên đăng nhập</label>
                        <input id="empid" type="text" value={empid} onChange={(e) => setEmpid(e.target.value)} required className="w-full px-3 py-2 mt-1 border border-gray-300 rounded-md shadow-sm focus:ring-indigo-500 focus:border-indigo-500" placeholder="Nhập mã nhân viên hoặc tài khoản admin" />
                    </div>
                    <div>
                        <label htmlFor="password" className="text-sm font-medium text-gray-700">Mật khẩu (Ngày sinh)</label>
                        <div className="relative mt-1">
                            <input id="password" type={showPassword ? 'text' : 'password'} value={password} onChange={(e) => setPassword(e.target.value)} required className="w-full px-3 py-2 pr-10 border border-gray-300 rounded-md shadow-sm focus:ring-indigo-500 focus:border-indigo-500" placeholder="Nhập theo định dạng ddmmyyyy" />
                            <button type="button" onClick={() => setShowPassword(!showPassword)} className="absolute inset-y-0 right-0 flex items-center px-3 text-gray-400 hover:text-gray-600" aria-label={showPassword ? "Ẩn mật khẩu" : "Hiện mật khẩu"}>
                                {showPassword ? (
                                    <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                                      <path strokeLinecap="round" strokeLinejoin="round" d="M13.875 18.825A10.05 10.05 0 0112 19c-4.478 0-8.268-2.943-9.543-7a9.97 9.97 0 011.563-3.029m5.858.908a3 3 0 114.243 4.243M9.878 9.878l4.242 4.242" />
                                    </svg>
                                ) : (
                                    <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                                      <path strokeLinecap="round" strokeLinejoin="round" d="M15 12a3 3 0 11-6 0 3 3 0 016 0z" />
                                      <path strokeLinecap="round" strokeLinejoin="round" d="M2.458 12C3.732 7.943 7.523 5 12 5c4.478 0 8.268 2.943 9.542 7-1.274 4.057-5.27 7-9.542 7S3.732 16.057 2.458 12z" />
                                    </svg>
                                )}
                            </button>
                        </div>
                    </div>

                    {error && <p className="text-sm text-center text-red-600">{error}</p>}
                    <div>
                        <button type="submit" disabled={isLoading} className="w-full px-4 py-2 font-medium text-white bg-indigo-600 rounded-md hover:bg-indigo-700 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-indigo-500 disabled:bg-gray-400">
                            {isLoading ? 'Đang xử lý...' : 'Đăng nhập'}
                        </button>
                    </div>
                </form>
            </div>
        </div>
    );
}

/*
================================================================================
File: src/components/ChangePasswordModal.js
Mô tả: Component modal để đổi mật khẩu.
================================================================================
*/
import React, { useState } from 'react';
import { apiChangePassword } from '../api'; // Giả sử api.js ở thư mục cha

export default function ChangePasswordModal({ user, onClose }) {
    const [oldPassword, setOldPassword] = useState('');
    const [newPassword, setNewPassword] = useState('');
    const [confirmPassword, setConfirmPassword] = useState('');
    const [error, setError] = useState('');
    const [isLoading, setIsLoading] = useState(false);

    const handleSubmit = async (e) => {
        e.preventDefault();
        setError('');
        if (newPassword !== confirmPassword) {
            setError('Mật khẩu mới không khớp.');
            return;
        }
        setIsLoading(true);
        try {
            await apiChangePassword(user.id, oldPassword, newPassword);
            alert('Đổi mật khẩu thành công!');
            onClose();
        } catch (err) {
            setError(err.message);
        } finally {
            setIsLoading(false);
        }
    };

    return (
        <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50 p-4">
            <div className="bg-white rounded-lg shadow-xl p-6 w-full max-w-md">
                <h2 className="text-xl font-bold mb-4">Đổi Mật Khẩu</h2>
                <form onSubmit={handleSubmit} className="space-y-4">
                    <div>
                        <label className="block text-sm font-medium text-gray-700">Mật khẩu cũ</label>
                        <input type="password" value={oldPassword} onChange={(e) => setOldPassword(e.target.value)} required className="w-full px-3 py-2 mt-1 border rounded-md" />
                    </div>
                     <div>
                        <label className="block text-sm font-medium text-gray-700">Mật khẩu mới</label>
                        <input type="password" value={newPassword} onChange={(e) => setNewPassword(e.target.value)} required className="w-full px-3 py-2 mt-1 border rounded-md" />
                    </div>
                     <div>
                        <label className="block text-sm font-medium text-gray-700">Xác nhận mật khẩu mới</label>
                        <input type="password" value={confirmPassword} onChange={(e) => setConfirmPassword(e.target.value)} required className="w-full px-3 py-2 mt-1 border rounded-md" />
                    </div>
                    {error && <p className="text-sm text-red-600">{error}</p>}
                    <div className="flex justify-end gap-4 mt-6">
                        <button type="button" onClick={onClose} className="px-4 py-2 bg-gray-200 rounded-md">Hủy</button>
                        <button type="submit" disabled={isLoading} className="px-4 py-2 bg-indigo-600 text-white rounded-md disabled:bg-indigo-400">
                            {isLoading ? 'Đang lưu...' : 'Lưu thay đổi'}
                        </button>
                    </div>
                </form>
            </div>
        </div>
    );
}

/*
================================================================================
File: src/features/common/TimesheetTable.js
Mô tả: Component hiển thị bảng chấm công (dùng chung).
================================================================================
*/
import React from 'react';

export default function TimesheetTable({ data }) {
    if (!Array.isArray(data) || data.length === 0) {
        return <p className="text-center text-gray-500 mt-4">Không có dữ liệu chấm công cho tháng này.</p>;
    }

    const totals = data.reduce((acc, row) => {
        acc.hoursWorked += row.hoursWorked || 0;
        acc.leaveHours += row.leaveHours || 0;
        acc.h1 += row.h1 || 0;
        acc.h2 += row.h2 || 0;
        acc.h3 += row.h3 || 0;
        acc.b3 += row.b3 || 0;
        acc.b4 += row.b4 || 0;
        return acc;
    }, { hoursWorked: 0, leaveHours: 0, h1: 0, h2: 0, h3: 0, b3: 0, b4: 0 });

    const formatCell = (value) => !value || value === 0 ? '-' : value;
    const formatHoursCell = (value) => !value || value === 0 ? '-' : value.toFixed(1);

    const getStatusClass = (status) => {
        switch (status) {
            case 'Đi làm': return 'text-green-700 bg-green-50';
            case 'Nghỉ phép': return 'text-blue-700 bg-blue-50';
            case 'Đi làm & Nghỉ phép': return 'text-purple-700 bg-purple-50';
            default: return 'text-gray-700 bg-gray-50';
        }
    };

    return (
        <div className="overflow-x-auto mt-4">
            <table className="min-w-full bg-white border border-gray-200">
                <thead className="bg-gray-50">
                    <tr>
                        <th className="px-4 py-2 text-left text-xs text-gray-800 uppercase tracking-wider font-bold">Ngày</th>
                        <th className="px-4 py-2 text-left text-xs text-gray-800 uppercase tracking-wider font-bold">Giờ vào</th>
                        <th className="px-4 py-2 text-left text-xs text-gray-800 uppercase tracking-wider font-bold">Giờ ra</th>
                        <th className="px-4 py-2 text-left text-xs text-gray-800 uppercase tracking-wider font-bold">Số giờ</th>
                        <th className="px-4 py-2 text-left text-xs text-gray-800 uppercase tracking-wider font-bold">Tăng ca 1.5</th>
                        <th className="px-4 py-2 text-left text-xs text-gray-800 uppercase tracking-wider font-bold">Tăng ca 2.0</th>
                        <th className="px-4 py-2 text-left text-xs text-gray-800 uppercase tracking-wider font-bold">Tăng ca 3.0</th>
                        <th className="px-4 py-2 text-left text-xs text-gray-800 uppercase tracking-wider font-bold">Tăng ca đêm</th>
                        <th className="px-4 py-2 text-left text-xs text-gray-800 uppercase tracking-wider font-bold">Phụ cấp 0.5</th>
                        <th className="px-4 py-2 text-left text-xs text-gray-800 uppercase tracking-wider font-bold">Trạng thái</th>
                        <th className="px-4 py-2 text-left text-xs text-blue-800 uppercase tracking-wider font-bold">Giờ Phép</th>
                        <th className="px-4 py-2 text-left text-xs text-blue-800 uppercase tracking-wider font-bold">Loại Phép</th>
                    </tr>
                </thead>
                <tbody className="divide-y divide-gray-200">
                    {data.map((row, index) => (
                        <tr key={index} className="hover:bg-gray-50">
                            <td className="px-4 py-2 whitespace-nowrap text-sm text-gray-800 font-bold">{new Date(row.date).toLocaleDateString('vi-VN')}</td>
                            <td className="px-4 py-2 whitespace-nowrap text-sm text-gray-500">{row.checkIn || '-'}</td>
                            <td className="px-4 py-2 whitespace-nowrap text-sm text-gray-500">{row.checkOut || '-'}</td>
                            <td className="px-4 py-2 whitespace-nowrap text-sm text-gray-500">{formatHoursCell(row.hoursWorked)}</td>
                            <td className="px-4 py-2 whitespace-nowrap text-sm text-gray-500">{formatCell(row.h1)}</td>
                            <td className="px-4 py-2 whitespace-nowrap text-sm text-gray-500">{formatCell(row.h2)}</td>
                            <td className="px-4 py-2 whitespace-nowrap text-sm text-gray-500">{formatCell(row.h3)}</td>
                            <td className="px-4 py-2 whitespace-nowrap text-sm text-gray-500">{formatCell(row.b3)}</td>
                            <td className="px-4 py-2 whitespace-nowrap text-sm text-gray-500">{formatCell(row.b4)}</td>
                            <td className="px-4 py-2 whitespace-nowrap text-sm font-medium">
                                <span className={`px-2 inline-flex text-xs leading-5 font-semibold rounded-full ${getStatusClass(row.status)}`}>
                                    {row.status}
                                </span>
                            </td>
                            <td className="px-4 py-2 whitespace-nowrap text-sm text-blue-600 font-semibold">{formatCell(row.leaveHours)}</td>
                            <td className="px-4 py-2 whitespace-nowrap text-sm text-blue-600">{row.leaveType || '-'}</td>
                        </tr>
                    ))}
                </tbody>
                <tfoot className="bg-gray-100 font-bold">
                    <tr>
                        <td colSpan="3" className="px-4 py-2 text-right text-gray-700">Tổng cộng</td>
                        <td className="px-4 py-2 text-gray-800">{totals.hoursWorked.toFixed(1)}</td>
                        <td className="px-4 py-2 text-gray-800">{totals.h1}</td>
                        <td className="px-4 py-2 text-gray-800">{totals.h2}</td>
                        <td className="px-4 py-2 text-gray-800">{totals.h3}</td>
                        <td className="px-4 py-2 text-gray-800">{totals.b3}</td>
                        <td className="px-4 py-2 text-gray-800">{totals.b4}</td>
                        <td className="px-4 py-2"></td>
                        <td className="px-4 py-2 text-blue-800">{totals.leaveHours.toFixed(1)}</td>
                        <td className="px-4 py-2"></td>
                    </tr>
                </tfoot>
            </table>
        </div>
    );
}

/*
================================================================================
File: src/features/common/PayrollDetails.js
Mô tả: Component hiển thị chi tiết phiếu lương (dùng chung).
================================================================================
*/
import React from 'react';

export default function PayrollDetails({ data, isAdminView = false }) {
    if (!data) return <p className="text-center text-gray-500 mt-4">Không có dữ liệu lương cho tháng này.</p>;
    if (data.approved === false && !isAdminView) {
        return <p className="text-center text-blue-600 bg-blue-50 p-4 rounded-md mt-4">{data.message}</p>;
    }
    const { employeeInfo = {}, earnings = [], deductions = [], overtimeAndBonus = [], summary = {} } = data;
    const formatCurrency = (amount) => {
        if (typeof amount !== 'number' || isNaN(amount)) return '0 ₫';
        return amount.toLocaleString('vi-VN', { style: 'currency', currency: 'VND' });
    };
    const totalEarnings = earnings.reduce((sum, item) => sum + (item.value || 0), 0);
    const totalDeductions = deductions.reduce((sum, item) => sum + (item.value || 0), 0);
    const totalOvertimeAndBonus = overtimeAndBonus.reduce((sum, item) => sum + (item.amount || 0), 0);
    const DetailRow = ({ label, value, colorClass = 'text-gray-800' }) => (
        <div className="flex justify-between items-center py-2 border-b border-gray-100">
            <p className="text-sm text-gray-600">{label}</p>
            <p className={`text-sm font-medium ${colorClass}`}>{formatCurrency(value)}</p>
        </div>
    );
    const OvertimeBonusTable = ({ data }) => {
        if (!data || data.length === 0) return null;
        return (
            <div className="mt-6">
                <h3 className="text-lg font-bold text-gray-700 mb-2">Chi tiết Tăng ca & Thưởng</h3>
                <div className="overflow-x-auto bg-gray-50 p-4 rounded-lg">
                    <table className="min-w-full">
                        <thead>
                            <tr>
                                <th className="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">Hạng mục</th>
                                <th className="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">Số giờ</th>
                                <th className="px-4 py-2 text-right text-xs font-medium text-gray-500 uppercase">Thành tiền</th>
                            </tr>
                        </thead>
                        <tbody className="divide-y divide-gray-200">
                            {data.map((item, index) => item.amount > 0 && (
                                <tr key={index}>
                                    <td className="px-4 py-2 text-sm text-gray-800">{item.label}</td>
                                    <td className="px-4 py-2 text-sm text-gray-500">{item.hours}</td>
                                    <td className="px-4 py-2 text-sm text-gray-800 text-right">{formatCurrency(item.amount)}</td>
                                </tr>
                            ))}
                        </tbody>
                        <tfoot className="border-t-2 border-gray-300">
                             <tr>
                                <td colSpan="2" className="px-4 py-2 text-right font-bold text-gray-700">Tổng cộng:</td>
                                <td className="px-4 py-2 text-right font-bold text-gray-800">{formatCurrency(totalOvertimeAndBonus)}</td>
                            </tr>
                        </tfoot>
                    </table>
                </div>
            </div>
        );
    }

    return (
        <div className="mt-6 bg-white p-6 rounded-2xl shadow-lg font-sans transition-all duration-300">
            <div className="text-center mb-6">
                <h2 className="text-2xl font-bold text-gray-800">BẢNG LƯƠNG CHI TIẾT</h2>
                <p className="text-md text-gray-500">Kỳ lương: Tháng {employeeInfo.thang}/{employeeInfo.nam}</p>
            </div>
            <div className="bg-gray-50 p-4 rounded-lg mb-6">
                <div className="grid grid-cols-3 gap-x-4 gap-y-2 text-sm">
                    <div><p className="font-semibold text-gray-500">NHÂN VIÊN</p><p className="font-bold text-gray-800">{employeeInfo.hoTen}</p></div>
                    <div><p className="font-semibold text-gray-500">MÃ SỐ</p><p className="font-bold text-gray-800">{employeeInfo.soThe}</p></div>
                     <div><p className="font-semibold text-gray-500">TIỀN CÔNG GIỜ</p><p className="font-bold text-gray-800">{formatCurrency(summary.tinhLuongMoiGio)}</p></div>
                    <div><p className="font-semibold text-gray-500">CHỨC VỤ</p><p className="font-bold text-gray-800">{employeeInfo.chucVu}</p></div>
                    <div><p className="font-semibold text-gray-500">ĐƠN VỊ</p><p className="font-bold text-gray-800">{employeeInfo.donVi}</p></div>
                </div>
            </div>
            <div className="bg-indigo-600 text-white p-6 rounded-xl text-center mb-6 shadow-indigo-200 shadow-md">
                <p className="text-lg font-semibold opacity-80">LƯƠNG THỰC LÃNH</p>
                <p className="text-4xl font-bold tracking-tight">{formatCurrency(summary.luongThucLanh)}</p>
            </div>
            <div className="grid md:grid-cols-2 gap-6">
                <div className="bg-green-50 p-4 rounded-lg">
                    <h3 className="font-bold text-green-800 mb-3">CÁC KHOẢN THU NHẬP</h3>
                    <div className="space-y-1">
                        {earnings.map((item, index) => (<DetailRow key={index} label={item.label} value={item.value} />))}
                         <div className="pt-2 mt-2 border-t-2 border-green-200">
                            <DetailRow label="TỔNG THU NHẬP (chưa gồm Tăng ca)" value={totalEarnings} colorClass="text-green-700 font-bold" />
                         </div>
                    </div>
                </div>
                <div className="bg-red-50 p-4 rounded-lg">
                    <h3 className="font-bold text-red-800 mb-3">CÁC KHOẢN KHẤU TRỪ</h3>
                    <div className="space-y-1">
                        {deductions.map((item, index) => (<DetailRow key={index} label={item.label} value={item.value} />))}
                        <div className="pt-2 mt-2 border-t-2 border-red-200">
                            <DetailRow label="TỔNG KHẤU TRỪ" value={totalDeductions} colorClass="text-red-700 font-bold" />
                        </div>
                    </div>
                </div>
            </div>
            <OvertimeBonusTable data={overtimeAndBonus} />
        </div>
    );
}

/*
================================================================================
File: src/features/common/HolidayTable.js
Mô tả: Component hiển thị bảng nghỉ phép (dùng chung).
================================================================================
*/
import React from 'react';

export default function HolidayTable({ data, summary }) {
    return (
        <div className="mt-4">
            {summary.isCurrentYear ? (
                <div className="bg-blue-50 border-l-4 border-blue-500 text-blue-800 p-4 rounded-r-lg mb-6">
                    <p className="font-bold">Phép năm còn lại tính tới tháng hiện tại</p>
                    <p className="text-3xl font-bold">{summary?.remaining || 0} Giờ </p>
                    <p className="text-sm mt-1">{summary.isForeigner ? 'Chế độ: Lao động nước ngoài (16 giờ/tháng)' : 'Chế độ: Lao động trong nước (8 giờ/tháng)'}</p>
                </div>
            ) : (
                <div className="bg-blue-50 border-l-4 border-blue-500 text-blue-800 p-4 rounded-r-lg mb-6">
                    <p className="font-bold">Việc tính toán phép năm chỉ áp dụng cho năm hiện tại.</p>
                </div>
            )}
            {data.length === 0 ? (
                 <p className="text-center text-gray-500 mt-4">Không có dữ liệu chi tiết ngày nghỉ cho năm này.</p>
            ) : (
                <div className="overflow-x-auto">
                    <table className="min-w-full bg-white border border-gray-200">
                        <thead className="bg-gray-50">
                            <tr>
                                <th className="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Ngày nghỉ</th>
                                <th className="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Số giờ nghỉ</th>
                                <th className="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Loại nghỉ phép</th>
                            </tr>
                        </thead>
                        <tbody className="divide-y divide-gray-200">
                            {data.map((row, index) => (
                                <tr key={index} className="hover:bg-gray-50">
                                    <td className="px-4 py-2 whitespace-nowrap text-sm text-gray-800">{new Date(row.date).toLocaleDateString('vi-VN')}</td>
                                    <td className="px-4 py-2 whitespace-nowrap text-sm text-gray-500">{row.hours} giờ</td>
                                    <td className="px-4 py-2 whitespace-nowrap text-sm text-gray-500">
                                        {row.reason}
                                        {row.memo && row.memo.trim().toLowerCase() === 'khang cong' && (
                                            <span className="ml-1 font-semibold italic text-indigo-700">
                                                ({row.memo}) 🌟
                                            </span>
                                        )}
                                    </td>
                                </tr>
                            ))}
                        </tbody>
                    </table>
                </div>
            )}
        </div>
    );
}

/*
================================================================================
File: src/features/common/LuongT13Details.js
Mô tả: Component hiển thị chi tiết lương tháng 13 (dùng chung).
================================================================================
*/
import React from 'react';

export default function LuongT13Details({ data, year }) {
    if (!data) return <p className="text-center text-gray-500 mt-4">Dữ liệu lương tháng 13 cho năm {year} chưa được cập nhật.</p>;
    const formatCurrency = (amount) => (typeof amount !== 'number' || isNaN(amount)) ? '0' : Math.round(amount).toLocaleString('vi-VN');
    const tongLuongTinhThuong = data.TongLuong - (data.ChuyenCan || 0);
    return (
        <div className="bg-slate-50 p-4 sm:p-6 lg:p-8 rounded-2xl max-w-4xl mx-auto font-sans">
            <header className="text-center mb-8"><h2 className="text-3xl font-bold text-gray-800">Phiếu Lương Thưởng Tháng 13</h2><p className="text-lg text-gray-500">Năm {year}</p></header>
            <div className="bg-white p-4 rounded-lg shadow-sm mb-6 flex justify-between items-center">
                <div><p className="text-lg font-bold text-indigo-700">{data.HoTen}</p><p className="text-sm text-gray-500">MSNV: {data.MSNV}</p></div>
                {data.ChucVu && <p className="text-sm text-gray-600 font-medium bg-gray-100 px-3 py-1 rounded-full">{data.ChucVu}</p>}
            </div>
            <div className="bg-gradient-to-r from-green-500 to-teal-500 text-white p-6 rounded-xl text-center mb-8 shadow-lg shadow-green-200"><p className="text-lg font-semibold uppercase tracking-wider opacity-80">Thực Lãnh</p><p className="text-5xl font-bold tracking-tight">{formatCurrency(data.ThucLanh)} <span className="text-3xl opacity-80">VNĐ</span></p></div>
            <div className="grid md:grid-cols-2 gap-6 mb-8">
                <div className="bg-white rounded-lg shadow-sm border border-gray-200 p-6"><div className="flex items-center mb-4"><div className="bg-green-100 text-green-600 p-2 rounded-full mr-4"><svg xmlns="http://www.w3.org/2000/svg" className="h-6 w-6" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}><path strokeLinecap="round" strokeLinejoin="round" d="M12 8c-1.657 0-3 .895-3 2s1.343 2 3 2 3 .895 3 2-1.343 2-3 2m0-8c1.11 0 2.08.402 2.599 1M12 8V7m0 1v.01" /></svg></div><h3 className="text-xl font-bold text-gray-700">Thu Nhập</h3></div><div className="space-y-3"><div className="flex justify-between items-center text-base"><span className="text-gray-600">Thưởng tháng 13</span><span className="font-semibold text-gray-800">{formatCurrency(data.TienThuongThang13)}</span></div><div className="flex justify-between items-center text-base"><span className="text-gray-600">Tiền phép năm</span><span className="font-semibold text-gray-800">{formatCurrency(data.TienPhepNam)}</span></div></div><div className="border-t my-4"></div><div className="flex justify-between items-center text-lg"><span className="font-bold text-gray-600">Tổng cộng</span><span className="font-bold text-green-600">{formatCurrency(data.TongCong)}</span></div></div>
                <div className="bg-white rounded-lg shadow-sm border border-gray-200 p-6"><div className="flex items-center mb-4"><div className="bg-red-100 text-red-600 p-2 rounded-full mr-4"><svg xmlns="http://www.w3.org/2000/svg" className="h-6 w-6" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}><path strokeLinecap="round" strokeLinejoin="round" d="M18 12H6" /></svg></div><h3 className="text-xl font-bold text-gray-700">Khoản Trừ</h3></div><div className="space-y-3"><div className="flex justify-between items-center text-base"><span className="text-gray-600">Trừ khác (biên bản...)</span><span className="font-semibold text-gray-800">{formatCurrency(data.TienTruKhac)}</span></div></div><div className="border-t my-4"></div><div className="flex justify-between items-center text-lg"><span className="font-bold text-gray-600">Tổng trừ</span><span className="font-bold text-red-600">{formatCurrency(data.TienTruKhac)}</span></div></div>
            </div>
            <details className="bg-white rounded-lg shadow-sm border border-gray-200 p-4 group"><summary className="font-semibold text-gray-700 cursor-pointer list-none flex justify-between items-center">Xem chi tiết & các chỉ số tham chiếu<svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5 transition-transform duration-300 group-open:rotate-180" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19 9l-7 7-7-7" /></svg></summary><div className="mt-4 pt-4 border-t grid grid-cols-2 md:grid-cols-3 gap-x-6 gap-y-3 text-sm"><div><p className="text-gray-500">Lương cơ bản</p><p className="font-semibold">{formatCurrency(data.LuongCoBan)}</p></div><div><p className="text-gray-500">Tổng lương (không chuyên cần)</p><p className="font-semibold">{formatCurrency(tongLuongTinhThuong)}</p></div><div><p className="text-gray-500">Hệ số thưởng</p><p className="font-semibold">{(data.HeSoThuong || 0).toFixed(2)}</p></div><div><p className="text-gray-500">Số ngày công</p><p className="font-semibold">{data.SoNgayCongThucTe}</p></div><div><p className="text-gray-500">Số tiếng phép còn lại</p><p className="font-semibold">{data.SoTiengPhepNamConLai}</p></div><div><p className="text-gray-500">Số ngày nghỉ không lương</p><p className="font-semibold">{data.SoNgayNghiKhongLuong}</p></div></div></details>
            <div className="mt-6 bg-gray-100 p-4 rounded-lg text-xs text-gray-600 space-y-1"><p className="font-bold text-gray-800">GHI CHÚ:</p><p><span className="font-semibold">* Tiền thưởng tháng 13</span> = Tổng lương(Không tính chuyên cần)/365* số ngày làm việc thực tế * hệ số thưởng</p><p><span className="font-semibold">* Tiền phép năm</span> = (Lương cơ bản + P/C chức vụ + P/C kỹ thuật + P/C Điện thoại + P/C Xăng xe + P/C Nhà ở+chuyên cần)/26/8*số tiếng phép năm còn lại</p><p><span className="font-semibold">* Hệ số</span> = Số ngày tính hệ số thưởng/30</p><p><span className="font-semibold">* Thực lãnh</span> = Tiền thưởng tháng 13 + Tiền phép năm - Tiền bị trừ khi lập biên bản - tiền khống công</p><p>(Ghi chú : 1 lần bị lập biên bản sẽ bị trừ tương ứng 5 ngày làm việc thực tế)</p></div>
        </div>
    );
}

/*
================================================================================
File: src/features/admin/UserManagementComponent.js
Mô tả: Component quản lý người dùng (reset mật khẩu).
================================================================================
*/
import React, { useState } from 'react';
import { apiAdminResetPassword, apiFetchEmployeeInfo } from '../../api';

export default function UserManagementComponent() {
    const [employeeId, setEmployeeId] = useState('');
    const [isLoading, setIsLoading] = useState(false);
    const [isChecking, setIsChecking] = useState(false);
    const [message, setMessage] = useState({ type: '', text: '' });
    const [employeeInfo, setEmployeeInfo] = useState(null);

    const handleCheckEmployee = async () => {
        if (!employeeId) {
            setMessage({ type: 'error', text: 'Vui lòng nhập Mã số nhân viên.' });
            return;
        }
        setIsChecking(true);
        setMessage({ type: '', text: '' });
        setEmployeeInfo(null);
        try {
            const info = await apiFetchEmployeeInfo(employeeId);
            setEmployeeInfo({ id: employeeId, ...info });
        } catch (error) {
            setMessage({ type: 'error', text: `Lỗi: ${error.message}` });
        } finally {
            setIsChecking(false);
        }
    };

    const handleResetPassword = async () => {
        if (!employeeInfo) {
            setMessage({ type: 'error', text: 'Vui lòng kiểm tra thông tin nhân viên trước khi reset.' });
            return;
        }
        if (!window.confirm(`Bạn có chắc muốn reset mật khẩu cho nhân viên ${employeeInfo.name} (${employeeInfo.id})? Mật khẩu của họ sẽ được đặt lại về ngày sinh.`)) {
            return;
        }
        setIsLoading(true);
        setMessage({ type: '', text: '' });
        try {
            const result = await apiAdminResetPassword(employeeInfo.id);
            setMessage({ type: 'success', text: result.message });
            setEmployeeId('');
            setEmployeeInfo(null);
        } catch (error) {
            setMessage({ type: 'error', text: `Lỗi: ${error.message}` });
        } finally {
            setIsLoading(false);
        }
    };

    const handleInputChange = (e) => {
        setEmployeeId(e.target.value.toUpperCase());
        if (employeeInfo) {
            setEmployeeInfo(null);
        }
        if (message.text) {
            setMessage({ type: '', text: '' });
        }
    };

    return (
        <div>
            <h2 className="text-2xl font-bold text-gray-800 mb-4">Quản lý Người dùng</h2>
            <div className="border-2 border-dashed border-gray-300 rounded-lg p-6">
                <h3 className="text-lg font-semibold text-gray-700 mb-2">Reset Mật khẩu Nhân viên</h3>
                <p className="text-gray-500 mb-4">
                    Nhập MSNV và nhấn "Kiểm tra" để xác nhận thông tin. Sau đó, nhấn "Reset Mật khẩu" để đưa mật khẩu về mặc định (ngày sinh).
                </p>
                <div className="flex flex-col sm:flex-row items-start gap-4">
                    <div className="flex-grow w-full">
                        <label htmlFor="reset-empid" className="block text-sm font-medium text-gray-700">Mã Nhân Viên (MSNV)</label>
                        <div className="mt-1 flex gap-2">
                            <input
                                id="reset-empid"
                                type="text"
                                value={employeeId}
                                onChange={handleInputChange}
                                onKeyDown={(e) => { if (e.key === 'Enter') { e.preventDefault(); handleCheckEmployee(); } }}
                                className="w-full px-3 py-2 border border-gray-300 rounded-md shadow-sm focus:ring-indigo-500 focus:border-indigo-500"
                                placeholder="Nhập MSNV cần kiểm tra"
                            />
                            <button type="button" onClick={handleCheckEmployee} disabled={isChecking || !employeeId} className="px-4 py-2 font-semibold text-white bg-blue-600 rounded-md hover:bg-blue-700 disabled:bg-gray-400">
                                {isChecking ? 'Đang...' : 'Kiểm tra'}
                            </button>
                        </div>
                    </div>
                </div>

                {employeeInfo && (
                    <div className="mt-4 bg-green-50 border-l-4 border-green-500 text-green-800 p-4 rounded-r-lg">
                        <p><span className="font-bold">Tên:</span> {employeeInfo.name}</p>
                        <p><span className="font-bold">Bộ phận:</span> {employeeInfo.department || 'Không rõ'}</p>
                        <button onClick={handleResetPassword} disabled={isLoading} className="mt-4 w-full sm:w-auto px-4 py-2 font-semibold text-white bg-orange-600 rounded-md hover:bg-orange-700 disabled:bg-gray-400">
                            {isLoading ? 'Đang xử lý...' : `Reset Mật khẩu cho ${employeeInfo.name}`}
                        </button>
                    </div>
                )}

                 {message.text && (
                    <p className={`mt-4 text-sm font-semibold text-center ${message.type === 'error' ? 'text-red-600' : 'text-green-600'}`}>
                        {message.text}
                    </p>
                )}
            </div>
        </div>
    );
}

/*
================================================================================
File: src/features/admin/AdminDashboard.js
Mô tả: Component chính cho giao diện quản trị.
================================================================================
*/
import React, { useState } from 'react';
// import LeaveManagementComponent from './LeaveManagementComponent';
// import ApprovePayrollComponent from './ApprovePayrollComponent';
// import AdminEmployeeCheck from './AdminEmployeeCheck';
// import AdminManualLeaveEntry from './AdminManualLeaveEntry';
// import ExportTimesheetComponent from './ExportTimesheetComponent';
// import LuongT13ManagementComponent from './LuongT13ManagementComponent';
// import UserManagementComponent from './UserManagementComponent';

// Do đang ở trong 1 file, ta sẽ định nghĩa các component này ở đây luôn
// Trong dự án thực tế, bạn sẽ import chúng từ các file riêng.
const LeaveManagementComponent = () => { /* ... code ... */ };
const ApprovePayrollComponent = () => { /* ... code ... */ };
const AdminEmployeeCheck = () => { /* ... code ... */ };
const AdminManualLeaveEntry = () => { /* ... code ... */ };
const ExportTimesheetComponent = () => { /* ... code ... */ };
const LuongT13ManagementComponent = () => { /* ... code ... */ };
// UserManagementComponent đã được định nghĩa ở trên

export default function AdminDashboard({ user, onLogout }) {
    const [view, setView] = useState('check-employee');

    const renderView = () => {
        switch (view) {
            case 'check-employee': return <AdminEmployeeCheck />;
            case 'manual-leave': return <AdminManualLeaveEntry />;
            case 'timesheet-export': return <ExportTimesheetComponent />;
            case 'leave-management': return <LeaveManagementComponent />;
            case 'approve': return <ApprovePayrollComponent />;
            case 'luong-t13': return <LuongT13ManagementComponent />;
            case 'user-management': return <UserManagementComponent />;
            default: return <AdminEmployeeCheck />;
        }
    };

    return (
        <div className="min-h-screen bg-gray-50">
            <header className="bg-white shadow-sm">
                 <div className="max-w-7xl mx-auto py-4 px-4 sm:px-6 lg:px-8 flex justify-between items-center">
                    <h1 className="text-xl font-semibold text-gray-900">Trang quản trị viên</h1>
                    <button onClick={onLogout} className="px-4 py-2 text-sm font-medium text-white bg-red-600 rounded-md hover:bg-red-700">Đăng xuất</button>
                </div>
            </header>
             <main className="max-w-7xl mx-auto py-6 sm:px-6 lg:px-8">
                <div className="border-b border-gray-200 mb-4">
                    <nav className="-mb-px flex space-x-8 overflow-x-auto" aria-label="Tabs">
                        <button onClick={() => setView('check-employee')} className={`${view === 'check-employee' ? 'border-indigo-500 text-indigo-600' : 'border-transparent text-gray-500 hover:text-gray-700 hover:border-gray-300'} whitespace-nowrap py-4 px-1 border-b-2 font-medium text-sm`}>
                            Kiểm tra NV
                        </button>
                        <button onClick={() => setView('manual-leave')} className={`${view === 'manual-leave' ? 'border-indigo-500 text-indigo-600' : 'border-transparent text-gray-500 hover:text-gray-700 hover:border-gray-300'} whitespace-nowrap py-4 px-1 border-b-2 font-medium text-sm`}>
                            Nhập Phép Thủ Công
                        </button>
                        <button onClick={() => setView('timesheet-export')} className={`${view === 'timesheet-export' ? 'border-indigo-500 text-indigo-600' : 'border-transparent text-gray-500 hover:text-gray-700 hover:border-gray-300'} whitespace-nowrap py-4 px-1 border-b-2 font-medium text-sm`}>
                            Xuất Báo Cáo Chấm Công
                        </button>
                        <button onClick={() => setView('leave-management')} className={`${view === 'leave-management' ? 'border-indigo-500 text-indigo-600' : 'border-transparent text-gray-500 hover:text-gray-700 hover:border-gray-300'} whitespace-nowrap py-4 px-1 border-b-2 font-medium text-sm`}>
                            Quản lý Phép năm
                        </button>
                        <button onClick={() => setView('approve')} className={`${view === 'approve' ? 'border-indigo-500 text-indigo-600' : 'border-transparent text-gray-500 hover:text-gray-700 hover:border-gray-300'} whitespace-nowrap py-4 px-1 border-b-2 font-medium text-sm`}>
                            Duyệt Phiếu lương
                        </button>
                       <button onClick={() => setView('luong-t13')} className={`${view === 'luong-t13' ? 'border-indigo-500 text-indigo-600' : 'border-transparent text-gray-500 hover:text-gray-700 hover:border-gray-300'} whitespace-nowrap py-4 px-1 border-b-2 font-medium text-sm`}>
                            QL Lương T13
                        </button>
                        <button onClick={() => setView('user-management')} className={`${view === 'user-management' ? 'border-indigo-500 text-indigo-600' : 'border-transparent text-gray-500 hover:text-gray-700 hover:border-gray-300'} whitespace-nowrap py-4 px-1 border-b-2 font-medium text-sm`}>
                            Quản lý Người dùng
                        </button>
                    </nav>
                </div>
                <div className="px-4 py-6 sm:px-0 bg-white rounded-lg shadow p-6">
                    {renderView()}
                </div>
            </main>
        </div>
    );
}

/*
================================================================================
File: src/features/employee/EmployeeDashboard.js
Mô tả: Component chính cho giao diện nhân viên.
================================================================================
*/
import React, { useState, useEffect } from 'react';
import ChangePasswordModal from '../../components/ChangePasswordModal';
import TimesheetTable from '../common/TimesheetTable';
import PayrollDetails from '../common/PayrollDetails';
import HolidayTable from '../common/HolidayTable';
import LuongT13Details from '../common/LuongT13Details';
import { apiFetchTimesheet, apiFetchPayroll, apiFetchHolidays, apiFetchLuongT13 } from '../../api';

export default function EmployeeDashboard({ user, onLogout }) {
    const [view, setView] = useState('timesheet');
    const [currentYear, setCurrentYear] = useState(new Date().getFullYear());
    const [currentMonth, setCurrentMonth] = useState(new Date().getMonth());
    const [timesheetData, setTimesheetData] = useState([]);
    const [payrollData, setPayrollData] = useState(null);
    const [holidayData, setHolidayData] = useState([]);
    const [holidaySummary, setHolidaySummary] = useState({ remaining: 0, isCurrentYear: false, isForeigner: false });
    const [luongT13Data, setLuongT13Data] = useState(null);
    const [isLoading, setIsLoading] = useState(false);
    const [showChangePassword, setShowChangePassword] = useState(false);

    useEffect(() => {
        if (!user) return;
        const fetchData = async () => {
            setIsLoading(true);
            try {
                if (view === 'timesheet' || view === 'payroll') {
                    const monthString = (currentMonth + 1).toString().padStart(2, '0');
                    const yearMonth = `${currentYear}-${monthString}`;
                    if (view === 'timesheet') {
                        setTimesheetData(await apiFetchTimesheet(user.id, yearMonth));
                    } else {
                        setPayrollData(await apiFetchPayroll(user.id, yearMonth));
                    }
                } else if (view === 'holiday') {
                    const { holidayList, summary, isForeigner } = await apiFetchHolidays(user.id, currentYear);
                    setHolidayData(holidayList);
                    setHolidaySummary({ ...summary, isForeigner });
                } else if (view === 'luongT13') {
                     setLuongT13Data(await apiFetchLuongT13(user.id, currentYear));
                }
            } catch (error) {
                console.error("Lỗi tải dữ liệu:", error);
            }
            setIsLoading(false);
        };
        fetchData();
    }, [user, view, currentYear, currentMonth]);

    const startYear = new Date().getFullYear() + 1;
    const years = Array.from({ length: 10 }, (_, i) => startYear - i);
    const months = Array.from({ length: 12 }, (_, i) => ({ value: i, name: `Tháng ${i + 1}` }));

    const renderContent = () => {
        if (isLoading) return <p className="text-center">Đang tải...</p>;
        switch(view) {
            case 'timesheet': return <TimesheetTable data={timesheetData} />;
            case 'payroll': return <PayrollDetails data={payrollData} />;
            case 'holiday': return <HolidayTable data={holidayData} summary={holidaySummary} />;
            case 'luongT13': return <LuongT13Details data={luongT13Data} year={currentYear} />;
            default: return null;
        }
    };

    return (
        <div className="min-h-screen bg-gray-50">
            {showChangePassword && <ChangePasswordModal user={user} onClose={() => setShowChangePassword(false)} />}
            <header className="bg-white shadow-sm">
                <div className="max-w-7xl mx-auto py-4 px-4 sm:px-6 lg:px-8 flex flex-col sm:flex-row sm:justify-between items-start sm:items-center">
                    <div className="mb-2 sm:mb-0 text-center sm:text-left">
                        <h1 className="text-xl font-semibold text-gray-900">Chào, {user.name}!</h1>
                        {user.workDuration && <p className="text-sm text-gray-500 mt-1">Thời gian làm việc: {user.workDuration}</p>}
                    </div>
                    <div className="flex flex-wrap gap-2">
                        <button onClick={() => setShowChangePassword(true)} className="px-4 py-2 text-sm font-medium text-indigo-600 bg-indigo-100 rounded-md hover:bg-indigo-200">
                            Đổi mật khẩu
                        </button>
                        <button onClick={onLogout} className="px-4 py-2 text-sm font-medium text-white bg-red-600 rounded-md hover:bg-red-700">Đăng xuất</button>
                    </div>
                </div>
            </header>
            <main className="max-w-7xl mx-auto py-6 px-2 sm:px-6 lg:px-8">
                <div className="bg-white rounded-lg shadow p-4 sm:p-6">
                    <div className="flex flex-col lg:flex-row lg:items-center lg:justify-between gap-4">
                         <h2 className="text-2xl font-bold text-gray-800 flex-shrink-0">
                           {view === 'timesheet' ? 'Bảng chấm công' : view === 'payroll' ? 'Bảng lương' : view === 'holiday' ? 'Ngày nghỉ phép' : 'Lương Tháng 13'}
                         </h2>
                        <div className="flex flex-wrap items-center gap-2">
                            <div className="flex items-center gap-2">
                                <select value={currentYear} onChange={(e) => setCurrentYear(parseInt(e.target.value))} className="px-3 py-2 border border-gray-300 rounded-md shadow-sm w-28">{years.map(year => <option key={year} value={year}>{year}</option>)}</select>
                                {(view === 'timesheet' || view === 'payroll') && (<select value={currentMonth} onChange={(e) => setCurrentMonth(parseInt(e.target.value))} className="px-3 py-2 border border-gray-300 rounded-md shadow-sm w-36">{months.map(month => <option key={month.value} value={month.value}>{month.name}</option>)}</select>)}
                            </div>
                            <div className="flex-shrink-0 grid grid-cols-2 sm:grid-cols-4 gap-2">
                                <button onClick={() => setView('timesheet')} className={`px-3 py-2 rounded-md text-sm font-medium ${view === 'timesheet' ? 'bg-indigo-600 text-white' : 'bg-gray-200 text-gray-700'}`}>Chấm công</button>
                                <button onClick={() => setView('payroll')} className={`px-3 py-2 rounded-md text-sm font-medium ${view === 'payroll' ? 'bg-indigo-600 text-white' : 'bg-gray-200 text-gray-700'}`}>Bảng lương</button>
                                <button onClick={() => setView('holiday')} className={`px-3 py-2 rounded-md text-sm font-medium ${view === 'holiday' ? 'bg-indigo-600 text-white' : 'bg-gray-200 text-gray-700'}`}>Nghỉ phép</button>
                                <button onClick={() => setView('luongT13')} className={`px-3 py-2 rounded-md text-sm font-medium ${view === 'luongT13' ? 'bg-indigo-600 text-white' : 'bg-gray-200 text-gray-700'}`}>Lương T13</button>
                            </div>
                        </div>
                    </div>
                    <div className="mt-6">
                        {renderContent()}
                    </div>
                </div>
            </main>
        </div>
    );
}


/*
================================================================================
File: src/App.js
Mô tả: Component gốc của ứng dụng, quản lý trạng thái đăng nhập và routing.
================================================================================
*/
import React, { useState, useEffect } from 'react';
import { apiLogin, apiCheckSession, apiLogout } from './api';
import LoginForm from './components/LoginForm';
import AdminDashboard from './features/admin/AdminDashboard';
import EmployeeDashboard from './features/employee/EmployeeDashboard';

export default function App() {
    const [user, setUser] = useState(null);
    const [loginError, setLoginError] = useState('');
    const [isLoading, setIsLoading] = useState(true);

    useEffect(() => {
        const checkUserSession = async () => {
            try {
                const sessionUser = await apiCheckSession();
                if (sessionUser) {
                    setUser(sessionUser);
                }
            } catch (error) {
                // No valid session, do nothing
            } finally {
                setIsLoading(false);
            }
        };
        checkUserSession();
    }, []);

    const handleLogin = async (empid, password) => {
        setIsLoading(true);
        setLoginError('');
        try {
            const loggedInUser = await apiLogin(empid, password);
            setUser(loggedInUser);
        } catch (error) {
            setLoginError(error.message);
        } finally {
            setIsLoading(false);
        }
    };

    const handleLogout = async () => {
        setIsLoading(true);
        try {
            await apiLogout();
        } catch (error) {
            console.error("Lỗi khi đăng xuất:", error);
        } finally {
            setUser(null);
            setIsLoading(false);
        }
    };

    if (isLoading && !user) {
        return <div className="flex justify-center items-center min-h-screen"><p>Đang tải ứng dụng...</p></div>;
    }

    if (user) {
        return user.isAdmin
            ? <AdminDashboard user={user} onLogout={handleLogout} />
            : <EmployeeDashboard user={user} onLogout={handleLogout} />;
    }

    return <LoginForm
        onLogin={handleLogin}
        error={loginError}
        isLoading={isLoading}
    />;
}
