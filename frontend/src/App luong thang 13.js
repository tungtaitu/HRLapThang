import React, { useState, useEffect, useRef, useCallback } from 'react';
import { 
    startRegistration, 
    startAuthentication 
} from '@simplewebauthn/browser';

// --- HÀM GỌI API THỰC TẾ ---
const API_URL = 'https://api.nhansulapthang.io.vn'; // Thay đổi URL này nếu cần

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
    const response = await fetch(url, options);
    if (!response.ok) {
        let errorMessage = `Lỗi ${response.status}: ${response.statusText}`;
        try {
            const errorData = await response.json();
            if (errorData && errorData.message) {
                errorMessage = errorData.message;
            }
        } catch (jsonError) {
            console.error("Không thể phân tích JSON từ phản hồi lỗi:", jsonError);
        }
        throw new Error(errorMessage);
    }
    const contentType = response.headers.get("content-type");
    if (contentType && contentType.indexOf("application/json") !== -1) {
        return response.json();
    }
    if (contentType && contentType.indexOf("text/plain") !== -1) {
        return response.text();
    }
    return { success: true }; 
};

// --- Các hàm gọi API ---
const apiLogin = (empid, password) => customFetch(`${API_URL}/api/login`, { method: 'POST', body: { empid, password } });
const apiChangePassword = (userId, oldPassword, newPassword, isAdmin) => {
    const url = isAdmin ? `${API_URL}/api/admin/change-password` : `${API_URL}/api/user/change-password`;
    return customFetch(url, { method: 'POST', body: { userId, oldPassword, newPassword } });
};
const apiAdminSubmitLeave = (leaveData) => customFetch(`${API_URL}/api/admin/submit-leave`, { method: 'POST', body: leaveData });
const apiFetchTimesheet = (userId, yearMonth) => customFetch(`${API_URL}/api/timesheet/${userId}/${yearMonth}`);
const apiFetchPayroll = (userId, yearMonth) => customFetch(`${API_URL}/api/payroll/${userId}/${yearMonth}`);
const apiFetchHolidays = (userId, year) => customFetch(`${API_URL}/api/holidays/${userId}/${year}`);

// ====================== BẮT ĐẦU THAY ĐỔI ======================
const apiFetchAllPayrolls = (yearMonth, groupId = 'ALL') => customFetch(`${API_URL}/api/admin/all-payrolls/${yearMonth}?groupId=${groupId}`);
// ====================== KẾT THÚC THAY ĐỔI ======================

const apiApprovePayroll = (yearMonth) => customFetch(`${API_URL}/api/admin/approve-payroll`, { method: 'POST', body: { yearMonth } });
const apiUploadLeaveFile = async (file) => {
    const formData = new FormData();
    formData.append('leaveFile', file);
    return customFetch(`${API_URL}/api/admin/upload-leave`, { method: 'POST', body: formData });
};

// ====================== BẮT ĐẦU THAY ĐỔI ======================
const apiExportLeaveFile = async (year, groupId = 'ALL') => {
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

const apiExportPayrolls = async (yearMonth, groupId) => {
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
// ====================== KẾT THÚC THAY ĐỔI ======================

const apiCheckSession = () => customFetch(`${API_URL}/api/check-session`);
const apiLogout = () => customFetch(`${API_URL}/api/logout`, { method: 'POST' });
const apiGenerateRegistrationOptions = () => customFetch(`${API_URL}/api/webauthn/generate-registration-options`, { method: 'POST' });
const apiVerifyRegistration = (data) => customFetch(`${API_URL}/api/webauthn/verify-registration`, { method: 'POST', body: data });
const apiGenerateAuthenticationOptions = () => customFetch(`${API_URL}/api/webauthn/generate-authentication-options`);
const apiVerifyAuthentication = (data) => customFetch(`${API_URL}/api/webauthn/verify-authentication`, { method: 'POST', body: data });

const apiGetGroups = () => customFetch(`${API_URL}/api/groups`);

const apiExportTimesheet = async (yearMonth, groupId) => {
    console.log(`Đang yêu cầu xuất file chấm công cho tháng: ${yearMonth}, bộ phận: ${groupId}`);
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

// --- COMPONENT GIAO DIỆN ---
function LoginForm({ onLogin, onBiometricLogin, error, isLoading }) {
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
                <div className="relative my-4">
                    <div className="absolute inset-0 flex items-center"><div className="w-full border-t border-gray-300" /></div>
                    <div className="relative flex justify-center text-sm"><span className="px-2 bg-white text-gray-500">Hoặc</span></div>
                </div>
                <div>
                    <button type="button" onClick={onBiometricLogin} className="w-full flex justify-center items-center px-4 py-2 font-medium text-gray-700 bg-white border border-gray-300 rounded-md hover:bg-gray-50 disabled:opacity-50">
                        <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth="1.5" stroke="currentColor" className="size-6">
                        <path strokeLinecap="round" strokeLinejoin="round" d="M7.864 4.243A7.5 7.5 0 0 1 19.5 10.5c0 2.92-.556 5.709-1.568 8.268M5.742 6.364A7.465 7.465 0 0 0 4.5 10.5a7.464 7.464 0 0 1-1.15 3.993m1.989 3.559A11.209 11.209 0 0 0 8.25 10.5a3.75 3.75 0 1 1 7.5 0c0 .527-.021 1.049-.064 1.565M12 10.5a14.94 14.94 0 0 1-3.6 9.75m6.633-4.596a18.666 18.666 0 0 1-2.485 5.33" />
                        </svg>
                        Đăng nhập bằng vân tay / Face ID
                    </button>
                </div>
            </div>
        </div>
    );
}
function ChangePasswordModal({ user, onClose }) {
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
            await apiChangePassword(user.id, oldPassword, newPassword, user.isAdmin);
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

function TimesheetTable({ data }) {
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

function PayrollDetails({ data }) {
    if (!data) return <p className="text-center text-gray-500 mt-4">Không có dữ liệu lương cho tháng này.</p>;
    if (data.approved === false) {
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

function HolidayTable({ data, summary }) {
    return (
        <div className="mt-4">
            {summary.isCurrentYear ? (
                <div className="bg-blue-50 border-l-4 border-blue-500 text-blue-800 p-4 rounded-r-lg mb-6">
                    <p className="font-bold">Phép năm còn lại tính tới tháng hiện tại</p>
                    <p className="text-3xl font-bold">{summary?.remaining || 0} Giờ </p>
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
                                    <td className="px-4 py-2 whitespace-nowrap text-sm text-gray-500">{row.reason}</td>
                                </tr>
                            ))}
                        </tbody>
                    </table>
                </div>
            )}
        </div>
    );
}
function AdminManualLeaveEntry() {
    const [employeeId, setEmployeeId] = useState('');
    const [employeeInfo, setEmployeeInfo] = useState(null);
    const [isLoading, setIsLoading] = useState(false);
    const [message, setMessage] = useState({ type: '', text: '' });
    
    const [startDate, setStartDate] = useState('');
    const [endDate, setEndDate] = useState('');
    const [leaveType, setLeaveType] = useState('E');
    const [startTime, setStartTime] = useState('');
    const [endTime, setEndTime] = useState('');
    const [reason, setReason] = useState('');
    const [dailyBreakdown, setDailyBreakdown] = useState([]);

    const allRefs = {
        employeeId: useRef(null), startDate: useRef(null), endDate: useRef(null),
        leaveType: useRef(null), startTime: useRef(null), endTime: useRef(null),
        reason: useRef(null), submit: useRef(null)
    };

    const PUBLIC_HOLIDAYS_2025 = [
        '2025-01-01', '2025-01-28', '2025-01-29', '2025-01-30', '2025-01-31', 
        '2025-02-01', '2025-02-02', '2025-02-03', '2025-04-08', '2025-04-30', 
        '2025-05-01', '2025-05-02', '2025-09-01', '2025-09-02',
    ];

    const handleKeyDown = (e, nextRefKey) => {
        if (e.key === 'Enter') {
            e.preventDefault();
            if (nextRefKey && allRefs[nextRefKey] && allRefs[nextRefKey].current) {
                allRefs[nextRefKey].current.focus();
            }
        }
    };
    
    const calculateHours = useCallback((startStr, endStr) => {
        if (!startStr || !endStr || startStr.length < 4 || endStr.length < 4) return 0;
        const startH = parseInt(startStr.substring(0, 2)), startM = parseInt(startStr.substring(2, 4));
        const endH = parseInt(endStr.substring(0, 2)), endM = parseInt(endStr.substring(2, 4));
        if (isNaN(startH) || isNaN(startM) || isNaN(endH) || isNaN(endM)) return 0;
        
        const start = new Date(1970, 0, 1, startH, startM);
        const end = new Date(1970, 0, 1, endH, endM);
        if (end <= start) return 0;

        const morningStart = new Date(1970, 0, 1, 8, 0), morningEnd = new Date(1970, 0, 1, 12, 0);
        const afternoonStart = new Date(1970, 0, 1, 13, 0), afternoonEnd = new Date(1970, 0, 1, 17, 0);
        let totalMs = 0;
        
        const morningOverlapStart = Math.max(start, morningStart), morningOverlapEnd = Math.min(end, morningEnd);
        if (morningOverlapEnd > morningOverlapStart) totalMs += morningOverlapEnd - morningOverlapStart;
        const afternoonOverlapStart = Math.max(start, afternoonStart), afternoonOverlapEnd = Math.min(end, afternoonEnd);
        if (afternoonOverlapEnd > afternoonOverlapStart) totalMs += afternoonOverlapEnd - afternoonOverlapStart;
        
        return Math.round((totalMs / 3600000) * 10) / 10;
    }, []);

    const parseDateForFrontend = (ddmmyyyy) => {
        if (!ddmmyyyy || ddmmyyyy.length !== 8) return null;
        const day = parseInt(ddmmyyyy.substring(0, 2));
        const month = parseInt(ddmmyyyy.substring(2, 4)) - 1;
        const year = parseInt(ddmmyyyy.substring(4, 8));
        const date = new Date(year, month, day);
        if (date.getFullYear() !== year || date.getMonth() !== month || date.getDate() !== day) return null;
        return date;
    };
    
    useEffect(() => {
        const sDate = parseDateForFrontend(startDate);
        const eDate = parseDateForFrontend(endDate || startDate);
        if (!sDate || !eDate || eDate < sDate || !startTime || !endTime) {
            setDailyBreakdown([]);
            return;
        }

        let breakdown = [];
        let currentDate = new Date(sDate);
        while(currentDate <= eDate) {
            const dayOfWeek = currentDate.getDay(); 
            const dateString = `${currentDate.getFullYear()}-${String(currentDate.getMonth() + 1).padStart(2, '0')}-${String(currentDate.getDate()).padStart(2, '0')}`;
            
            let hours = 0;
            let note = '';

            if (dayOfWeek === 0) {
                note = 'Chủ Nhật';
            } else if (PUBLIC_HOLIDAYS_2025.includes(dateString)) {
                note = 'Ngày Lễ';
            } else {
                const isSameDayPeriod = sDate.getTime() === eDate.getTime();
                const isFirstDay = currentDate.getTime() === sDate.getTime();
                const isLastDay = currentDate.getTime() === eDate.getTime();

                if (isSameDayPeriod) { hours = calculateHours(startTime, endTime); }
                else if (isFirstDay) { hours = calculateHours(startTime, '1700'); }
                else if (isLastDay) { hours = calculateHours('0800', endTime); }
                else { hours = 8; }
            }
            
            breakdown.push({ date: currentDate.toLocaleDateString('vi-VN'), hours, note });
            currentDate.setDate(currentDate.getDate() + 1);
        }
        setDailyBreakdown(breakdown);
    }, [startDate, endDate, startTime, endTime, calculateHours]);
    
    useEffect(() => {
        if (employeeInfo && allRefs.startDate.current) {
            allRefs.startDate.current.focus();
        }
    }, [employeeInfo, allRefs.startDate]);

    const handleCheckEmployee = useCallback(async () => {
        if (!employeeId) return;
        setIsLoading(true); setEmployeeInfo(null); setMessage({ type: '', text: '' });
        try {
            const data = await apiFetchHolidays(employeeId, new Date().getFullYear());
            setEmployeeInfo({ 
                id: employeeId, 
                name: data.employeeName,
                remaining: data.summary.remaining 
            });
        } catch (error) {
            setMessage({ type: 'error', text: `Không tìm thấy nhân viên: ${error.message}` });
        } finally {
            setIsLoading(false);
        }
    }, [employeeId]);
    
    const resetFormFields = useCallback(() => {
        setStartDate(''); setEndDate(''); setStartTime(''); setEndTime('');
        setReason(''); setDailyBreakdown([]);
        setMessage({ type: 'success', text: 'Gửi thành công! Sẵn sàng nhập lượt tiếp theo.' });
        setTimeout(() => {
            setMessage({ type: '', text: '' });
            if (allRefs.startDate.current) allRefs.startDate.current.focus();
        }, 2000);
    }, [allRefs.startDate]);

    const handleSubmitLeave = useCallback(async (e) => {
        e.preventDefault();
        if (!employeeInfo) return;
        setIsLoading(true); setMessage({ type: '', text: '' });
        try {
            const leaveData = {
                userId: employeeInfo.id,
                startDate, endDate: endDate || startDate,
                leaveType, startTime, endTime, reason
            };
            await apiAdminSubmitLeave(leaveData);
            const updatedData = await apiFetchHolidays(employeeId, new Date().getFullYear());
            setEmployeeInfo({ ...employeeInfo, remaining: updatedData.summary.remaining, name: updatedData.employeeName });
            resetFormFields();
        } catch (error) {
             setMessage({ type: 'error', text: `Lỗi khi gửi đơn: ${error.message}` });
        } finally {
            setIsLoading(false);
        }
    }, [employeeInfo, employeeId, startDate, endDate, leaveType, startTime, endTime, reason, resetFormFields]);

    const totalHours = dailyBreakdown.reduce((acc, day) => acc + day.hours, 0);

    return (
        <div>
            <h2 className="text-2xl font-bold text-gray-800 mb-4">Nhập Phép Nhân Viên</h2>
            <div className="space-y-4 max-w-2xl mx-auto">
                <div className="bg-gray-50 p-4 rounded-lg">
                    <label htmlFor="employeeIdInput" className="block text-sm font-medium text-gray-700">Mã số nhân viên (MSNV)</label>
                    <div className="mt-1 flex gap-2">
                        <input type="text" id="employeeIdInput" ref={allRefs.employeeId} value={employeeId}
                               onChange={(e) => setEmployeeId(e.target.value.toUpperCase())}
                               onKeyDown={(e) => { if (e.key === 'Enter') { e.preventDefault(); handleCheckEmployee(); } }}
                               className="w-full px-3 py-2 border rounded-md" placeholder="Nhập MSNV rồi nhấn Enter..."/>
                        <button onClick={handleCheckEmployee} disabled={isLoading}
                                className="px-4 py-2 font-semibold text-white bg-indigo-600 rounded-md hover:bg-indigo-700 disabled:bg-gray-400">
                            {isLoading && !employeeInfo ? 'Đang...' : 'Kiểm tra'}
                        </button>
                    </div>
                </div>

                {employeeInfo && (
                    <form onSubmit={handleSubmitLeave} className="bg-white p-6 rounded-lg shadow-md space-y-4">
                        <div className="flex justify-between items-center bg-blue-50 border-l-4 border-blue-500 p-3 rounded-md">
                           <div>
                                <p>Đang nhập phép cho MSNV: <span className="font-bold">{employeeInfo.id} - {employeeInfo.name}</span></p>
                                <p>Số giờ phép năm còn lại: <span className="font-bold text-xl">{employeeInfo.remaining}</span> giờ</p>
                           </div>
                            <button type="button" onClick={() => { setEmployeeInfo(null); setEmployeeId(''); allRefs.employeeId.current.focus();}}
                                className="text-sm text-red-500 hover:text-red-700">
                                Đổi NV
                            </button>
                        </div>
                        
                        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                            <div>
                                <label htmlFor="startDate" className="block text-sm font-medium">Từ ngày (ddmmyyyy)</label>
                                <input type="text" id="startDate" ref={allRefs.startDate} value={startDate}
                                    onChange={e => setStartDate(e.target.value)}
                                    onKeyDown={(e) => handleKeyDown(e, 'endDate')}
                                    className="mt-1 w-full px-3 py-2 border rounded-md" maxLength={8}
                                    placeholder="Ví dụ: 18062025" required />
                            </div>
                            <div>
                                <label htmlFor="endDate" className="block text-sm font-medium">Đến ngày (ddmmyyyy)</label>
                                <input type="text" id="endDate" ref={allRefs.endDate} value={endDate}
                                    onChange={e => setEndDate(e.target.value)}
                                    onKeyDown={(e) => handleKeyDown(e, 'leaveType')}
                                    className="mt-1 w-full px-3 py-2 border rounded-md" maxLength={8}
                                    placeholder="Bỏ trống nếu nghỉ 1 ngày" />
                            </div>
                        </div>
                         <div>
                           <label htmlFor="leaveType" className="block text-sm font-medium">Loại phép</label>
                            <select id="leaveType" ref={allRefs.leaveType} value={leaveType}
                                    onChange={e => setLeaveType(e.target.value)}
                                    onKeyDown={(e) => handleKeyDown(e, 'startTime')}
                                    className="mt-1 w-full px-3 py-2 border rounded-md bg-white">
                                <option value="E">E: P.Năm </option>
                                <option value="A">A: P.Việc riêng</option>
                                <option value="B">B: P.Bệnh</option>
                                <option value="C">C: Nghỉ kết hôn</option>
                                <option value="D">D: P.Tang </option>
                                <option value="F">F: Nghỉ thai sản </option>
                                <option value="G">G: Nghỉ C.Tác </option>
                                <option value="H">H: Nghỉ C.Thường</option>
                                <option value="I">I: Đi đường</option>
                                <option value="J">J: Không lương (Absent)</option>
                            </select>
                        </div>
                        <div className="grid grid-cols-2 gap-4">
                             <div>
                                <label htmlFor="startTime" className="block text-sm font-medium">Giờ bắt đầu (hhmm)</label>
                                <input type="text" id="startTime" ref={allRefs.startTime} value={startTime}
                                       onChange={e => setStartTime(e.target.value)}
                                       onKeyDown={(e) => handleKeyDown(e, 'endTime')}
                                       maxLength={4} placeholder="Ví dụ: 0800"
                                       required className="mt-1 w-full px-3 py-2 border rounded-md"/>
                            </div>
                             <div>
                                <label htmlFor="endTime" className="block text-sm font-medium">Giờ kết thúc (hhmm)</label>
                                <input type="text" id="endTime" ref={allRefs.endTime} value={endTime}
                                       onChange={e => setEndTime(e.target.value)}
                                       onKeyDown={(e) => handleKeyDown(e, 'reason')}
                                       maxLength={4} placeholder="Ví dụ: 1700"
                                       required className="mt-1 w-full px-3 py-2 border rounded-md"/>
                            </div>
                        </div>
                        {dailyBreakdown.length > 0 && (
                            <div className="text-sm bg-indigo-50 p-3 rounded-md">
                                <h4 className="font-bold text-gray-700 mb-2">Chi tiết giờ nghỉ:</h4>
                                <ul className="list-disc list-inside space-y-1">
                                    {dailyBreakdown.map((day, index) => (
                                        <li key={index} className={`text-gray-800 ${day.note ? 'text-gray-500' : ''}`}>
                                            Ngày {day.date}:{' '}
                                            {day.note ? (
                                                <span className="font-semibold">{day.note} - 0 giờ</span>
                                            ) : (
                                                <span className="font-semibold">{day.hours} giờ</span>
                                            )}
                                        </li>
                                    ))}
                                </ul>
                                <div className="border-t mt-2 pt-2">
                                    <p className="font-bold text-indigo-700">Tổng cộng giờ phép tính: {totalHours} giờ</p>
                                </div>
                            </div>
                        )}

                        <div>
                            <label htmlFor="reason" className="block text-sm font-medium">Lý do (nếu có)</label>
                            <textarea id="reason" ref={allRefs.reason} value={reason}
                                      onChange={e => setReason(e.target.value)}
                                      onKeyDown={(e) => handleKeyDown(e, 'submit')}
                                      rows="2" className="mt-1 w-full px-3 py-2 border rounded-md"/>
                        </div>
                         <button type="submit" ref={allRefs.submit} disabled={isLoading || totalHours <= 0}
                                className="w-full px-4 py-3 font-semibold text-white bg-green-600 rounded-md hover:bg-green-700 disabled:bg-gray-400 focus:ring-2 focus:ring-green-500">
                            {isLoading ? 'Đang lưu...' : `Lưu và Nhập Phép (${totalHours} giờ)`}
                        </button>
                    </form>
                )}

                {message.text && (
                   <p className={`mt-4 text-center text-sm font-semibold ${message.type === 'error' ? 'text-red-600' : 'text-green-600'}`}>
                        {message.text}
                    </p>
                )}
            </div>
        </div>
    );
}
function ExportTimesheetComponent() {
    const [yearMonth, setYearMonth] = useState(`${new Date().getFullYear()}-${String(new Date().getMonth() + 1).padStart(2, '0')}`);
    const [groups, setGroups] = useState([]);
    const [selectedGroupId, setSelectedGroupId] = useState('ALL');
    const [isExporting, setIsExporting] = useState(false);

    useEffect(() => {
        const fetchGroups = async () => {
            try {
                const data = await apiGetGroups();
                setGroups(data);
            } catch (error) {
                console.error("Lỗi khi tải danh sách bộ phận:", error);
                if (error.message.includes('Chưa đăng nhập') || error.message.includes('401')) {
                    alert(error.message);
                    window.location.reload();
                }
            }
        };
        fetchGroups();
    }, []);

    const handleExport = async () => {
        setIsExporting(true);
        await apiExportTimesheet(yearMonth, selectedGroupId);
        setIsExporting(false);
    };

    return (
        <div>
            <h2 className="text-2xl font-bold text-gray-800 mb-4">XUẤT BÁO CÁO CHẤM CÔNG</h2>
            <div className="border-2 border-dashed border-gray-300 rounded-lg p-6 space-y-4">
                <p className="text-gray-500">Chọn tháng và bộ phận để xuất báo cáo tổng hợp giờ làm và tăng ca.</p>
                <div className="flex flex-col sm:flex-row justify-center items-center gap-4">
                    <div>
                        <label htmlFor="timesheet-month" className="block text-sm font-medium text-gray-700">Chọn tháng</label>
                        <input
                            type="month"
                            id="timesheet-month"
                            value={yearMonth}
                            onChange={(e) => setYearMonth(e.target.value)}
                            className="mt-1 px-3 py-2 border border-gray-300 rounded-md shadow-sm focus:ring-indigo-500 focus:border-indigo-500"
                        />
                    </div>
                    <div>
                        <label htmlFor="group-select" className="block text-sm font-medium text-gray-700">Chọn bộ phận</label>
                        <select
                            id="group-select"
                            value={selectedGroupId}
                            onChange={(e) => setSelectedGroupId(e.target.value)}
                            className="mt-1 block w-full pl-3 pr-10 py-2 text-base border-gray-300 focus:outline-none focus:ring-indigo-500 focus:border-indigo-500 sm:text-sm rounded-md"
                        >
                            <option value="ALL">Tất cả bộ phận</option>
                            {groups.map(group => (
                                <option key={group.groupId} value={group.groupId}>
                                    {group.groupId} - {group.groupName}
                                </option>
                            ))}
                        </select>
                    </div>
                </div>
                <div className="text-center mt-4">
                     <button onClick={handleExport} disabled={isExporting} className="px-6 py-2 font-semibold text-white bg-teal-600 rounded-md hover:bg-teal-700 disabled:bg-gray-400">
                        {isExporting ? 'Đang xử lý...' : 'Xuất Excel'}
                    </button>
                </div>
            </div>
        </div>
    );
}
function LuongT13ManagementComponent() {
    const [selectedFile, setSelectedFile] = useState(null);
    const [isUploading, setIsUploading] = useState(false);
    const [message, setMessage] = useState('');
    const [uploadYear, setUploadYear] = useState(new Date().getFullYear());
    const handleFileChange = (event) => { setSelectedFile(event.target.files[0]); setMessage(''); };
    const handleUpload = async () => {
        if (!selectedFile) { setMessage({ type: 'error', text: 'Vui lòng chọn một file Excel để upload.' }); return; }
        setIsUploading(true);
        setMessage({ type: '', text: '' });
        try {
            const result = await apiUploadLuongT13(selectedFile, uploadYear);
            setMessage({ type: 'success', text: result.message });
        } catch (error) {
            setMessage({ type: 'error', text: `Lỗi: ${error.message}` });
        } finally {
            setIsUploading(false);
            setSelectedFile(null);
            if (document.getElementById('luong-t13-file-input')) document.getElementById('luong-t13-file-input').value = null;
        }
    };
    const startYear = new Date().getFullYear() + 1;
    const years = Array.from({ length: 10 }, (_, i) => startYear - i);
    return (
        <div>
            <h2 className="text-2xl font-bold text-gray-800 mb-4">Upload Dữ liệu Lương Tháng 13</h2>
            <div className="border-2 border-dashed border-gray-300 rounded-lg p-6 text-center space-y-4">
                <p className="text-gray-500">Chọn năm áp dụng và file Excel chứa dữ liệu lương tháng 13. <br/><span className="font-semibold">Lưu ý: Tên cột trong file Excel phải khớp với mẫu chuẩn.</span></p>
                <div className="flex flex-col sm:flex-row justify-center items-center gap-4">
                    <div><label htmlFor="upload-year" className="sr-only">Chọn năm</label><select id="upload-year" value={uploadYear} onChange={(e) => setUploadYear(parseInt(e.target.value))} className="px-3 py-2 border border-gray-300 rounded-md shadow-sm">{years.map(year => <option key={year} value={year}>{year}</option>)}</select></div>
                    <input id="luong-t13-file-input" type="file" accept=".xlsx, .xls" onChange={handleFileChange} className="block w-full max-w-xs text-sm text-gray-500 file:mr-4 file:py-2 file:px-4 file:rounded-full file:border-0 file:text-sm file:font-semibold file:bg-indigo-50 file:text-indigo-700 hover:file:bg-indigo-100"/>
                </div>
                {selectedFile && <p className="text-sm text-gray-600">File đã chọn: {selectedFile.name}</p>}
                <button onClick={handleUpload} disabled={!selectedFile || isUploading} className="mt-4 px-6 py-2 font-semibold text-white bg-indigo-600 rounded-md hover:bg-indigo-700 disabled:bg-gray-400">{isUploading ? 'Đang xử lý...' : 'Upload và Cập nhật'}</button>
                {message.text && (<p className={`mt-4 text-sm font-semibold ${message.type === 'error' ? 'text-red-600' : 'text-green-600'}`}>{message.text}</p>)}
            </div>
        </div>
    );
}

function LuongT13Details({ data, year }) {
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
function AdminDashboard({ user, onLogout }) {
    const [view, setView] = useState('manual-leave'); 
    
    // ====================== BẮT ĐẦU THAY ĐỔI ======================
    const LeaveManagementComponent = () => {
        const [selectedFile, setSelectedFile] = useState(null);
        const [isUploading, setIsUploading] = useState(false);
        const [isExporting, setIsExporting] = useState(false);
        const [message, setMessage] = useState('');
        const [exportYear, setExportYear] = useState(new Date().getFullYear());
        const [groups, setGroups] = useState([]);
        const [selectedGroupId, setSelectedGroupId] = useState('ALL');

        useEffect(() => {
            const fetchGroups = async () => {
                try {
                    const data = await apiGetGroups();
                    setGroups(data);
                } catch (error) {
                    console.error("Lỗi khi tải danh sách bộ phận:", error);
                }
            };
            fetchGroups();
        }, []);

        const handleFileChange = (event) => { setSelectedFile(event.target.files[0]); setMessage(''); };
        const handleUpload = async () => {
            if (!selectedFile) { setMessage('Vui lòng chọn một file Excel để upload.'); return; }
            setIsUploading(true);
            setMessage('');
            try {
                const result = await apiUploadLeaveFile(selectedFile);
                setMessage(result.message);
            } catch (error) {
                setMessage(`Lỗi: ${error.message}`);
            } finally {
                setIsUploading(false);
                setSelectedFile(null);
                if(document.getElementById('file-input')) {
                    document.getElementById('file-input').value = null;
                }
            }
        };
        const handleExport = async () => {
            setIsExporting(true);
            await apiExportLeaveFile(exportYear, selectedGroupId);
            setIsExporting(false);
        };
        const startYear = new Date().getFullYear() + 1;
        const years = Array.from({ length: 10 }, (_, i) => startYear - i);
        return (
            <div>
                 <div className="mb-8">
                     <h2 className="text-2xl font-bold text-gray-800 mb-4">Upload Dữ liệu Cấu hình Phép Năm</h2>
                     <div className="border-2 border-dashed border-gray-300 rounded-lg p-6 text-center">
                        <p className="mb-4 text-gray-500">Chọn file Excel (.xlsx) chứa dữ liệu phép năm.</p>
                        <input id="file-input" type="file" accept=".xlsx, .xls" onChange={handleFileChange} className="block w-full max-w-xs mx-auto text-sm text-gray-500 file:mr-4 file:py-2 file:px-4 file:rounded-full file:border-0 file:text-sm file:font-semibold file:bg-indigo-50 file:text-indigo-700 hover:file:bg-indigo-100" />
                        {selectedFile && <p className="mt-4 text-sm text-gray-600">File đã chọn: {selectedFile.name}</p>}
                        <button onClick={handleUpload} disabled={!selectedFile || isUploading} className="mt-6 px-6 py-2 font-semibold text-white bg-indigo-600 rounded-md hover:bg-indigo-700 disabled:bg-gray-400">
                            {isUploading ? 'Đang xử lý...' : 'Upload và Cập nhật'}
                        </button>
                        {message && <p className={`mt-4 text-sm ${message.startsWith('Lỗi') ? 'text-red-600' : 'text-green-600'}`}>{message}</p>}
                     </div>
                 </div>
                 <div>
                     <h2 className="text-2xl font-bold text-gray-800 mb-4">Xuất File Tổng hợp Phép Năm</h2>
                     <div className="border-2 border-dashed border-gray-300 rounded-lg p-6 text-center">
                        <p className="mb-4 text-gray-500">Chọn năm và bộ phận để xuất báo cáo.</p>
                        <div className="flex flex-col sm:flex-row justify-center items-center gap-4">
                            <div>
                                <label htmlFor="leave-year" className="block text-sm font-medium text-gray-700">Chọn năm</label>
                                <select id="leave-year" value={exportYear} onChange={(e) => setExportYear(parseInt(e.target.value))} className="mt-1 px-3 py-2 border border-gray-300 rounded-md shadow-sm">
                                    {years.map(year => <option key={year} value={year}>{year}</option>)}
                                </select>
                            </div>
                            <div>
                                <label htmlFor="leave-group" className="block text-sm font-medium text-gray-700">Chọn bộ phận</label>
                                <select id="leave-group" value={selectedGroupId} onChange={(e) => setSelectedGroupId(e.target.value)} className="mt-1 block w-full pl-3 pr-10 py-2 text-base border-gray-300 focus:outline-none focus:ring-indigo-500 focus:border-indigo-500 sm:text-sm rounded-md">
                                    <option value="ALL">Tất cả bộ phận</option>
                                    {groups.map(group => (<option key={group.groupId} value={group.groupId}>{group.groupId} - {group.groupName}</option>))}
                                </select>
                            </div>
                            <div className="self-end">
                                <button onClick={handleExport} disabled={isExporting} className="px-6 py-2 font-semibold text-white bg-green-600 rounded-md hover:bg-green-700 disabled:bg-gray-400">
                                    {isExporting ? 'Đang xử lý...' : 'Xuất Excel'}
                                </button>
                            </div>
                        </div>
                     </div>
                 </div>
            </div>
        );
    }
    const ApprovePayrollComponent = () => {
        const [currentDate, setCurrentDate] = useState(new Date('2025-05-01'));
        const [payrolls, setPayrolls] = useState([]);
        const [isApproved, setIsApproved] = useState(false);
        const [isLoading, setIsLoading] = useState(false);
        const [groups, setGroups] = useState([]);
        const [selectedGroupId, setSelectedGroupId] = useState('ALL');
        const [isExporting, setIsExporting] = useState(false);
        
        const formatCurrency = (amount) => (amount || 0).toLocaleString('vi-VN', { style: 'currency', currency: 'VND' });
        const yearMonth = `${currentDate.getFullYear()}-${(currentDate.getMonth() + 1).toString().padStart(2, '0')}`;

        const fetchData = useCallback(async () => {
            setIsLoading(true);
            try {
                const { payrolls, isApproved } = await apiFetchAllPayrolls(yearMonth, selectedGroupId);
                setPayrolls(payrolls);
                setIsApproved(isApproved);
            } catch (error) { console.error(error); } finally {
                setIsLoading(false);
            }
        }, [yearMonth, selectedGroupId]);

        useEffect(() => {
            fetchData();
        }, [fetchData]);

        useEffect(() => {
            const fetchGroups = async () => {
                try {
                    const data = await apiGetGroups();
                    setGroups(data);
                } catch (error) {
                    console.error("Lỗi khi tải danh sách bộ phận:", error);
                }
            };
            fetchGroups();
        }, []);

        const handleApprove = async () => {
            if (!window.confirm(`Bạn có chắc muốn phê duyệt lương tháng ${yearMonth}? Hành động này không thể hoàn tác.`)) return;
            try {
                await apiApprovePayroll(yearMonth);
                setIsApproved(true);
                alert("Phê duyệt thành công!");
            } catch (error) {
                alert("Lỗi: " + error.message);
            }
        };

        const handleExportPayroll = async () => {
            setIsExporting(true);
            await apiExportPayrolls(yearMonth, selectedGroupId);
            setIsExporting(false);
        };
        
        return (
            <div>
                 <h2 className="text-2xl font-bold text-gray-800 mb-4">Phê duyệt Bảng lương</h2>
                 <div className="flex flex-col md:flex-row md:items-center md:justify-between gap-4 mb-4">
                     <div className="flex items-center gap-4">
                        <input type="month" value={yearMonth} onChange={(e) => setCurrentDate(new Date(e.target.value))} className="px-3 py-2 border border-gray-300 rounded-md shadow-sm"/>
                        <select value={selectedGroupId} onChange={e => setSelectedGroupId(e.target.value)} className="block w-full pl-3 pr-10 py-2 text-base border-gray-300 focus:outline-none focus:ring-indigo-500 focus:border-indigo-500 sm:text-sm rounded-md">
                            <option value="ALL">Tất cả bộ phận</option>
                            {groups.map(group => (<option key={group.groupId} value={group.groupId}>{group.groupId} - {group.groupName}</option>))}
                        </select>
                     </div>
                     <div className="flex items-center gap-2">
                        <span className={`px-3 py-1 text-sm font-semibold rounded-full ${isApproved ? 'bg-green-100 text-green-800' : 'bg-yellow-100 text-yellow-800'}`}>
                            {isApproved ? 'ĐÃ PHÊ DUYỆT' : 'CHƯA PHÊ DUYỆT'}
                        </span>
                         <button onClick={handleExportPayroll} disabled={isExporting} className="px-4 py-2 text-sm font-medium text-white bg-blue-600 rounded-md hover:bg-blue-700 disabled:bg-gray-400">
                             {isExporting ? 'Đang xuất...' : 'Xuất Excel'}
                         </button>
                         <button onClick={handleApprove} disabled={isApproved} className="px-4 py-2 text-sm font-medium text-white bg-green-600 rounded-md hover:bg-green-700 disabled:bg-gray-400">
                             Phê duyệt
                         </button>
                     </div>
                </div>
                {isLoading ? <p>Đang tải...</p> : (
                    <div className="overflow-x-auto">
                         <table className="min-w-full bg-white border">
                            <thead className="bg-gray-50"><tr className="bg-gray-50"><th className="px-4 py-2 text-left">Mã Bộ Phận</th><th className="px-4 py-2 text-left">Tên Bộ Phận</th><th className="px-4 py-2 text-left">Mã NV</th><th className="px-4 py-2 text-left">Tên Nhân Viên</th><th className="px-4 py-2 text-right">Lương Thực Lãnh</th></tr></thead>
                            <tbody className="divide-y divide-gray-200">
                                {payrolls.map(p => (
                                    <tr key={p.EMPID} className="hover:bg-gray-50">
                                        <td className="px-4 py-2">{p.GROUPID}</td>
                                        <td className="px-4 py-2">{p.GroupName}</td>
                                        <td className="px-4 py-2">{p.EMPID}</td>
                                        <td className="px-4 py-2">{p.EMPNAM_VN}</td>
                                        <td className="px-4 py-2 text-right font-semibold">{formatCurrency(p.REAL_TOTAL)}</td>
                                    </tr>
                                ))}
                            </tbody>
                         </table>
                    </div>
                )}
            </div>
        );
    }
    // ====================== KẾT THÚC THAY ĐỔI ======================

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
                    <nav className="-mb-px flex space-x-8" aria-label="Tabs">
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
                    </nav>
                </div>
                <div className="px-4 py-6 sm:px-0 bg-white rounded-lg shadow p-6">
                    {view === 'manual-leave' && <AdminManualLeaveEntry />}
                    {view === 'timesheet-export' && <ExportTimesheetComponent />}
                    {view === 'leave-management' && <LeaveManagementComponent />}
                    {view === 'approve' && <ApprovePayrollComponent />}
                    {view === 'luong-t13' && <LuongT13ManagementComponent />}
                </div>
            </main>
        </div>
    );
}

function EmployeeDashboard({ user, onLogout }) {
    const [view, setView] = useState('timesheet');
    const [currentYear, setCurrentYear] = useState(new Date().getFullYear());
    const [currentMonth, setCurrentMonth] = useState(new Date().getMonth());
    const [timesheetData, setTimesheetData] = useState([]);
    const [payrollData, setPayrollData] = useState(null);
    const [holidayData, setHolidayData] = useState([]);
    const [holidaySummary, setHolidaySummary] = useState({ remaining: 0, isCurrentYear: false });
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
                    const { holidayList, summary } = await apiFetchHolidays(user.id, currentYear);
                    setHolidayData(holidayList);
                    setHolidaySummary(summary);
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
                console.log('Chưa có session hợp lệ.');
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

    const handleBiometricLogin = async () => {
        setIsLoading(true);
        setLoginError('');
        try {
            const options = await apiGenerateAuthenticationOptions();
            const authResult = await startAuthentication(options);
            const loggedInUser = await apiVerifyAuthentication(authResult);
            if (loggedInUser) {
                setUser(loggedInUser);
            }
        } catch (error) {
            setLoginError(`Đăng nhập sinh trắc học thất bại: ${error.message || 'Vui lòng thử lại.'}`);
        } finally {
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
        onBiometricLogin={handleBiometricLogin}
        error={loginError} 
        isLoading={isLoading} 
    />;
}
