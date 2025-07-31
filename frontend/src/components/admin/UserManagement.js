/*
 * File: components/admin/UserManagement.js
 * Mô tả: Chức năng quản lý người dùng, cụ thể là reset mật khẩu.
 */
import React, { useState } from 'react';
import { apiGetEmployeeInfo, apiAdminResetPassword } from '../../api';

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
            const info = await apiGetEmployeeInfo(employeeId);
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
