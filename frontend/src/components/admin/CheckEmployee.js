/*
 * File: components/admin/CheckEmployee.js
 * Mô tả: Chức năng cho phép admin kiểm tra thông tin chấm công và lương của một nhân viên.
 */
import React, { useState } from 'react';
import { apiFetchTimesheet, apiFetchPayroll } from '../../api';
import TimesheetTable from '../timesheet/TimesheetTable';
import PayrollDetails from '../payroll/PayrollDetails';

export default function AdminEmployeeCheck() {
    const [employeeId, setEmployeeId] = useState('');
    const [searchedEmployee, setSearchedEmployee] = useState(null);
    const [yearMonth, setYearMonth] = useState(`${new Date().getFullYear()}-${String(new Date().getMonth() + 1).padStart(2, '0')}`);

    const [timesheetData, setTimesheetData] = useState(null);
    const [payrollData, setPayrollData] = useState(null);
    const [isLoading, setIsLoading] = useState(false);
    const [error, setError] = useState('');

    const handleSearch = async (e) => {
        e.preventDefault();
        if (!employeeId) return;

        setIsLoading(true);
        setError('');
        setTimesheetData(null);
        setPayrollData(null);
        setSearchedEmployee(null);

        try {
            const [timesheetRes, payrollRes] = await Promise.allSettled([
                apiFetchTimesheet(employeeId, yearMonth),
                apiFetchPayroll(employeeId, yearMonth)
            ]);

            const timesheet = timesheetRes.status === 'fulfilled' ? timesheetRes.value : null;
            const payroll = payrollRes.status === 'fulfilled' ? payrollRes.value : null;

            const hasTimesheetData = timesheet && Array.isArray(timesheet) && timesheet.length > 0;
            const hasPayrollData = payroll && payroll.employeeInfo;

            if (!hasTimesheetData && !hasPayrollData) {
                throw new Error("Không tìm thấy dữ liệu chấm công hoặc lương cho nhân viên này trong tháng đã chọn.");
            }

            setTimesheetData(timesheet);
            setPayrollData(payroll);

            const employeeName = payroll?.employeeInfo?.hoTen || 'Không tìm thấy tên';
            setSearchedEmployee({ id: employeeId, name: employeeName });

        } catch (err) {
            setError(err.message);
        } finally {
            setIsLoading(false);
        }
    };

    return (
        <div>
            <h2 className="text-2xl font-bold text-gray-800 mb-4">Kiểm tra Thông tin Nhân viên</h2>
            <form onSubmit={handleSearch} className="bg-gray-50 p-4 rounded-lg flex flex-col sm:flex-row items-center gap-4 mb-6">
                <div className="flex-grow">
                    <label htmlFor="employeeIdCheck" className="block text-sm font-medium text-gray-700">Mã Nhân Viên</label>
                    <input
                        id="employeeIdCheck"
                        type="text"
                        value={employeeId}
                        onChange={(e) => setEmployeeId(e.target.value.toUpperCase())}
                        className="mt-1 w-full px-3 py-2 border rounded-md"
                        placeholder="Nhập MSNV..."
                        required
                    />
                </div>
                <div>
                    <label htmlFor="checkMonth" className="block text-sm font-medium text-gray-700">Tháng/Năm</label>
                    <input
                        type="month"
                        id="checkMonth"
                        value={yearMonth}
                        onChange={(e) => setYearMonth(e.target.value)}
                        className="mt-1 px-3 py-2 border border-gray-300 rounded-md shadow-sm"
                    />
                </div>
                <div className="self-end mt-4 sm:mt-0">
                    <button type="submit" disabled={isLoading} className="w-full sm:w-auto px-6 py-2 font-semibold text-white bg-indigo-600 rounded-md hover:bg-indigo-700 disabled:bg-gray-400">
                        {isLoading ? 'Đang tìm...' : 'Kiểm tra'}
                    </button>
                </div>
            </form>

            {error && <p className="text-center text-red-600 bg-red-100 p-3 rounded-md">{error}</p>}

            {isLoading && <p className="text-center">Đang tải dữ liệu...</p>}

            {searchedEmployee && !isLoading && (
                <div className="space-y-8 mt-6">
                    <h3 className="text-xl font-bold text-center text-gray-700">
                        Kết quả cho: {searchedEmployee.name} ({searchedEmployee.id}) - Kỳ {yearMonth}
                    </h3>

                    <div>
                        <h4 className="text-lg font-semibold text-gray-800 mb-2 border-b-2 border-indigo-500 pb-1">Bảng Chấm Công</h4>
                        <TimesheetTable data={timesheetData} />
                    </div>

                    <div>
                        <h4 className="text-lg font-semibold text-gray-800 mb-2 border-b-2 border-indigo-500 pb-1">Phiếu Lương</h4>
                        <PayrollDetails data={payrollData} isAdminView={true} />
                    </div>
                </div>
            )}
        </div>
    );
}
