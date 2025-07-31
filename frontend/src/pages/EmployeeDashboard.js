/*
 * File: pages/EmployeeDashboard.js
 * Mô tả: Trang tổng quan dành cho nhân viên.
 */
import React, { useState, useEffect } from 'react';
import {
    apiFetchTimesheet,
    apiFetchPayroll,
    apiFetchHolidays,
    apiFetchLuongT13
} from '../api';
import ChangePasswordModal from '../components/common/ChangePasswordModal';
import TimesheetTable from '../components/timesheet/TimesheetTable';
import PayrollDetails from '../components/payroll/PayrollDetails';
import HolidayTable from '../components/leave/HolidayTable';
import LuongT13Details from '../components/payroll/LuongT13Details';

export default function EmployeeDashboard({ user, onLogout }) {
    const [view, setView] = useState('timesheet');
    const [currentYear, setCurrentYear] = useState(new Date().getFullYear());
    const [currentMonth, setCurrentMonth] = useState(new Date().getMonth());
    
    // States for data
    const [timesheetData, setTimesheetData] = useState([]);
    const [payrollData, setPayrollData] = useState(null);
    const [holidayData, setHolidayData] = useState([]);
    const [holidaySummary, setHolidaySummary] = useState({ remaining: 0, isCurrentYear: false, isForeigner: false });
    const [luongT13Data, setLuongT13Data] = useState(null);
    
    // Control states
    const [isLoading, setIsLoading] = useState(false);
    const [showChangePassword, setShowChangePassword] = useState(false);

    // Fetch data based on the current view
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
                // Optionally set an error state to show in the UI
            }
            setIsLoading(false);
        };
        fetchData();
    }, [user, view, currentYear, currentMonth]);

    // Options for year and month selectors
    const startYear = new Date().getFullYear() + 1;
    const years = Array.from({ length: 10 }, (_, i) => startYear - i);
    const months = Array.from({ length: 12 }, (_, i) => ({ value: i, name: `Tháng ${i + 1}` }));

    const renderContent = () => {
        if (isLoading) return <p className="text-center p-10">Đang tải...</p>;
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
