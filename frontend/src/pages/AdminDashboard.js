/*
 * File: pages/AdminDashboard.js
 * Mô tả: Trang tổng quan chính dành cho quản trị viên.
 * Cập nhật: Tích hợp React Router để quản lý các trang con, bao gồm cả module chấm công mới.
 */
import React from 'react';
import { Routes, Route, Link, useLocation, Navigate } from 'react-router-dom';

// Import các component sẽ được dùng làm trang con
import EmployeeManagement from '../components/admin/EmployeeManagement';
import LeaveAdministration from '../components/admin/LeaveAdministration';
import PayrollManagement from '../components/admin/PayrollManagement';
// Import component module mới cho toàn bộ tính năng chấm công
import TimesheetModule from '../components/timesheet/TimesheetModule'; 

export default function AdminDashboard({ user, onLogout }) {
    const location = useLocation();

    // Cấu trúc các mục điều hướng
    const navItems = [
        { path: '/employee', label: 'Quản lý Nhân viên' },
        { path: '/leave', label: 'Quản lý Phép' },
        { path: '/payroll', label: 'Quản lý Lương & Thưởng' },
        { path: '/timesheet', label: 'Quản lý Chấm công' },
    ];

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
                        {navItems.map(item => (
                            <Link 
                                key={item.path}
                                to={item.path}
                                className={`${location.pathname.startsWith(item.path) ? 'border-indigo-500 text-indigo-600' : 'border-transparent text-gray-500 hover:text-gray-700 hover:border-gray-300'} whitespace-nowrap py-4 px-1 border-b-2 font-medium text-sm`}
                            >
                                {item.label}
                            </Link>
                        ))}
                    </nav>
                </div>
                <div className="px-4 py-6 sm:px-0 bg-white rounded-lg shadow p-6">
                    <Routes>
                        <Route path="/employee" element={<EmployeeManagement />} />
                        <Route path="/leave" element={<LeaveAdministration />} />
                        <Route path="/payroll" element={<PayrollManagement />} />
                        {/* Route cho module chấm công sẽ quản lý các route con của nó */}
                        <Route path="/timesheet/*" element={<TimesheetModule />} />
                        
                        {/* Route mặc định */}
                        <Route path="*" element={<Navigate to="/employee" replace />} />
                    </Routes>
                </div>
            </main>
        </div>
    );
}
