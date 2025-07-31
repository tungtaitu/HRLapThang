/*
 * File: components/timesheet/TimesheetModule.js
 * Mô tả: Component cha quản lý tất cả các tính năng chấm công, sử dụng routing.
 * Cập nhật: Tích hợp component Nhập Công Lẻ.
 */
import React from 'react';
import { Routes, Route, Link, useLocation, Navigate } from 'react-router-dom';

// Import các component con cho từng chức năng
import UploadTimesheetComponent from './UploadTimesheetComponent';
import TimesheetQuery from './TimesheetQuery';
import ExportTimesheetComponent from './ExportTimesheet';
import TimesheetLog from './TimesheetLog'; // Import component Nhập Công Lẻ thật


export default function TimesheetModule() {
    const location = useLocation();

    // Cấu trúc các tab điều hướng đầy đủ
    const subNavItems = [
        { path: '/timesheet/upload', label: 'Upload Chấm Công' },
        { path: '/timesheet/query', label: 'Tra cứu & Sửa chữa' },
        { path: '/timesheet/log', label: 'Nhập Công Lẻ' },
        { path: '/timesheet/report', label: 'Báo cáo Chấm công' }
    ];

    return (
        <div>
            {/* Thanh điều hướng cho các tab con */}
            <div className="mb-6 border-b border-gray-200">
                <nav className="flex space-x-4" aria-label="Tabs">
                    {subNavItems.map(item => (
                        <Link 
                            key={item.path} 
                            to={item.path} 
                            className={`px-3 py-2 font-medium text-sm rounded-md ${location.pathname === item.path ? 'bg-indigo-100 text-indigo-700' : 'text-gray-500 hover:text-gray-700'}`}
                        >
                            {item.label}
                        </Link>
                    ))}
                </nav>
            </div>

            {/* Định nghĩa các trang con (nested routes) */}
            <Routes>
                <Route path="upload" element={<UploadTimesheetComponent />} />
                <Route path="query" element={<TimesheetQuery />} />
                {/* Sử dụng component Nhập Công Lẻ đã được import */}
                <Route path="log" element={<TimesheetLog />} />
                <Route path="report" element={<ExportTimesheetComponent />} />
                {/* Route mặc định cho /timesheet sẽ chuyển đến trang upload */}
                <Route index element={<Navigate to="upload" replace />} />
            </Routes>
        </div>
    );
}
