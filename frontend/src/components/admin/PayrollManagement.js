/*
 * File: components/admin/PayrollManagement.js
 * Mô tả: Component container gộp các chức năng "Duyệt Lương" và "Quản lý Lương T13".
 */
import React, { useState } from 'react';
import ApprovePayrollComponent from './ApprovePayroll';
import LuongT13ManagementComponent from './LuongT13Management';

export default function PayrollManagement() {
    const [subView, setSubView] = useState('approve');

    return (
        <div>
            <div className="mb-6 border-b border-gray-200">
                <nav className="flex space-x-4" aria-label="Tabs">
                    <button
                        onClick={() => setSubView('approve')}
                        className={`px-3 py-2 font-medium text-sm rounded-md ${subView === 'approve' ? 'bg-indigo-100 text-indigo-700' : 'text-gray-500 hover:text-gray-700'}`}
                    >
                        Duyệt Lương Hàng Tháng
                    </button>
                    <button
                        onClick={() => setSubView('bonus')}
                        className={`px-3 py-2 font-medium text-sm rounded-md ${subView === 'bonus' ? 'bg-indigo-100 text-indigo-700' : 'text-gray-500 hover:text-gray-700'}`}
                    >
                        Quản lý Lương Tháng 13
                    </button>
                </nav>
            </div>
            {subView === 'approve' && <ApprovePayrollComponent />}
            {subView === 'bonus' && <LuongT13ManagementComponent />}
        </div>
    );
}
