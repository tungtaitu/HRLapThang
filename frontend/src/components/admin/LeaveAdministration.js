/*
 * File: components/admin/LeaveAdministration.js
 * Mô tả: Component container gộp các chức năng "Quản lý Phép" và "Cấu hình Phép".
 */
import React, { useState } from 'react';
import LeaveManager from './LeaveManager';
import LeaveManagementComponent from './LeaveManagement';

export default function LeaveAdministration() {
    const [subView, setSubView] = useState('manage');

    return (
        <div>
            <div className="mb-6 border-b border-gray-200">
                <nav className="flex space-x-4" aria-label="Tabs">
                    <button
                        onClick={() => setSubView('manage')}
                        className={`px-3 py-2 font-medium text-sm rounded-md ${subView === 'manage' ? 'bg-indigo-100 text-indigo-700' : 'text-gray-500 hover:text-gray-700'}`}
                    >
                        Nhập / Sửa / Xóa Phép
                    </button>
                    <button
                        onClick={() => setSubView('config')}
                        className={`px-3 py-2 font-medium text-sm rounded-md ${subView === 'config' ? 'bg-indigo-100 text-indigo-700' : 'text-gray-500 hover:text-gray-700'}`}
                    >
                        Cấu hình Phép (Upload)
                    </button>
                </nav>
            </div>
            {subView === 'manage' && <LeaveManager />}
            {subView === 'config' && <LeaveManagementComponent />}
        </div>
    );
}
