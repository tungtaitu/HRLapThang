/*
 * File: components/admin/EmployeeManagement.js
 * Mô tả: Component container gộp các chức năng quản lý nhân viên.
 */
import React, { useState } from 'react';
import AdminEmployeeCheck from './CheckEmployee';
import UserManagementComponent from './UserManagement';
import AddEmployee from './AddEmployee';
import EmployeeList from './EmployeeList';

export default function EmployeeManagement() {
    const [subView, setSubView] = useState('list');

    return (
        <div>
            <div className="mb-6 border-b border-gray-200">
                <nav className="flex space-x-4" aria-label="Tabs">
                    <button
                        onClick={() => setSubView('list')}
                        className={`px-3 py-2 font-medium text-sm rounded-md ${subView === 'list' ? 'bg-indigo-100 text-indigo-700' : 'text-gray-500 hover:text-gray-700'}`}
                    >
                        Danh sách & Sửa/Xóa
                    </button>
                    <button
                        onClick={() => setSubView('add')}
                        className={`px-3 py-2 font-medium text-sm rounded-md ${subView === 'add' ? 'bg-indigo-100 text-indigo-700' : 'text-gray-500 hover:text-gray-700'}`}
                    >
                        Thêm Nhân viên Mới
                    </button>
                    <button
                        onClick={() => setSubView('check')}
                        className={`px-3 py-2 font-medium text-sm rounded-md ${subView === 'check' ? 'bg-indigo-100 text-indigo-700' : 'text-gray-500 hover:text-gray-700'}`}
                    >
                        Kiểm tra Thông tin
                    </button>
                    <button
                        onClick={() => setSubView('manage')}
                        className={`px-3 py-2 font-medium text-sm rounded-md ${subView === 'manage' ? 'bg-indigo-100 text-indigo-700' : 'text-gray-500 hover:text-gray-700'}`}
                    >
                        Quản lý Tài khoản
                    </button>
                </nav>
            </div>
            {subView === 'list' && <EmployeeList />}
            {subView === 'add' && <AddEmployee />}
            {subView === 'check' && <AdminEmployeeCheck />}
            {subView === 'manage' && <UserManagementComponent />}
        </div>
    );
}
