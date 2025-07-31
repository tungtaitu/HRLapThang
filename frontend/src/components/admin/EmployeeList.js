/*
 * File: components/admin/EmployeeList.js
 * Mô tả: Component hiển thị danh sách nhân viên, cho phép sửa và xóa (thôi việc).
 * Cập nhật: Sửa lỗi không hiển thị modal nhập ngày khi thực hiện chức năng thôi việc.
 */
import React, { useState, useEffect, useMemo } from 'react';
import { apiGetAllEmployees, apiDeleteEmployee } from '../../api';
import EditEmployeeModal from './EditEmployeeModal'; 

// --- COMPONENT MODAL THÔI VIỆC ---
const ResignModal = ({ employee, onClose, onConfirm }) => {
    const [outDate, setOutDate] = useState('');

    const handleConfirm = () => {
        if (!outDate) {
            alert('Vui lòng chọn ngày thôi việc.');
            return;
        }
        onConfirm(employee.EMPID, outDate);
    };

    return (
        <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50 p-4">
            <div className="bg-white rounded-lg shadow-xl p-6 w-full max-w-sm">
                <h2 className="text-lg font-bold mb-4">Xác nhận Thôi việc</h2>
                <p className="text-gray-600 mb-4">
                    Nhập ngày thôi việc cho nhân viên <span className="font-semibold">{employee.EMPNAM_VN} ({employee.EMPID})</span>.
                </p>
                <div>
                    <label htmlFor="outDate" className="block text-sm font-medium text-gray-700">Ngày thôi việc</label>
                    <input
                        type="date"
                        id="outDate"
                        value={outDate}
                        onChange={(e) => setOutDate(e.target.value)}
                        className="mt-1 w-full px-3 py-2 border rounded-md"
                        required
                    />
                </div>
                <div className="flex justify-end gap-4 mt-6">
                    <button onClick={onClose} className="px-4 py-2 bg-gray-200 rounded-md hover:bg-gray-300">Hủy</button>
                    <button onClick={handleConfirm} className="px-4 py-2 bg-red-600 text-white rounded-md hover:bg-red-700">Xác nhận</button>
                </div>
            </div>
        </div>
    );
};


export default function EmployeeList() {
    const [employees, setEmployees] = useState([]);
    const [isLoading, setIsLoading] = useState(true);
    const [error, setError] = useState('');
    const [searchTerm, setSearchTerm] = useState('');
    
    const [editingEmployeeId, setEditingEmployeeId] = useState(null);
    const [resigningEmployee, setResigningEmployee] = useState(null); // State cho modal thôi việc

    const fetchEmployees = async () => {
        setIsLoading(true);
        try {
            const data = await apiGetAllEmployees();
            setEmployees(data);
        } catch (err)
 {
            setError(err.message);
        } finally {
            setIsLoading(false);
        }
    };

    useEffect(() => {
        fetchEmployees();
    }, []);

    const handleConfirmResign = async (empid, outDate) => {
        try {
            await apiDeleteEmployee(empid, outDate);
            alert('Cập nhật thành công!');
            setResigningEmployee(null); // Đóng modal
            await fetchEmployees(); // Tải lại danh sách
        } catch (err) {
            alert(`Lỗi: ${err.message}`);
        }
    };

    const handleSaveSuccess = () => {
        setEditingEmployeeId(null); // Đóng modal
        fetchEmployees(); // Tải lại danh sách
    };

    const filteredEmployees = useMemo(() => {
        if (!searchTerm) return employees;
        return employees.filter(emp => 
            emp.EMPID.toLowerCase().includes(searchTerm.toLowerCase()) ||
            emp.EMPNAM_VN.toLowerCase().includes(searchTerm.toLowerCase())
        );
    }, [employees, searchTerm]);

    return (
        <div>
            {editingEmployeeId && <EditEmployeeModal employeeId={editingEmployeeId} onClose={() => setEditingEmployeeId(null)} onSaveSuccess={handleSaveSuccess} />}
            {resigningEmployee && (
                <ResignModal 
                    employee={resigningEmployee}
                    onClose={() => setResigningEmployee(null)}
                    onConfirm={handleConfirmResign}
                />
            )}
            <h2 className="text-2xl font-bold text-gray-800 mb-4">Danh sách Nhân viên</h2>
            <div className="mb-4">
                <input 
                    type="text"
                    placeholder="Tìm kiếm theo Tên hoặc MSNV..."
                    value={searchTerm}
                    onChange={e => setSearchTerm(e.target.value)}
                    className="w-full max-w-md px-3 py-2 border rounded-md"
                />
            </div>
            {isLoading && <p>Đang tải danh sách...</p>}
            {error && <p className="text-red-500">{error}</p>}
            {!isLoading && !error && (
                <div className="overflow-x-auto">
                    <table className="min-w-full bg-white border">
                        <thead className="bg-gray-100">
                            <tr>
                                <th className="px-4 py-2 text-left">MSNV</th>
                                <th className="px-4 py-2 text-left">Họ Tên</th>
                                <th className="px-4 py-2 text-left">Ngày vào làm</th>
                                <th className="px-4 py-2 text-left">Hành động</th>
                            </tr>
                        </thead>
                        <tbody className="divide-y">
                            {filteredEmployees.map(emp => (
                                <tr key={emp.EMPID}>
                                    <td className="px-4 py-2 font-semibold">{emp.EMPID}</td>
                                    <td className="px-4 py-2">{emp.EMPNAM_VN}</td>
                                    <td className="px-4 py-2">{new Date(emp.INDAT).toLocaleDateString('vi-VN')}</td>
                                    <td className="px-4 py-2 space-x-2">
                                        <button onClick={() => setEditingEmployeeId(emp.EMPID)} className="px-3 py-1 text-sm text-blue-600 bg-blue-100 rounded-md hover:bg-blue-200">
                                            Sửa
                                        </button>
                                        <button onClick={() => setResigningEmployee(emp)} className="px-3 py-1 text-sm text-red-600 bg-red-100 rounded-md hover:bg-red-200">
                                            Thôi việc
                                        </button>
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
