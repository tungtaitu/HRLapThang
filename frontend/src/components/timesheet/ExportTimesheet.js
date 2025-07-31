/*
 * File: components/timesheet/ExportTimesheet.js
 * Mô tả: Chức năng tra cứu và xuất báo cáo chấm công tổng hợp.
 * Cập nhật: Thay đổi cột hiển thị sang giờ phép và loại phép.
 */
import React, { useState, useEffect } from 'react';
import { useNavigate } from 'react-router-dom';
import { 
    apiGetGroups, 
    apiExportTimesheet, 
    apiGetMonthlyTimesheetSummary
} from '../../api';

export default function ExportTimesheetComponent() {
    const [yearMonth, setYearMonth] = useState(`${new Date().getFullYear()}-${String(new Date().getMonth() + 1).padStart(2, '0')}`);
    const [groups, setGroups] = useState([]);
    const [selectedGroupId, setSelectedGroupId] = useState('ALL');
    const [timesheetData, setTimesheetData] = useState([]);
    const [isLoading, setIsLoading] = useState(false);
    const [isExporting, setIsExporting] = useState(false);
    const [message, setMessage] = useState({ type: '', text: '' });
    const navigate = useNavigate();

    useEffect(() => {
        const fetchGroups = async () => {
            try {
                const data = await apiGetGroups();
                setGroups(data);
            } catch (error) {
                console.error("Lỗi khi tải danh sách bộ phận:", error);
                setMessage({ type: 'error', text: `Lỗi khi tải danh sách bộ phận: ${error.message}` });
            }
        };
        fetchGroups();
    }, []);

    const handleQuery = async () => {
        setIsLoading(true);
        setMessage({ type: '', text: '' });
        setTimesheetData([]);
        try {
            const yymm = yearMonth.replace('-', '');
            const data = await apiGetMonthlyTimesheetSummary(yymm, selectedGroupId);
            
            if (data.length === 0) {
                setMessage({ type: 'info', text: 'Không tìm thấy dữ liệu chấm công cho lựa chọn này.' });
            }
            setTimesheetData(data);
        } catch (error) {
            console.error("Lỗi khi tra cứu dữ liệu chấm công:", error);
            setMessage({ type: 'error', text: `Lỗi khi tra cứu: ${error.message}` });
        } finally {
            setIsLoading(false);
        }
    };

    const handleExport = async () => {
        if (timesheetData.length === 0) {
            setMessage({ type: 'error', text: 'Không có dữ liệu để xuất. Vui lòng tra cứu trước.' });
            return;
        }
        setIsExporting(true);
        setMessage({ type: 'info', text: 'Đang chuẩn bị file Excel...' });
        try {
            await apiExportTimesheet(yearMonth, selectedGroupId);
            setMessage({ type: 'success', text: 'Xuất file Excel thành công.' });
        } catch (error) {
            console.error("Lỗi khi xuất file Excel:", error);
            setMessage({ type: 'error', text: `Lỗi khi xuất file: ${error.message}` });
        } finally {
            setIsExporting(false);
        }
    };

    const handleNavigateToQuery = (employeeId) => {
        navigate('../query', { 
            state: { empid: employeeId, yymm: yearMonth } 
        });
    };

    return (
        <div>
            <h2 className="text-2xl font-bold text-gray-800 mb-4">BÁO CÁO TỔNG HỢP CHẤM CÔNG</h2>
            <div className="border-2 border-dashed border-gray-300 rounded-lg p-6 space-y-4 mb-6">
                <p className="text-gray-500">Chọn tháng và bộ phận để tra cứu dữ liệu. Nhấn vào tên nhân viên để xem và chỉnh sửa chi tiết.</p>
                <div className="flex flex-wrap justify-center items-end gap-4">
                    <div>
                        <label htmlFor="timesheet-month" className="block text-sm font-medium text-gray-700">Chọn tháng</label>
                        <input
                            type="month"
                            id="timesheet-month"
                            value={yearMonth}
                            onChange={(e) => setYearMonth(e.target.value)}
                            className="mt-1 px-3 py-2 border border-gray-300 rounded-md shadow-sm"
                        />
                    </div>
                    <div>
                        <label htmlFor="group-select" className="block text-sm font-medium text-gray-700">Chọn bộ phận</label>
                        <select
                            id="group-select"
                            value={selectedGroupId}
                            onChange={(e) => setSelectedGroupId(e.target.value)}
                            className="mt-1 block w-full pl-3 pr-10 py-2 text-base border-gray-300 rounded-md"
                        >
                            <option value="ALL">Tất cả bộ phận</option>
                            {groups.map(group => (
                                <option key={group.groupId} value={group.groupId}>
                                    {group.groupId} - {group.groupName}
                                </option>
                            ))}
                        </select>
                    </div>
                    <button onClick={handleQuery} disabled={isLoading} className="px-6 py-2 font-semibold text-white bg-blue-600 rounded-md hover:bg-blue-700 disabled:bg-gray-400">
                        {isLoading ? 'Đang tải...' : 'Tra cứu'}
                    </button>
                    <button onClick={handleExport} disabled={isExporting || timesheetData.length === 0} className="px-6 py-2 font-semibold text-white bg-teal-600 rounded-md hover:bg-teal-700 disabled:bg-gray-400">
                        {isExporting ? 'Đang xuất...' : 'Xuất Excel'}
                    </button>
                </div>
            </div>

            {message.text && (
                <div className={`my-4 p-3 rounded-md text-sm ${
                    message.type === 'error' ? 'bg-red-100 text-red-800' : 
                    message.type === 'success' ? 'bg-green-100 text-green-800' : 'bg-blue-100 text-blue-800'
                }`}>
                    {message.text}
                </div>
            )}

            {timesheetData.length > 0 && (
                 <div className="mt-6">
                    <div className="overflow-auto bg-white rounded-lg shadow" style={{ maxHeight: '70vh' }}>
                        <table className="min-w-full text-sm">
                            <thead>
                                <tr>
                                    {/* SỬA ĐỔI: Cập nhật lại tiêu đề bảng */}
                                    {['Mã NV', 'Họ Tên', 'Tổng giờ vắng', 'Tổng giờ làm', 'Tổng TC 1.5', 'Tổng TC 2.0', 'Tổng TC 3.0', 'Tổng TC Đêm', 'Tổng PC 0.5', 'Tổng PC 0.3', 'Tổng giờ phép', 'Loại phép'].map(header => (
                                        <th key={header} className="sticky top-0 z-10 p-2 text-left font-semibold text-gray-600 bg-gray-100 whitespace-nowrap border-b">
                                            {header}
                                        </th>
                                    ))}
                                </tr>
                            </thead>
                            <tbody className="divide-y divide-gray-200">
                                {timesheetData.map(row => (
                                    <tr key={row.empid} className="hover:bg-gray-50">
                                        <td className="p-2 whitespace-nowrap">{row.empid}</td>
                                        <td className="p-2 whitespace-nowrap">
                                            <button 
                                                onClick={() => handleNavigateToQuery(row.empid)} 
                                                className="text-blue-600 hover:text-blue-800 hover:underline font-semibold"
                                            >
                                                {row.empName}
                                            </button>
                                        </td>
                                        <td className="p-2 whitespace-nowrap text-center">{row.total_Kzhour}</td>
                                        <td className="p-2 whitespace-nowrap text-center">{row.total_TOTH}</td>
                                        <td className="p-2 whitespace-nowrap text-center">{row.total_H1}</td>
                                        <td className="p-2 whitespace-nowrap text-center">{row.total_H2}</td>
                                        <td className="p-2 whitespace-nowrap text-center">{row.total_H3}</td>
                                        <td className="p-2 whitespace-nowrap text-center">{row.total_B3}</td>
                                        <td className="p-2 whitespace-nowrap text-center">{row.total_B4}</td>
                                        <td className="p-2 whitespace-nowrap text-center">{row.total_B5}</td>
                                        {/* SỬA ĐỔI: Hiển thị dữ liệu phép mới */}
                                        <td className="p-2 whitespace-nowrap text-center text-blue-600 font-semibold">{row.total_leave_hours}</td>
                                        <td className="p-2 whitespace-nowrap text-center text-blue-600">{row.leave_types}</td>
                                    </tr>
                                ))}
                            </tbody>
                        </table>
                    </div>
                </div>
            )}
        </div>
    );
}
