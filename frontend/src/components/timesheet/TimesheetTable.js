/*
 * File: components/timesheet/TimesheetTable.js
 * Mô tả: Component hiển thị bảng chấm công chi tiết của nhân viên.
 * Cập nhật: Cố định dòng tiêu đề của bảng để dễ dàng theo dõi khi cuộn.
 */
import React from 'react';

export default function TimesheetTable({ data }) {
    if (!Array.isArray(data) || data.length === 0) {
        return <p className="text-center text-gray-500 mt-4">Không có dữ liệu chấm công cho tháng này.</p>;
    }

    const totals = data.reduce((acc, row) => {
        acc.hoursWorked += row.hoursWorked || 0;
        acc.leaveHours += row.leaveHours || 0;
        acc.h1 += row.h1 || 0;
        acc.h2 += row.h2 || 0;
        acc.h3 += row.h3 || 0;
        acc.b3 += row.b3 || 0;
        acc.b4 += row.b4 || 0;
        return acc;
    }, { hoursWorked: 0, leaveHours: 0, h1: 0, h2: 0, h3: 0, b3: 0, b4: 0 });

    const formatCell = (value) => !value || value === 0 ? '-' : value;
    const formatHoursCell = (value) => !value || value === 0 ? '-' : value.toFixed(1);

    const getStatusClass = (status) => {
        switch (status) {
            case 'Đi làm': return 'text-green-700 bg-green-50';
            case 'Nghỉ phép': return 'text-blue-700 bg-blue-50';
            case 'Đi làm & Nghỉ phép': return 'text-purple-700 bg-purple-50';
            default: return 'text-gray-700 bg-gray-50';
        }
    };

    return (
        // SỬA ĐỔI: Thêm container với chiều cao tối đa và thanh cuộn
        <div className="overflow-auto mt-4" style={{ maxHeight: '70vh' }}>
            <table className="min-w-full bg-white border border-gray-200">
                <thead>
                    <tr>
                        {/* SỬA ĐỔI: Thêm class `sticky` để cố định tiêu đề */}
                        <th className="sticky top-0 z-10 px-4 py-2 text-left text-xs text-gray-800 uppercase tracking-wider font-bold bg-gray-50 border-b">Ngày</th>
                        <th className="sticky top-0 z-10 px-4 py-2 text-left text-xs text-gray-800 uppercase tracking-wider font-bold bg-gray-50 border-b">Giờ vào</th>
                        <th className="sticky top-0 z-10 px-4 py-2 text-left text-xs text-gray-800 uppercase tracking-wider font-bold bg-gray-50 border-b">Giờ ra</th>
                        <th className="sticky top-0 z-10 px-4 py-2 text-left text-xs text-gray-800 uppercase tracking-wider font-bold bg-gray-50 border-b">Số giờ</th>
                        <th className="sticky top-0 z-10 px-4 py-2 text-left text-xs text-gray-800 uppercase tracking-wider font-bold bg-gray-50 border-b">Tăng ca 1.5</th>
                        <th className="sticky top-0 z-10 px-4 py-2 text-left text-xs text-gray-800 uppercase tracking-wider font-bold bg-gray-50 border-b">Tăng ca 2.0</th>
                        <th className="sticky top-0 z-10 px-4 py-2 text-left text-xs text-gray-800 uppercase tracking-wider font-bold bg-gray-50 border-b">Tăng ca 3.0</th>
                        <th className="sticky top-0 z-10 px-4 py-2 text-left text-xs text-gray-800 uppercase tracking-wider font-bold bg-gray-50 border-b">Tăng ca đêm</th>
                        <th className="sticky top-0 z-10 px-4 py-2 text-left text-xs text-gray-800 uppercase tracking-wider font-bold bg-gray-50 border-b">Phụ cấp 0.5</th>
                        <th className="sticky top-0 z-10 px-4 py-2 text-left text-xs text-gray-800 uppercase tracking-wider font-bold bg-gray-50 border-b">Trạng thái</th>
                        <th className="sticky top-0 z-10 px-4 py-2 text-left text-xs text-blue-800 uppercase tracking-wider font-bold bg-gray-50 border-b">Giờ Phép</th>
                        <th className="sticky top-0 z-10 px-4 py-2 text-left text-xs text-blue-800 uppercase tracking-wider font-bold bg-gray-50 border-b">Loại Phép</th>
                    </tr>
                </thead>
                <tbody className="divide-y divide-gray-200">
                    {data.map((row, index) => (
                        <tr key={index} className="hover:bg-gray-50">
                            <td className="px-4 py-2 whitespace-nowrap text-sm text-gray-800 font-bold">{new Date(row.date).toLocaleDateString('vi-VN')}</td>
                            <td className="px-4 py-2 whitespace-nowrap text-sm text-gray-500">{row.checkIn || '-'}</td>
                            <td className="px-4 py-2 whitespace-nowrap text-sm text-gray-500">{row.checkOut || '-'}</td>
                            <td className="px-4 py-2 whitespace-nowrap text-sm text-gray-500">{formatHoursCell(row.hoursWorked)}</td>
                            <td className="px-4 py-2 whitespace-nowrap text-sm text-gray-500">{formatCell(row.h1)}</td>
                            <td className="px-4 py-2 whitespace-nowrap text-sm text-gray-500">{formatCell(row.h2)}</td>
                            <td className="px-4 py-2 whitespace-nowrap text-sm text-gray-500">{formatCell(row.h3)}</td>
                            <td className="px-4 py-2 whitespace-nowrap text-sm text-gray-500">{formatCell(row.b3)}</td>
                            <td className="px-4 py-2 whitespace-nowrap text-sm text-gray-500">{formatCell(row.b4)}</td>
                            <td className="px-4 py-2 whitespace-nowrap text-sm font-medium">
                                <span className={`px-2 inline-flex text-xs leading-5 font-semibold rounded-full ${getStatusClass(row.status)}`}>
                                    {row.status}
                                </span>
                            </td>
                            <td className="px-4 py-2 whitespace-nowrap text-sm text-blue-600 font-semibold">{formatCell(row.leaveHours)}</td>
                            <td className="px-4 py-2 whitespace-nowrap text-sm text-blue-600">{row.leaveType || '-'}</td>
                        </tr>
                    ))}
                </tbody>
                <tfoot className="bg-gray-100 font-bold">
                    <tr>
                        <td colSpan="3" className="px-4 py-2 text-right text-gray-700">Tổng cộng</td>
                        <td className="px-4 py-2 text-gray-800">{totals.hoursWorked.toFixed(1)}</td>
                        <td className="px-4 py-2 text-gray-800">{totals.h1}</td>
                        <td className="px-4 py-2 text-gray-800">{totals.h2}</td>
                        <td className="px-4 py-2 text-gray-800">{totals.h3}</td>
                        <td className="px-4 py-2 text-gray-800">{totals.b3}</td>
                        <td className="px-4 py-2 text-gray-800">{totals.b4}</td>
                        <td className="px-4 py-2"></td>
                        <td className="px-4 py-2 text-blue-800">{totals.leaveHours.toFixed(1)}</td>
                        <td className="px-4 py-2"></td>
                    </tr>
                </tfoot>
            </table>
        </div>
    );
}
