/*
 * File: components/leave/HolidayTable.js
 * Mô tả: Component hiển thị thông tin tổng quan và chi tiết ngày nghỉ phép.
 */
import React from 'react';

export default function HolidayTable({ data, summary }) {
    return (
        <div className="mt-4">
            {summary.isCurrentYear ? (
                <div className="bg-blue-50 border-l-4 border-blue-500 text-blue-800 p-4 rounded-r-lg mb-6">
                    <p className="font-bold">Phép năm còn lại tính tới tháng hiện tại</p>
                    <p className="text-3xl font-bold">{summary?.remaining || 0} Giờ </p>
                    <p className="text-sm mt-1">{summary.isForeigner ? 'Chế độ: Lao động nước ngoài (16 giờ/tháng)' : ''}</p>
                </div>
            ) : (
                <div className="bg-blue-50 border-l-4 border-blue-500 text-blue-800 p-4 rounded-r-lg mb-6">
                    <p className="font-bold">Việc tính toán phép năm chỉ áp dụng cho năm hiện tại.</p>
                </div>
            )}
            {data.length === 0 ? (
                 <p className="text-center text-gray-500 mt-4">Không có dữ liệu chi tiết ngày nghỉ cho năm này.</p>
            ) : (
                <div className="overflow-x-auto">
                    <table className="min-w-full bg-white border border-gray-200">
                        <thead className="bg-gray-50">
                            <tr>
                                <th className="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Ngày nghỉ</th>
                                <th className="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Số giờ nghỉ</th>
                                <th className="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Loại nghỉ phép</th>
                            </tr>
                        </thead>
                        <tbody className="divide-y divide-gray-200">
                            {data.map((row, index) => (
                                <tr key={index} className="hover:bg-gray-50">
                                    <td className="px-4 py-2 whitespace-nowrap text-sm text-gray-800">{new Date(row.date).toLocaleDateString('vi-VN')}</td>
                                    <td className="px-4 py-2 whitespace-nowrap text-sm text-gray-500">{row.hours} giờ</td>
                                    <td className="px-4 py-2 whitespace-nowrap text-sm text-gray-500">
                                        {row.reason}
                                        {row.memo && row.memo.trim().toLowerCase() === 'khang cong' && (
                                            <span className="ml-1 font-semibold italic text-indigo-700">
                                                ({row.memo}) 🌟
                                            </span>
                                        )}
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
