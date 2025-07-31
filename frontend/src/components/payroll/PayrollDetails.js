/*
 * File: components/payroll/PayrollDetails.js
 * Mô tả: Component hiển thị phiếu lương chi tiết.
 */
import React from 'react';

export default function PayrollDetails({ data, isAdminView = false }) {
    if (!data) return <p className="text-center text-gray-500 mt-4">Không có dữ liệu lương cho tháng này.</p>;
    if (data.approved === false && !isAdminView) {
        return <p className="text-center text-blue-600 bg-blue-50 p-4 rounded-md mt-4">{data.message}</p>;
    }
    const { employeeInfo = {}, earnings = [], deductions = [], overtimeAndBonus = [], summary = {} } = data;
    const formatCurrency = (amount) => {
        if (typeof amount !== 'number' || isNaN(amount)) return '0 ₫';
        return amount.toLocaleString('vi-VN', { style: 'currency', currency: 'VND' });
    };
    const totalEarnings = earnings.reduce((sum, item) => sum + (item.value || 0), 0);
    const totalDeductions = deductions.reduce((sum, item) => sum + (item.value || 0), 0);
    const totalOvertimeAndBonus = overtimeAndBonus.reduce((sum, item) => sum + (item.amount || 0), 0);
    
    const DetailRow = ({ label, value, colorClass = 'text-gray-800' }) => (
        <div className="flex justify-between items-center py-2 border-b border-gray-100">
            <p className="text-sm text-gray-600">{label}</p>
            <p className={`text-sm font-medium ${colorClass}`}>{formatCurrency(value)}</p>
        </div>
    );

    const OvertimeBonusTable = ({ data }) => {
        if (!data || data.length === 0) return null;
        return (
            <div className="mt-6">
                <h3 className="text-lg font-bold text-gray-700 mb-2">Chi tiết Tăng ca & Thưởng</h3>
                <div className="overflow-x-auto bg-gray-50 p-4 rounded-lg">
                    <table className="min-w-full">
                        <thead>
                            <tr>
                                <th className="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">Hạng mục</th>
                                <th className="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">Số giờ</th>
                                <th className="px-4 py-2 text-right text-xs font-medium text-gray-500 uppercase">Thành tiền</th>
                            </tr>
                        </thead>
                        <tbody className="divide-y divide-gray-200">
                            {data.map((item, index) => item.amount > 0 && (
                                <tr key={index}>
                                    <td className="px-4 py-2 text-sm text-gray-800">{item.label}</td>
                                    <td className="px-4 py-2 text-sm text-gray-500">{item.hours}</td>
                                    <td className="px-4 py-2 text-sm text-gray-800 text-right">{formatCurrency(item.amount)}</td>
                                </tr>
                            ))}
                        </tbody>
                        <tfoot className="border-t-2 border-gray-300">
                             <tr>
                                <td colSpan="2" className="px-4 py-2 text-right font-bold text-gray-700">Tổng cộng:</td>
                                <td className="px-4 py-2 text-right font-bold text-gray-800">{formatCurrency(totalOvertimeAndBonus)}</td>
                            </tr>
                        </tfoot>
                    </table>
                </div>
            </div>
        );
    }

    return (
        <div className="mt-6 bg-white p-6 rounded-2xl shadow-lg font-sans transition-all duration-300">
            <div className="text-center mb-6">
                <h2 className="text-2xl font-bold text-gray-800">BẢNG LƯƠNG CHI TIẾT</h2>
                <p className="text-md text-gray-500">Kỳ lương: Tháng {employeeInfo.thang}/{employeeInfo.nam}</p>
            </div>
            <div className="bg-gray-50 p-4 rounded-lg mb-6">
                <div className="grid grid-cols-3 gap-x-4 gap-y-2 text-sm">
                    <div><p className="font-semibold text-gray-500">NHÂN VIÊN</p><p className="font-bold text-gray-800">{employeeInfo.hoTen}</p></div>
                    <div><p className="font-semibold text-gray-500">MÃ SỐ</p><p className="font-bold text-gray-800">{employeeInfo.soThe}</p></div>
                     <div><p className="font-semibold text-gray-500">TIỀN CÔNG GIỜ</p><p className="font-bold text-gray-800">{formatCurrency(summary.tinhLuongMoiGio)}</p></div>
                    <div><p className="font-semibold text-gray-500">CHỨC VỤ</p><p className="font-bold text-gray-800">{employeeInfo.chucVu}</p></div>
                    <div><p className="font-semibold text-gray-500">ĐƠN VỊ</p><p className="font-bold text-gray-800">{employeeInfo.donVi}</p></div>
                </div>
            </div>
            <div className="bg-indigo-600 text-white p-6 rounded-xl text-center mb-6 shadow-indigo-200 shadow-md">
                <p className="text-lg font-semibold opacity-80">LƯƠNG THỰC LÃNH</p>
                <p className="text-4xl font-bold tracking-tight">{formatCurrency(summary.luongThucLanh)}</p>
            </div>
            <div className="grid md:grid-cols-2 gap-6">
                <div className="bg-green-50 p-4 rounded-lg">
                    <h3 className="font-bold text-green-800 mb-3">CÁC KHOẢN THU NHẬP</h3>
                    <div className="space-y-1">
                        {earnings.map((item, index) => (<DetailRow key={index} label={item.label} value={item.value} />))}
                         <div className="pt-2 mt-2 border-t-2 border-green-200">
                            <DetailRow label="TỔNG THU NHẬP (chưa gồm Tăng ca)" value={totalEarnings} colorClass="text-green-700 font-bold" />
                         </div>
                    </div>
                </div>
                <div className="bg-red-50 p-4 rounded-lg">
                    <h3 className="font-bold text-red-800 mb-3">CÁC KHOẢN KHẤU TRỪ</h3>
                    <div className="space-y-1">
                        {deductions.map((item, index) => (<DetailRow key={index} label={item.label} value={item.value} />))}
                        <div className="pt-2 mt-2 border-t-2 border-red-200">
                            <DetailRow label="TỔNG KHẤU TRỪ" value={totalDeductions} colorClass="text-red-700 font-bold" />
                        </div>
                    </div>
                </div>
            </div>
            <OvertimeBonusTable data={overtimeAndBonus} />
        </div>
    );
}
