/*
 * File: components/admin/ApprovePayroll.js
 * Mô tả: Chức năng duyệt bảng lương hàng tháng.
 */
import React, { useState, useEffect, useCallback } from 'react';
import { apiFetchAllPayrolls, apiApprovePayroll, apiGetGroups, apiExportPayrolls } from '../../api';

export default function ApprovePayrollComponent() {
    const [currentDate, setCurrentDate] = useState(new Date());
    const [payrolls, setPayrolls] = useState([]);
    const [isApproved, setIsApproved] = useState(false);
    const [isLoading, setIsLoading] = useState(false);
    const [groups, setGroups] = useState([]);
    const [selectedGroupId, setSelectedGroupId] = useState('ALL');
    const [isExporting, setIsExporting] = useState(false);

    const formatCurrency = (amount) => (amount || 0).toLocaleString('vi-VN', { style: 'currency', currency: 'VND' });
    const yearMonth = `${currentDate.getFullYear()}-${(currentDate.getMonth() + 1).toString().padStart(2, '0')}`;

    const fetchData = useCallback(async () => {
        setIsLoading(true);
        try {
            const { payrolls, isApproved } = await apiFetchAllPayrolls(yearMonth, selectedGroupId);
            setPayrolls(payrolls);
            setIsApproved(isApproved);
        } catch (error) { console.error(error); } finally {
            setIsLoading(false);
        }
    }, [yearMonth, selectedGroupId]);

    useEffect(() => {
        fetchData();
    }, [fetchData]);

    useEffect(() => {
        const fetchGroups = async () => {
            try {
                const data = await apiGetGroups();
                setGroups(data);
            } catch (error) {
                console.error("Lỗi khi tải danh sách bộ phận:", error);
            }
        };
        fetchGroups();
    }, []);

    const handleApprove = async () => {
        if (!window.confirm(`Bạn có chắc muốn phê duyệt lương tháng ${yearMonth}? Hành động này không thể hoàn tác.`)) return;
        try {
            await apiApprovePayroll(yearMonth);
            setIsApproved(true);
            alert("Phê duyệt thành công!");
        } catch (error) {
            alert("Lỗi: " + error.message);
        }
    };

    const handleExportPayroll = async () => {
        setIsExporting(true);
        await apiExportPayrolls(yearMonth, selectedGroupId);
        setIsExporting(false);
    };

    return (
        <div>
             <h2 className="text-2xl font-bold text-gray-800 mb-4">Phê duyệt Bảng lương</h2>
             <div className="flex flex-col md:flex-row md:items-center md:justify-between gap-4 mb-4">
                 <div className="flex items-center gap-4">
                    <input type="month" value={yearMonth} onChange={(e) => setCurrentDate(new Date(e.target.value))} className="px-3 py-2 border border-gray-300 rounded-md shadow-sm"/>
                    <select value={selectedGroupId} onChange={e => setSelectedGroupId(e.target.value)} className="block w-full pl-3 pr-10 py-2 text-base border-gray-300 focus:outline-none focus:ring-indigo-500 focus:border-indigo-500 sm:text-sm rounded-md">
                        <option value="ALL">Tất cả bộ phận</option>
                        {groups.map(group => (<option key={group.groupId} value={group.groupId}>{group.groupId} - {group.groupName}</option>))}
                    </select>
                 </div>
                 <div className="flex items-center gap-2">
                    <span className={`px-3 py-1 text-sm font-semibold rounded-full ${isApproved ? 'bg-green-100 text-green-800' : 'bg-yellow-100 text-yellow-800'}`}>
                        {isApproved ? 'ĐÃ PHÊ DUYỆT' : 'CHƯA PHÊ DUYỆT'}
                    </span>
                     <button onClick={handleExportPayroll} disabled={isExporting} className="px-4 py-2 text-sm font-medium text-white bg-blue-600 rounded-md hover:bg-blue-700 disabled:bg-gray-400">
                         {isExporting ? 'Đang xuất...' : 'Xuất Excel'}
                     </button>
                     <button onClick={handleApprove} disabled={isApproved} className="px-4 py-2 text-sm font-medium text-white bg-green-600 rounded-md hover:bg-green-700 disabled:bg-gray-400">
                         Phê duyệt
                     </button>
                 </div>
            </div>
            {isLoading ? <p>Đang tải...</p> : (
                <div className="overflow-x-auto">
                     <table className="min-w-full bg-white border">
                        <thead className="bg-gray-50"><tr className="bg-gray-50"><th className="px-4 py-2 text-left">Mã Bộ Phận</th><th className="px-4 py-2 text-left">Tên Bộ Phận</th><th className="px-4 py-2 text-left">Mã NV</th><th className="px-4 py-2 text-left">Tên Nhân Viên</th><th className="px-4 py-2 text-right">Lương Thực Lãnh</th></tr></thead>
                        <tbody className="divide-y divide-gray-200">
                            {payrolls.map(p => (
                                <tr key={p.EMPID} className="hover:bg-gray-50">
                                    <td className="px-4 py-2">{p.GROUPID}</td>
                                    <td className="px-4 py-2">{p.GroupName}</td>
                                    <td className="px-4 py-2">{p.EMPID}</td>
                                    <td className="px-4 py-2">{p.EMPNAM_VN}</td>
                                    <td className="px-4 py-2 text-right font-semibold">{formatCurrency(p.REAL_TOTAL)}</td>
                                </tr>
                            ))}
                        </tbody>
                     </table>
                </div>
            )}
        </div>
    );
}
