/*
 * File: components/timesheet/TimesheetQuery.js
 * Mô tả: Giao diện tra cứu và chỉnh sửa dữ liệu chấm công.
 * Cập nhật: Cố định dòng tiêu đề của bảng để dễ dàng theo dõi khi cuộn.
 */
import React, { useState, useEffect, useCallback } from 'react';
import { useLocation, useNavigate } from 'react-router-dom';
import { apiGetTimesheetForEmployee, apiUpdateTimesheetEntry, apiGetEmployeeInfo } from '../../api';

// --- HÀM HỖ TRỢ TÍNH TOÁN (Không thay đổi) ---
const roundUpToHour = (timeStr, baseDate) => {
    if (!timeStr) return null;
    const [hours, minutes] = timeStr.split(':').map(Number);
    const date = new Date(baseDate);
    date.setHours(hours, minutes, 0, 0);
    if (minutes > 0) {
        date.setHours(date.getHours() + 1, 0, 0, 0);
    }
    return date;
};

const calculateNightAllowance_0_3 = (checkInStr, workDate, isNightShift, standardHours) => {
    if (!isNightShift || !checkInStr || standardHours <= 0) return 0;
    const baseDate = new Date(workDate);
    baseDate.setHours(0, 0, 0, 0);
    const officialStartTime = roundUpToHour(checkInStr, baseDate);
    if (!officialStartTime) return 0;
    const officialWorkEnd = new Date(officialStartTime.getTime() + standardHours * 60 * 60 * 1000);
    const nightWindowStart = new Date(baseDate);
    nightWindowStart.setHours(22, 0, 0, 0);
    const nightWindowEnd = new Date(baseDate);
    nightWindowEnd.setDate(nightWindowEnd.getDate() + 1);
    nightWindowEnd.setHours(6, 0, 0, 0);
    const overlapStart = Math.max(officialStartTime, nightWindowStart);
    const overlapEnd = Math.min(officialWorkEnd, nightWindowEnd);
    const diffMs = overlapEnd - overlapStart;
    if (diffMs > 0) {
        return parseFloat((diffMs / (1000 * 60 * 60)).toFixed(2));
    }
    return 0;
};


export default function TimesheetQuery() {
    const location = useLocation();
    const navigate = useNavigate();

    const getCurrentYearMonth = () => {
        const now = new Date();
        const year = now.getFullYear();
        const month = (now.getMonth() + 1).toString().padStart(2, '0');
        return `${year}-${month}`;
    };

    const [searchParams, setSearchParams] = useState({ empid: '', yymm: getCurrentYearMonth() });
    const [employeeInfo, setEmployeeInfo] = useState(null);
    const [results, setResults] = useState([]);
    const [originalResults, setOriginalResults] = useState([]);
    const [editedRows, setEditedRows] = useState({});
    const [isLoading, setIsLoading] = useState(false);
    const [message, setMessage] = useState({ type: '', text: '' });
    
    const SPECIAL_ALLOWANCE_GROUPS = ['A059', 'A073', 'A067'];

    const executeSearch = useCallback(async (empIdToSearch, yearMonthToSearch) => {
        if (!empIdToSearch || !yearMonthToSearch) {
            setMessage({ type: 'error', text: 'Vui lòng nhập Mã NV và Tháng/Năm.' });
            return;
        }
        setIsLoading(true);
        setMessage({ type: '', text: '' });
        setResults([]);
        setOriginalResults([]);
        setEditedRows({});
        setEmployeeInfo(null);

        try {
            const info = await apiGetEmployeeInfo(empIdToSearch);
            setEmployeeInfo(info);
            const data = await apiGetTimesheetForEmployee(empIdToSearch, yearMonthToSearch.replace('-', ''));
            if (data.length === 0) {
                setMessage({ type: 'info', text: 'Không tìm thấy dữ liệu chấm công.' });
            }
            setResults(data);
            setOriginalResults(JSON.parse(JSON.stringify(data))); 
        } catch (err) {
            setEmployeeInfo(null);
            setMessage({ type: 'error', text: err.message });
        } finally {
            setIsLoading(false);
        }
    }, []);

    useEffect(() => {
        if (location.state?.empid && location.state?.yymm) {
            const { empid, yymm } = location.state;
            navigate(location.pathname, { replace: true, state: null });
            setSearchParams({ empid, yymm });
            executeSearch(empid, yymm);
        }
    }, [location.state, executeSearch, navigate, location.pathname]);


    const handleParamChange = (e) => {
        const { name, value } = e.target;
        setSearchParams(prev => ({ ...prev, [name]: value.toUpperCase() }));
        if (name === 'empid') {
            setEmployeeInfo(null);
        }
    };

    const handleFormSubmit = (e) => {
        e.preventDefault();
        executeSearch(searchParams.empid, searchParams.yymm);
    };

    const handleInputChange = (autoid, field, value) => {
        let newResults = results.map(row => 
            row.autoid === autoid ? { ...row, [field]: value } : row
        );

        const changedRow = newResults.find(r => r.autoid === autoid);

        if (['timeup', 'timedown', 'Kzhour'].includes(field) && changedRow) {
            const workDate = new Date(
                changedRow.workdat.substring(0, 4), 
                parseInt(changedRow.workdat.substring(4, 6), 10) - 1, 
                changedRow.workdat.substring(6, 8)
            );
            const timeup = formatTime(changedRow.timeup);
            const timedown = formatTime(changedRow.timedown);
            const isNightShift = timeup > timedown;
            const standardHours = parseFloat(changedRow.Kzhour) || 0;

            let phuCap_0_5 = 0;
            if (employeeInfo && SPECIAL_ALLOWANCE_GROUPS.includes(employeeInfo.GROUPID) && standardHours >= 8) {
                phuCap_0_5 = 0.5;
            }
            changedRow.B4 = phuCap_0_5;

            changedRow.B5 = calculateNightAllowance_0_3(timeup, workDate, isNightShift, standardHours);
        }

        setResults(newResults);
        setEditedRows(prev => ({ ...prev, [autoid]: true }));
    };

    const handleSaveAll = async () => {
        const changedData = results.filter(row => editedRows[row.autoid]);
        if (changedData.length === 0) return;

        setIsLoading(true);
        setMessage({ type: 'info', text: `Đang lưu ${changedData.length} thay đổi...` });

        try {
            await Promise.all(changedData.map(row => apiUpdateTimesheetEntry(row.autoid, row)));
            setMessage({ type: 'success', text: 'Tất cả thay đổi đã được lưu thành công!' });
            setOriginalResults(JSON.parse(JSON.stringify(results)));
            setEditedRows({});
        } catch (err) {
            setMessage({ type: 'error', text: `Lỗi khi lưu thay đổi: ${err.message}` });
        } finally {
            setIsLoading(false);
        }
    };

    const handleCancelChanges = () => {
        setResults(originalResults);
        setEditedRows({});
        setMessage({ type: '', text: '' });
    };
    
    const formatTime = (timeStr) => {
        if (!timeStr || timeStr.length < 4) return '';
        return `${timeStr.substring(0, 2)}:${timeStr.substring(2, 4)}`;
    };

    return (
        <div className="max-w-7xl mx-auto">
            <h2 className="text-2xl font-bold text-gray-800 mb-4">Tra cứu & Chỉnh sửa Chấm công</h2>
            <form onSubmit={handleFormSubmit} className="bg-gray-50 p-4 rounded-lg flex items-end gap-4 mb-6">
                <div>
                    <label className="block text-sm font-medium text-gray-700">Mã Nhân viên</label>
                    <input type="text" name="empid" value={searchParams.empid} onChange={handleParamChange} className="mt-1 p-2 border rounded-md w-40"/>
                </div>
                {employeeInfo && <div className="self-center pb-2 text-blue-600 font-semibold">{employeeInfo.EMPNAM_VN}</div>}
                <div>
                    <label className="block text-sm font-medium text-gray-700">Tháng (YYYY-MM)</label>
                    <input type="month" name="yymm" value={searchParams.yymm} onChange={handleParamChange} className="mt-1 p-2 border rounded-md"/>
                </div>
                <button type="submit" disabled={isLoading} className="px-6 py-2 font-semibold text-white bg-blue-600 rounded-md hover:bg-blue-700 disabled:bg-gray-400">
                    {isLoading ? 'Đang tìm...' : 'Tìm kiếm'}
                </button>
            </form>

            {message.text && (
                <div className={`p-3 rounded-md text-sm ${
                    message.type === 'error' ? 'bg-red-100 text-red-800' : 
                    message.type === 'success' ? 'bg-green-100 text-green-800' : 'bg-blue-100 text-blue-800'
                }`}>
                    {message.text}
                </div>
            )}

            {results.length > 0 && (
                 <div className="mt-6">
                    {/* SỬA ĐỔI: Thêm container với chiều cao tối đa và thanh cuộn */}
                    <div className="overflow-auto bg-white rounded-lg shadow" style={{ maxHeight: '70vh' }}>
                        <table className="min-w-full text-sm">
                            <thead>
                                <tr>
                                    {['Ngày', 'Giờ vào', 'Giờ ra', 'Giờ vắng', 'Tổng giờ', 'TC 1.5', 'TC 2.0', 'TC 3.0', 'TC Đêm', 'PC 0.5', 'PC 0.3'].map(header => (
                                        // SỬA ĐỔI: Thêm class `sticky` để cố định tiêu đề
                                        <th key={header} className="sticky top-0 z-10 p-2 text-left font-semibold text-gray-600 bg-gray-100 whitespace-nowrap border-b">
                                            {header}
                                        </th>
                                    ))}
                                </tr>
                            </thead>
                            <tbody className="divide-y divide-gray-200">
                                {results.map(row => (
                                    <tr key={row.autoid} className={`hover:bg-gray-50 ${editedRows[row.autoid] ? 'bg-yellow-50' : ''}`}>
                                        <td className="p-2 whitespace-nowrap">{row.workdat}</td>
                                        <td className="p-1"><input type="text" value={formatTime(row.timeup)} onChange={e => handleInputChange(row.autoid, 'timeup', e.target.value.replace(':', '') + '00')} className="w-20 p-1 border rounded-md bg-white"/></td>
                                        <td className="p-1"><input type="text" value={formatTime(row.timedown)} onChange={e => handleInputChange(row.autoid, 'timedown', e.target.value.replace(':', '') + '00')} className="w-20 p-1 border rounded-md bg-white"/></td>
                                        <td className="p-1"><input type="number" step="0.5" value={row.Kzhour} onChange={e => handleInputChange(row.autoid, 'Kzhour', e.target.value)} className="w-16 p-1 border rounded-md"/></td>
                                        <td className="p-1"><input type="number" step="0.5" value={row.TOTH} onChange={e => handleInputChange(row.autoid, 'TOTH', e.target.value)} className="w-16 p-1 border rounded-md"/></td>
                                        <td className="p-1"><input type="number" step="0.5" value={row.H1} onChange={e => handleInputChange(row.autoid, 'H1', e.target.value)} className="w-16 p-1 border rounded-md"/></td>
                                        <td className="p-1"><input type="number" step="0.5" value={row.H2} onChange={e => handleInputChange(row.autoid, 'H2', e.target.value)} className="w-16 p-1 border rounded-md"/></td>
                                        <td className="p-1"><input type="number" step="0.5" value={row.H3} onChange={e => handleInputChange(row.autoid, 'H3', e.target.value)} className="w-16 p-1 border rounded-md"/></td>
                                        <td className="p-1"><input type="number" step="0.5" value={row.B3} onChange={e => handleInputChange(row.autoid, 'B3', e.target.value)} className="w-16 p-1 border rounded-md"/></td>
                                        <td className="p-1"><input type="number" step="0.5" value={row.B4} onChange={e => handleInputChange(row.autoid, 'B4', e.target.value)} className="w-16 p-1 border rounded-md"/></td>
                                        <td className="p-1"><input type="number" step="0.5" value={row.B5} onChange={e => handleInputChange(row.autoid, 'B5', e.target.value)} className="w-16 p-1 border rounded-md"/></td>
                                    </tr>
                                ))}
                            </tbody>
                        </table>
                    </div>
                    {Object.keys(editedRows).length > 0 && (
                        <div className="mt-6 flex justify-end gap-4">
                            <button onClick={handleCancelChanges} type="button" className="px-6 py-2 font-semibold text-gray-700 bg-gray-200 rounded-md hover:bg-gray-300">
                                Hủy
                            </button>
                            <button onClick={handleSaveAll} disabled={isLoading} className="px-6 py-2 font-semibold text-white bg-green-600 rounded-md hover:bg-green-700 disabled:bg-gray-400">
                                {isLoading ? 'Đang lưu...' : `Lưu ${Object.keys(editedRows).length} thay đổi`}
                            </button>
                        </div>
                    )}
                </div>
            )}
        </div>
    );
}
