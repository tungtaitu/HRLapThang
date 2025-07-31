/*
 * File: components/admin/LeaveManager.js
 * Mô tả: Giao diện hợp nhất cho phép admin quản lý toàn bộ nghiệp vụ nghỉ phép.
 * Cập nhật: Cập nhật logic xử lý giờ vắng (Kzhour) linh hoạt theo số giờ phép.
 */
import React, { useState, useCallback, useEffect, useRef } from 'react';
import { 
    apiGetLeaveEntries, 
    apiDeleteLeaveEntry, 
    apiGetEmployeeInfo, 
    apiAdminSubmitLeave, 
    apiFetchHolidays,
    apiUpdateLeaveEntry,
    apiGetTimesheetForEmployee,
    apiUpdateTimesheetEntry
} from '../../api';

// --- DANH SÁCH CÁC LOẠI PHÉP ---
const LEAVE_TYPE_OPTIONS = {
    'E': 'Phép năm',
    'A': 'Việc riêng',
    'B': 'Phép Bệnh',
    'C': 'Nghỉ kết hôn',
    'D': 'Phép Tang',
    'F': 'Nghỉ thai sản',
    'G': 'Nghỉ công tác',
    'H': 'Nghỉ C.Thường',
    'I': 'Đi đường',
    'K': 'Không lương'
};

// --- COMPONENT MODAL CHỈNH SỬA ---
const EditModal = ({ entry, onSave, onCancel, isLoading }) => {
    const [formData, setFormData] = useState({
        DateUP: new Date(entry.DateUP).toISOString().split('T')[0], 
        TimeUP: entry.TimeUP.replace(':', ''),
        TimeDown: entry.TimeDown.replace(':', ''),
        JiaType: entry.JiaType.trim(),
        memo: entry.memo || ''
    });

    const handleChange = (e) => {
        const { name, value } = e.target;
        setFormData(prev => ({ ...prev, [name]: value }));
    };

    const handleSubmit = (e) => {
        e.preventDefault();
        const payload = {
            ...formData,
            TimeUP: `${formData.TimeUP.substring(0, 2)}:${formData.TimeUP.substring(2, 4)}`,
            TimeDown: `${formData.TimeDown.substring(0, 2)}:${formData.TimeDown.substring(2, 4)}`
        };
        onSave(entry.id, payload);
    };

    return (
        <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50 p-4">
            <form onSubmit={handleSubmit} className="bg-white rounded-lg shadow-xl p-6 w-full max-w-md space-y-4">
                <h2 className="text-xl font-bold mb-4">Chỉnh sửa ngày phép</h2>
                <div>
                    <label className="block text-sm font-medium">Ngày nghỉ</label>
                    <input type="date" name="DateUP" value={formData.DateUP} onChange={handleChange} className="mt-1 w-full px-3 py-2 border rounded-md" required />
                </div>
                <div className="grid grid-cols-2 gap-4">
                    <div>
                        <label className="block text-sm font-medium">Giờ bắt đầu (hhmm)</label>
                        <input type="text" name="TimeUP" value={formData.TimeUP} onChange={handleChange} maxLength={4} className="mt-1 w-full px-3 py-2 border rounded-md" required />
                    </div>
                    <div>
                        <label className="block text-sm font-medium">Giờ kết thúc (hhmm)</label>
                        <input type="text" name="TimeDown" value={formData.TimeDown} onChange={handleChange} maxLength={4} className="mt-1 w-full px-3 py-2 border rounded-md" required />
                    </div>
                </div>
                <div>
                   <label className="block text-sm font-medium">Loại phép</label>
                    <select name="JiaType" value={formData.JiaType} onChange={handleChange} className="mt-1 w-full px-3 py-2 border rounded-md bg-white">
                        {Object.entries(LEAVE_TYPE_OPTIONS).map(([key, value]) => (
                            <option key={key} value={key}>{key}: {value}</option>
                        ))}
                    </select>
                </div>
                <div>
                    <label className="block text-sm font-medium">Lý do (nếu có)</label>
                    <textarea name="memo" value={formData.memo} onChange={handleChange} rows="2" className="mt-1 w-full px-3 py-2 border rounded-md"/>
                </div>
                <div className="flex justify-end gap-4 pt-4">
                    <button type="button" onClick={onCancel} className="px-4 py-2 bg-gray-200 rounded-md">Hủy</button>
                    <button type="submit" disabled={isLoading} className="px-4 py-2 bg-blue-600 text-white rounded-md disabled:bg-blue-400">
                        {isLoading ? 'Đang lưu...' : 'Lưu thay đổi'}
                    </button>
                </div>
            </form>
        </div>
    );
};

const ConfirmationModal = ({ message, onConfirm, onCancel }) => (
    <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50 p-4">
        <div className="bg-white rounded-lg shadow-xl p-6 w-full max-w-sm">
            <h2 className="text-lg font-bold mb-4">Xác nhận hành động</h2>
            <p className="text-gray-600 mb-6">{message}</p>
            <div className="flex justify-end gap-4">
                <button onClick={onCancel} className="px-4 py-2 bg-gray-200 rounded-md hover:bg-gray-300">Hủy</button>
                <button onClick={onConfirm} className="px-4 py-2 bg-red-600 text-white rounded-md hover:bg-red-700">Xác nhận Xóa</button>
            </div>
        </div>
    </div>
);

const calculateHours = (startStr, endStr) => {
    if (!startStr || !endStr || startStr.length < 4 || endStr.length < 4) return 0;
    const startH = parseInt(startStr.substring(0, 2)), startM = parseInt(startStr.substring(2, 4));
    const endH = parseInt(endStr.substring(0, 2)), endM = parseInt(endStr.substring(2, 4));
    if (isNaN(startH) || isNaN(startM) || isNaN(endH) || isNaN(endM)) return 0;
    const start = new Date(1970, 0, 1, startH, startM);
    const end = new Date(1970, 0, 1, endH, endM);
    if (end <= start) return 0;
    const morningStart = new Date(1970, 0, 1, 8, 0), morningEnd = new Date(1970, 0, 1, 12, 0);
    const afternoonStart = new Date(1970, 0, 1, 13, 0), afternoonEnd = new Date(1970, 0, 1, 17, 0);
    let totalMs = 0;
    const morningOverlapStart = Math.max(start, morningStart), morningOverlapEnd = Math.min(end, morningEnd);
    if (morningOverlapEnd > morningOverlapStart) totalMs += morningOverlapEnd - morningOverlapStart;
    const afternoonOverlapStart = Math.max(start, afternoonStart), afternoonOverlapEnd = Math.min(end, afternoonEnd);
    if (afternoonOverlapEnd > afternoonOverlapStart) totalMs += afternoonOverlapEnd - afternoonOverlapStart;
    return Math.round((totalMs / 3600000) * 10) / 10;
};

const parseDateForFrontend = (ddmmyyyy) => {
    if (!ddmmyyyy || ddmmyyyy.length !== 8) return null;
    const day = parseInt(ddmmyyyy.substring(0, 2));
    const month = parseInt(ddmmyyyy.substring(2, 4)) - 1;
    const year = parseInt(ddmmyyyy.substring(4, 8));
    const date = new Date(year, month, day);
    if (date.getFullYear() !== year || date.getMonth() !== month || date.getDate() !== day) return null;
    return date;
};

const PUBLIC_HOLIDAYS = {
    '2025-01-01': 'Tết Dương lịch', '2025-01-28': 'Tết Nguyên Đán', '2025-01-29': 'Tết Nguyên Đán',
    '2025-01-30': 'Tết Nguyên Đán', '2025-01-31': 'Tết Nguyên Đán', '2025-02-03': 'Nghỉ bù Tết Nguyên Đán',
    '2025-04-08': 'Giỗ Tổ Hùng Vương', '2025-04-30': 'Ngày Giải phóng miền Nam', '2025-05-01': 'Ngày Quốc tế Lao động',
    '2025-05-02': 'Nghỉ bù', '2025-09-01': 'Nghỉ Quốc Khánh', '2025-09-02': 'Quốc Khánh',
};

export default function LeaveManager() {
    const [employeeId, setEmployeeId] = useState('');
    const [year, setYear] = useState(new Date().getFullYear());
    const [employeeInfo, setEmployeeInfo] = useState(null);
    const [leaveEntries, setLeaveEntries] = useState([]);
    const [formState, setFormState] = useState({ startDate: '', endDate: '', leaveType: 'E', startTime: '', endTime: '', reason: '' });
    const [isLoading, setIsLoading] = useState(false);
    const [message, setMessage] = useState({ type: '', text: '' });
    const [showConfirmModal, setShowConfirmModal] = useState(false);
    const [entryToDelete, setEntryToDelete] = useState(null);
    const [showEditModal, setShowEditModal] = useState(false);
    const [editingEntry, setEditingEntry] = useState(null);
    const [dailyBreakdown, setDailyBreakdown] = useState([]);

    const startDateRef = useRef(null);
    const endDateRef = useRef(null);
    const leaveTypeRef = useRef(null);
    const startTimeRef = useRef(null);
    const endTimeRef = useRef(null);
    const reasonRef = useRef(null);
    const submitButtonRef = useRef(null);

    const handleKeyDown = (e, nextFieldRef) => {
        if (e.key === 'Enter') {
            e.preventDefault();
            if (nextFieldRef && nextFieldRef.current) {
                nextFieldRef.current.focus();
            }
        }
    };

    const handleSearch = useCallback(async (e) => {
        if (e) e.preventDefault();
        if (!employeeId) {
            setMessage({ type: 'error', text: 'Vui lòng nhập Mã số nhân viên.' });
            return;
        }
        setIsLoading(true);
        setMessage({ type: '', text: '' });
        setLeaveEntries([]);
        setEmployeeInfo(null);
        try {
            const [empInfo, entriesData, summaryRes] = await Promise.all([
                apiGetEmployeeInfo(employeeId),
                apiGetLeaveEntries(employeeId, year),
                apiFetchHolidays(employeeId, new Date().getFullYear())
            ]);
            setEmployeeInfo({ name: empInfo.EMPNAM_VN, department: empInfo.department, ...summaryRes.summary });
            setLeaveEntries(entriesData);
        } catch (error) {
            setMessage({ type: 'error', text: error.message });
        } finally {
            setIsLoading(false);
        }
    }, [employeeId, year]);

    const handleSubmitLeave = async (e) => {
        e.preventDefault();
        setIsLoading(true);
        setMessage({ type: '', text: '' });

        try {
            const updatesToPerform = [];
            const monthsToFetch = new Set();
            const leaveDayDetails = new Map();

            dailyBreakdown.forEach(day => {
                if (day.hours > 0) {
                    const parts = day.date.split('/');
                    const dateObj = new Date(parts[2], parts[1] - 1, parts[0]);
                    const yyyymm = `${dateObj.getFullYear()}${(dateObj.getMonth() + 1).toString().padStart(2, '0')}`;
                    const yyyymmdd = `${yyyymm}${dateObj.getDate().toString().padStart(2, '0')}`;
                    
                    monthsToFetch.add(yyyymm);
                    leaveDayDetails.set(yyyymmdd, day.hours);
                }
            });

            if (monthsToFetch.size > 0) {
                const monthFetchPromises = Array.from(monthsToFetch).map(yymm =>
                    apiGetTimesheetForEmployee(employeeId, yymm)
                );
                const monthlyTimesheets = await Promise.all(monthFetchPromises);
                const allTimesheetEntries = monthlyTimesheets.flat();

                allTimesheetEntries.forEach(entry => {
                    if (leaveDayDetails.has(entry.workdat)) {
                        let id = entry.autoid || entry.AUTOID;
                        if (Array.isArray(id)) id = id[0];

                        if (id && !isNaN(parseInt(id)) && parseInt(id) > 0) {
                            // CẬP NHẬT LOGIC: Trừ giờ vắng bằng đúng số giờ phép nhập
                            const leaveHours = leaveDayDetails.get(entry.workdat);
                            const currentKzhour = parseFloat(entry.Kzhour || 0);
                            const newKzhour = Math.max(0, currentKzhour - leaveHours); // Đảm bảo không âm
                            
                            const updatedEntry = { ...entry, Kzhour: newKzhour };
                            updatesToPerform.push(apiUpdateTimesheetEntry(id, updatedEntry));
                        } else {
                            console.error("Bỏ qua cập nhật giờ vắng: bản ghi chấm công có 'autoid' không hợp lệ.", entry);
                        }
                    }
                });

                if (updatesToPerform.length > 0) {
                    await Promise.all(updatesToPerform);
                }
            }

            const leaveData = { userId: employeeId, ...formState, endDate: formState.endDate || formState.startDate };
            await apiAdminSubmitLeave(leaveData);

            setMessage({ type: 'success', text: 'Nhập phép thành công và giờ vắng đã được cập nhật!' });
            setFormState({ startDate: '', endDate: '', leaveType: 'E', startTime: '', endTime: '', reason: '' });
            await handleSearch();

        } catch (error) {
            setMessage({ type: 'error', text: `Lỗi khi nhập phép: ${error.message}` });
        } finally {
            setIsLoading(false);
        }
    };
    
    const handleEditClick = (entry) => {
        setEditingEntry(entry);
        setShowEditModal(true);
    };

    const handleSaveEdit = async (id, updatedData) => {
        setIsLoading(true);
        try {
            await apiUpdateLeaveEntry(id, updatedData);
            setMessage({ type: 'success', text: 'Cập nhật ngày phép thành công!' });
            setShowEditModal(false);
            setEditingEntry(null);
            await handleSearch();
        } catch (error) {
            setMessage({ type: 'error', text: `Lỗi khi cập nhật: ${error.message}` });
        } finally {
            setIsLoading(false);
        }
    };
    
    const handleDeleteClick = (entry) => {
        setEntryToDelete(entry);
        setShowConfirmModal(true);
    };

    const confirmDelete = async () => {
        if (!entryToDelete) return;
        setIsLoading(true);
        setShowConfirmModal(false);
        setMessage({ type: '', text: '' });

        try {
            const leaveDate = new Date(entryToDelete.DateUP);
            const yyyymm = `${leaveDate.getFullYear()}${(leaveDate.getMonth() + 1).toString().padStart(2, '0')}`;
            const yyyymmdd = `${yyyymm}${leaveDate.getDate().toString().padStart(2, '0')}`;

            const timesheetData = await apiGetTimesheetForEmployee(employeeId, yyyymm);
            const timesheetEntryForDay = timesheetData.find(entry => entry.workdat === yyyymmdd);

            if (timesheetEntryForDay) {
                let id = timesheetEntryForDay.autoid || timesheetEntryForDay.AUTOID;
                if (Array.isArray(id)) id = id[0];
                
                if (id && !isNaN(parseInt(id)) && parseInt(id) > 0) {
                    // CẬP NHẬT LOGIC: Cộng lại giờ vắng bằng đúng số giờ phép đã xóa
                    const leaveHoursToDelete = parseFloat(entryToDelete.HHour || 0);
                    const currentKzhour = parseFloat(timesheetEntryForDay.Kzhour || 0);
                    const newKzhour = currentKzhour + leaveHoursToDelete;

                    const updatedEntry = { ...timesheetEntryForDay, Kzhour: newKzhour };
                    await apiUpdateTimesheetEntry(id, updatedEntry);
                } else {
                    console.error("Bỏ qua cập nhật giờ vắng khi xóa phép: bản ghi chấm công không có 'autoid' hợp lệ.", timesheetEntryForDay);
                }
            }

            await apiDeleteLeaveEntry(entryToDelete.id);

            setMessage({ type: 'success', text: 'Xóa ngày phép thành công và đã cập nhật lại giờ vắng!' });
            await handleSearch();

        } catch (error) {
            setMessage({ type: 'error', text: `Lỗi khi xóa: ${error.message}` });
        } finally {
            setIsLoading(false);
            setEntryToDelete(null);
        }
    };

    const handleFormChange = (e) => {
        const { name, value } = e.target;
        setFormState(prev => ({ ...prev, [name]: value }));
    };

    useEffect(() => {
        const { startDate, endDate, startTime, endTime } = formState;
        const sDate = parseDateForFrontend(startDate);
        const eDate = parseDateForFrontend(endDate || startDate);
        if (!sDate || !eDate || eDate < sDate || !startTime || !endTime) {
            setDailyBreakdown([]);
            return;
        }
        let breakdown = [];
        let currentDate = new Date(sDate);
        while(currentDate <= eDate) {
            const dayOfWeek = currentDate.getDay();
            const dateString = `${currentDate.getFullYear()}-${String(currentDate.getMonth() + 1).padStart(2, '0')}-${String(currentDate.getDate()).padStart(2, '0')}`;
            let hours = 0;
            let note = '';
            if (dayOfWeek === 0) {
                note = 'Chủ Nhật';
            } else if (PUBLIC_HOLIDAYS[dateString]) {
                note = PUBLIC_HOLIDAYS[dateString];
            } else {
                const isSameDayPeriod = sDate.getTime() === eDate.getTime();
                const isFirstDay = currentDate.getTime() === sDate.getTime();
                const isLastDay = currentDate.getTime() === eDate.getTime();
                if (isSameDayPeriod) { hours = calculateHours(startTime, endTime); }
                else if (isFirstDay) { hours = calculateHours(startTime, '1700'); }
                else if (isLastDay) { hours = calculateHours('0800', endTime); }
                else { hours = 8; }
            }
            breakdown.push({ date: currentDate.toLocaleDateString('vi-VN'), hours, note });
            currentDate.setDate(currentDate.getDate() + 1);
        }
        setDailyBreakdown(breakdown);
    }, [formState.startDate, formState.endDate, formState.startTime, formState.endTime]);

    const totalHours = dailyBreakdown.reduce((acc, day) => acc + day.hours, 0);
    const years = Array.from({ length: 10 }, (_, i) => new Date().getFullYear() + 1 - i);

    return (
        <div>
            {showConfirmModal && ( <ConfirmationModal message={`Bạn có chắc muốn xóa ngày phép ${new Date(entryToDelete.DateUP).toLocaleDateString('vi-VN')}?`} onConfirm={confirmDelete} onCancel={() => setShowConfirmModal(false)} /> )}
            {showEditModal && ( <EditModal entry={editingEntry} onSave={handleSaveEdit} onCancel={() => setShowEditModal(false)} isLoading={isLoading} /> )}

            <h2 className="text-2xl font-bold text-gray-800 mb-4">Quản lý Phép Nhân viên</h2>
            
            <form onSubmit={handleSearch} className="bg-gray-50 p-4 rounded-lg flex flex-col sm:flex-row items-center gap-4 mb-6">
                <div className="flex-grow">
                    <label htmlFor="employeeId" className="block text-sm font-medium text-gray-700">Mã Nhân Viên</label>
                    <input 
                        id="employeeId" 
                        type="text" 
                        value={employeeId} 
                        onChange={(e) => setEmployeeId(e.target.value.toUpperCase())}
                        className="mt-1 w-full px-3 py-2 border rounded-md" 
                        required 
                    />
                </div>
                <div>
                    <label htmlFor="year" className="block text-sm font-medium text-gray-700">Năm</label>
                    <select id="year" value={year} onChange={(e) => setYear(parseInt(e.target.value))} className="mt-1 px-3 py-2 border border-gray-300 rounded-md shadow-sm">
                        {years.map(y => <option key={y} value={y}>{y}</option>)}
                    </select>
                </div>
                <div className="self-end mt-4 sm:mt-0">
                    <button type="submit" disabled={isLoading} className="w-full sm:w-auto px-6 py-2 font-semibold text-white bg-indigo-600 rounded-md hover:bg-indigo-700 disabled:bg-gray-400">
                        {isLoading ? 'Đang tìm...' : 'Tìm kiếm'}
                    </button>
                </div>
            </form>

            {message.text && ( <p className={`text-center p-3 rounded-md mb-4 text-sm ${ message.type === 'error' ? 'bg-red-100 text-red-700' : message.type === 'success' ? 'bg-green-100 text-green-700' : 'bg-blue-100 text-blue-700' }`}> {message.text} </p> )}
            
            {employeeInfo && (
                <div className="space-y-8">
                    <div className="bg-white p-6 rounded-lg shadow-md border border-gray-200">
                        <div className="bg-blue-50 border-l-4 border-blue-500 p-4 rounded-md mb-6">
                            <p>Đang thao tác cho: <span className="font-bold">{employeeInfo.name} ({employeeId})</span></p>
                            <p>Số giờ phép năm còn lại: <span className="font-bold text-xl text-blue-700">{employeeInfo.remaining}</span> giờ</p>
                        </div>
                        <h3 className="text-xl font-bold text-gray-800 mb-4">Nhập Phép Mới</h3>
                        <form onSubmit={handleSubmitLeave} className="space-y-4">
                            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                                <div>
                                    <label className="block text-sm font-medium">Từ ngày (ddmmyyyy)</label>
                                    <input 
                                        type="text" 
                                        name="startDate" 
                                        value={formState.startDate} 
                                        onChange={handleFormChange} 
                                        className="mt-1 w-full px-3 py-2 border rounded-md" 
                                        maxLength={8} 
                                        required 
                                        ref={startDateRef}
                                        onKeyDown={(e) => handleKeyDown(e, endDateRef)}
                                    />
                                </div>
                                <div>
                                    <label className="block text-sm font-medium">Đến ngày (ddmmyyyy)</label>
                                    <input 
                                        type="text" 
                                        name="endDate" 
                                        value={formState.endDate} 
                                        onChange={handleFormChange} 
                                        className="mt-1 w-full px-3 py-2 border rounded-md" 
                                        maxLength={8} 
                                        placeholder="Bỏ trống nếu nghỉ 1 ngày" 
                                        ref={endDateRef}
                                        onKeyDown={(e) => handleKeyDown(e, leaveTypeRef)}
                                    />
                                </div>
                            </div>
                            <div>
                               <label className="block text-sm font-medium">Loại phép</label>
                                <select 
                                    name="leaveType" 
                                    value={formState.leaveType} 
                                    onChange={handleFormChange} 
                                    className="mt-1 w-full px-3 py-2 border rounded-md bg-white"
                                    ref={leaveTypeRef}
                                    onKeyDown={(e) => handleKeyDown(e, startTimeRef)}
                                >
                                    {Object.entries(LEAVE_TYPE_OPTIONS).map(([key, value]) => (
                                        <option key={key} value={key}>{key}: {value}</option>
                                    ))}
                                </select>
                            </div>
                            <div className="grid grid-cols-2 gap-4">
                                 <div>
                                    <label className="block text-sm font-medium">Giờ bắt đầu (hhmm)</label>
                                    <input 
                                        type="text" 
                                        name="startTime" 
                                        value={formState.startTime} 
                                        onChange={handleFormChange} 
                                        maxLength={4} 
                                        required 
                                        className="mt-1 w-full px-3 py-2 border rounded-md"
                                        ref={startTimeRef}
                                        onKeyDown={(e) => handleKeyDown(e, endTimeRef)}
                                    />
                                </div>
                                 <div>
                                    <label className="block text-sm font-medium">Giờ kết thúc (hhmm)</label>
                                    <input 
                                        type="text" 
                                        name="endTime" 
                                        value={formState.endTime} 
                                        onChange={handleFormChange} 
                                        maxLength={4} 
                                        required 
                                        className="mt-1 w-full px-3 py-2 border rounded-md"
                                        ref={endTimeRef}
                                        onKeyDown={(e) => handleKeyDown(e, reasonRef)}
                                    />
                                </div>
                            </div>
                            {dailyBreakdown.length > 0 && (
                                <div className="text-sm bg-indigo-50 p-3 rounded-md">
                                    <h4 className="font-bold text-gray-700 mb-2">Chi tiết giờ nghỉ:</h4>
                                    <ul className="list-disc list-inside space-y-1">
                                        {dailyBreakdown.map((day, index) => (
                                            <li key={index} className={`text-gray-800 ${day.note ? 'text-gray-500' : ''}`}>
                                                Ngày {day.date}:{' '}
                                                {day.note ? ( <span className="font-semibold">{day.note} - 0 giờ</span> ) : ( <span className="font-semibold">{day.hours} giờ</span> )}
                                            </li>
                                        ))}
                                    </ul>
                                    <div className="border-t mt-2 pt-2">
                                        <p className="font-bold text-indigo-700">Tổng cộng giờ phép tính: {totalHours} giờ</p>
                                    </div>
                                </div>
                            )}
                            <div>
                                <label className="block text-sm font-medium">Lý do (nếu có)</label>
                                <textarea 
                                    name="reason" 
                                    value={formState.reason} 
                                    onChange={handleFormChange} 
                                    rows="2" 
                                    className="mt-1 w-full px-3 py-2 border rounded-md"
                                    ref={reasonRef}
                                    onKeyDown={(e) => {
                                        if (e.key === 'Enter' && !e.shiftKey) {
                                            e.preventDefault();
                                            submitButtonRef.current?.focus();
                                        }
                                    }}
                                />
                            </div>
                            <button 
                                type="submit" 
                                disabled={isLoading || totalHours <= 0} 
                                className="w-full px-4 py-3 font-semibold text-white bg-green-600 rounded-md hover:bg-green-700 disabled:bg-gray-400"
                                ref={submitButtonRef}
                            >
                                {isLoading ? 'Đang lưu...' : `Lưu và Nhập Phép (${totalHours} giờ)`}
                            </button>
                        </form>
                    </div>
                    {leaveEntries.length > 0 && (
                        <div className="bg-white p-6 rounded-lg shadow-md border border-gray-200">
                            <h3 className="text-xl font-bold text-gray-800 mb-4">Lịch sử nghỉ phép năm {year}</h3>
                            <div className="overflow-x-auto">
                                <table className="min-w-full bg-white border">
                                    <thead className="bg-gray-100">
                                        <tr>
                                            <th className="px-4 py-2 text-left text-sm font-semibold">Ngày nghỉ</th>
                                            <th className="px-4 py-2 text-left text-sm font-semibold">Loại phép</th>
                                            <th className="px-4 py-2 text-left text-sm font-semibold">Số giờ</th>
                                            <th className="px-4 py-2 text-left text-sm font-semibold">Hành động</th>
                                        </tr>
                                    </thead>
                                    <tbody className="divide-y">
                                        {leaveEntries.map(entry => (
                                            <tr key={entry.id}>
                                                <td className="px-4 py-2">{new Date(entry.DateUP).toLocaleDateString('vi-VN')}</td>
                                                <td className="px-4 py-2">{entry.JiaTypeName}</td>
                                                <td className="px-4 py-2">{entry.HHour}</td>
                                                <td className="px-4 py-2 space-x-2">
                                                    <button onClick={() => handleEditClick(entry)} disabled={isLoading} className="px-3 py-1 text-sm text-blue-600 bg-blue-100 rounded-md hover:bg-blue-200 disabled:opacity-50">
                                                        Sửa
                                                    </button>
                                                    <button onClick={() => handleDeleteClick(entry)} disabled={isLoading} className="px-3 py-1 text-sm text-red-600 bg-red-100 rounded-md hover:bg-red-200 disabled:opacity-50">
                                                        Xóa
                                                    </button>
                                                </td>
                                            </tr>
                                        ))}
                                    </tbody>
                                </table>
                            </div>
                        </div>
                    )}
                </div>
            )}
        </div>
    );
}
