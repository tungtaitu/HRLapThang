/*
 * File: components/admin/ManualLeaveEntry.js
 * Mô tả: Chức năng cho phép admin nhập phép thủ công cho nhân viên.
 */
import React, { useState, useEffect, useRef, useCallback } from 'react';
import { apiFetchHolidays, apiAdminSubmitLeave } from '../../api';

export default function AdminManualLeaveEntry() {
    const [employeeId, setEmployeeId] = useState('');
    const [employeeInfo, setEmployeeInfo] = useState(null);
    const [isLoading, setIsLoading] = useState(false);
    const [message, setMessage] = useState({ type: '', text: '' });

    const [startDate, setStartDate] = useState('');
    const [endDate, setEndDate] = useState('');
    const [leaveType, setLeaveType] = useState('E');
    const [startTime, setStartTime] = useState('');
    const [endTime, setEndTime] = useState('');
    const [reason, setReason] = useState('');
    const [dailyBreakdown, setDailyBreakdown] = useState([]);

    const allRefs = {
        employeeId: useRef(null), startDate: useRef(null), endDate: useRef(null),
        leaveType: useRef(null), startTime: useRef(null), endTime: useRef(null),
        reason: useRef(null), submit: useRef(null)
    };

    const PUBLIC_HOLIDAYS = {
        '2025-01-01': 'Tết Dương lịch', '2025-01-28': 'Tết Nguyên Đán', '2025-01-29': 'Tết Nguyên Đán',
        '2025-01-30': 'Tết Nguyên Đán', '2025-01-31': 'Tết Nguyên Đán', '2025-02-03': 'Nghỉ bù Tết Nguyên Đán',
        '2025-04-08': 'Giỗ Tổ Hùng Vương', '2025-04-30': 'Ngày Giải phóng miền Nam', '2025-05-01': 'Ngày Quốc tế Lao động',
        '2025-05-02': 'Nghỉ bù', '2025-09-01': 'Nghỉ Quốc Khánh', '2025-09-02': 'Quốc Khánh',
    };

    const handleKeyDown = (e, nextRefKey) => {
        if (e.key === 'Enter') {
            e.preventDefault();
            if (nextRefKey && allRefs[nextRefKey] && allRefs[nextRefKey].current) {
                allRefs[nextRefKey].current.focus();
            }
        }
    };

    const calculateHours = useCallback((startStr, endStr) => {
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
    }, []);

    const parseDateForFrontend = (ddmmyyyy) => {
        if (!ddmmyyyy || ddmmyyyy.length !== 8) return null;
        const day = parseInt(ddmmyyyy.substring(0, 2));
        const month = parseInt(ddmmyyyy.substring(2, 4)) - 1;
        const year = parseInt(ddmmyyyy.substring(4, 8));
        const date = new Date(year, month, day);
        if (date.getFullYear() !== year || date.getMonth() !== month || date.getDate() !== day) return null;
        return date;
    };

    useEffect(() => {
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
    }, [startDate, endDate, startTime, endTime, calculateHours]);

    useEffect(() => {
        if (employeeInfo && allRefs.startDate.current) {
            allRefs.startDate.current.focus();
        }
    }, [employeeInfo, allRefs.startDate]);

    const handleCheckEmployee = useCallback(async () => {
        if (!employeeId) return;
        setIsLoading(true); setEmployeeInfo(null); setMessage({ type: '', text: '' });
        try {
            const data = await apiFetchHolidays(employeeId, new Date().getFullYear());
            setEmployeeInfo({
                id: employeeId,
                name: data.employeeName,
                remaining: data.summary.remaining,
                isForeigner: data.isForeigner
            });
        } catch (error) {
            setMessage({ type: 'error', text: `Không tìm thấy nhân viên: ${error.message}` });
        } finally {
            setIsLoading(false);
        }
    }, [employeeId]);

    const resetFormFields = useCallback(() => {
        setStartDate(''); setEndDate(''); setStartTime(''); setEndTime('');
        setReason(''); setDailyBreakdown([]);
        setMessage({ type: 'success', text: 'Gửi thành công! Sẵn sàng nhập lượt tiếp theo.' });
        setTimeout(() => {
            setMessage({ type: '', text: '' });
            if (allRefs.startDate.current) allRefs.startDate.current.focus();
        }, 2000);
    }, [allRefs.startDate]);

    const handleSubmitLeave = useCallback(async (e) => {
        e.preventDefault();
        if (!employeeInfo) return;
        setIsLoading(true); setMessage({ type: '', text: '' });
        try {
            const leaveData = {
                userId: employeeInfo.id,
                startDate, endDate: endDate || startDate,
                leaveType, startTime, endTime, reason
            };
            await apiAdminSubmitLeave(leaveData);
            const updatedData = await apiFetchHolidays(employeeId, new Date().getFullYear());
            setEmployeeInfo({ ...employeeInfo, remaining: updatedData.summary.remaining, name: updatedData.employeeName, isForeigner: updatedData.isForeigner });
            resetFormFields();
        } catch (error) {
             setMessage({ type: 'error', text: `Lỗi khi gửi đơn: ${error.message}` });
        } finally {
            setIsLoading(false);
        }
    }, [employeeInfo, employeeId, startDate, endDate, leaveType, startTime, endTime, reason, resetFormFields]);

    const totalHours = dailyBreakdown.reduce((acc, day) => acc + day.hours, 0);

    return (
        <div>
            <h2 className="text-2xl font-bold text-gray-800 mb-4">Nhập Phép Nhân Viên</h2>
            <div className="space-y-4 max-w-2xl mx-auto">
                <div className="bg-gray-50 p-4 rounded-lg">
                    <label htmlFor="employeeIdInput" className="block text-sm font-medium text-gray-700">Mã số nhân viên (MSNV)</label>
                    <div className="mt-1 flex gap-2">
                        <input type="text" id="employeeIdInput" ref={allRefs.employeeId} value={employeeId}
                               onChange={(e) => setEmployeeId(e.target.value.toUpperCase())}
                               onKeyDown={(e) => { if (e.key === 'Enter') { e.preventDefault(); handleCheckEmployee(); } }}
                               className="w-full px-3 py-2 border rounded-md" placeholder="Nhập MSNV rồi nhấn Enter..."/>
                        <button onClick={handleCheckEmployee} disabled={isLoading}
                                className="px-4 py-2 font-semibold text-white bg-indigo-600 rounded-md hover:bg-indigo-700 disabled:bg-gray-400">
                            {isLoading && !employeeInfo ? 'Đang...' : 'Kiểm tra'}
                        </button>
                    </div>
                </div>

                {employeeInfo && (
                    <form onSubmit={handleSubmitLeave} className="bg-white p-6 rounded-lg shadow-md space-y-4">
                        <div className="bg-blue-50 border-l-4 border-blue-500 p-3 rounded-md">
                           <div className="flex justify-between items-start">
                                <div>
                                    <p>Đang nhập phép cho MSNV: <span className="font-bold">{employeeInfo.id} - {employeeInfo.name}</span></p>
                                    <p>Số giờ phép năm còn lại: <span className="font-bold text-xl">{employeeInfo.remaining}</span> giờ</p>
                                </div>
                                <button type="button" onClick={() => { setEmployeeInfo(null); setEmployeeId(''); allRefs.employeeId.current.focus();}}
                                    className="text-sm text-red-500 hover:text-red-700 flex-shrink-0">
                                    Đổi NV
                                </button>
                           </div>
                            <p className={`mt-2 text-sm font-semibold ${employeeInfo.isForeigner ? 'text-teal-700' : 'text-gray-600'}`}>
                                Đối tượng: {employeeInfo.isForeigner ? 'Lao động nước ngoài (16h/tháng)' : 'Lao động trong nước (8h/tháng)'}
                            </p>
                        </div>

                        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                            <div>
                                <label htmlFor="startDate" className="block text-sm font-medium">Từ ngày (ddmmyyyy)</label>
                                <input type="text" id="startDate" ref={allRefs.startDate} value={startDate}
                                    onChange={e => setStartDate(e.target.value)}
                                    onKeyDown={(e) => handleKeyDown(e, 'endDate')}
                                    className="mt-1 w-full px-3 py-2 border rounded-md" maxLength={8}
                                    placeholder="Ví dụ: 22072025" required />
                            </div>
                            <div>
                                <label htmlFor="endDate" className="block text-sm font-medium">Đến ngày (ddmmyyyy)</label>
                                <input type="text" id="endDate" ref={allRefs.endDate} value={endDate}
                                    onChange={e => setEndDate(e.target.value)}
                                    onKeyDown={(e) => handleKeyDown(e, 'leaveType')}
                                    className="mt-1 w-full px-3 py-2 border rounded-md" maxLength={8}
                                    placeholder="Bỏ trống nếu nghỉ 1 ngày" />
                            </div>
                        </div>
                         <div>
                           <label htmlFor="leaveType" className="block text-sm font-medium">Loại phép</label>
                            <select id="leaveType" ref={allRefs.leaveType} value={leaveType}
                                    onChange={e => setLeaveType(e.target.value)}
                                    onKeyDown={(e) => handleKeyDown(e, 'startTime')}
                                    className="mt-1 w-full px-3 py-2 border rounded-md bg-white">
                                <option value="E">E: P.Năm </option>
                                <option value="A">A: P.Việc riêng</option>
                                <option value="B">B: P.Bệnh</option>
                                <option value="C">C: Nghỉ kết hôn</option>
                                <option value="D">D: P.Tang </option>
                                <option value="F">F: Nghỉ thai sản </option>
                                <option value="G">G: Nghỉ C.Tác </option>
                                <option value="H">H: Nghỉ C.Thường</option>
                                <option value="I">I: Đi đường</option>
                                <option value="K">K: Không lương</option>
                            </select>
                        </div>
                        <div className="grid grid-cols-2 gap-4">
                             <div>
                                <label htmlFor="startTime" className="block text-sm font-medium">Giờ bắt đầu (hhmm)</label>
                                <input type="text" id="startTime" ref={allRefs.startTime} value={startTime}
                                       onChange={e => setStartTime(e.target.value)}
                                       onKeyDown={(e) => handleKeyDown(e, 'endTime')}
                                       maxLength={4} placeholder="Ví dụ: 0800"
                                       required className="mt-1 w-full px-3 py-2 border rounded-md"/>
                            </div>
                             <div>
                                <label htmlFor="endTime" className="block text-sm font-medium">Giờ kết thúc (hhmm)</label>
                                <input type="text" id="endTime" ref={allRefs.endTime} value={endTime}
                                       onChange={e => setEndTime(e.target.value)}
                                       onKeyDown={(e) => handleKeyDown(e, 'reason')}
                                       maxLength={4} placeholder="Ví dụ: 1700"
                                       required className="mt-1 w-full px-3 py-2 border rounded-md"/>
                            </div>
                        </div>
                        {dailyBreakdown.length > 0 && (
                            <div className="text-sm bg-indigo-50 p-3 rounded-md">
                                <h4 className="font-bold text-gray-700 mb-2">Chi tiết giờ nghỉ:</h4>
                                <ul className="list-disc list-inside space-y-1">
                                    {dailyBreakdown.map((day, index) => (
                                        <li key={index} className={`text-gray-800 ${day.note ? 'text-gray-500' : ''}`}>
                                            Ngày {day.date}:{' '}
                                            {day.note ? (
                                                <span className="font-semibold">{day.note} - 0 giờ</span>
                                            ) : (
                                                <span className="font-semibold">{day.hours} giờ</span>
                                            )}
                                        </li>
                                    ))}
                                </ul>
                                <div className="border-t mt-2 pt-2">
                                    <p className="font-bold text-indigo-700">Tổng cộng giờ phép tính: {totalHours} giờ</p>
                                </div>
                            </div>
                        )}

                        <div>
                            <label htmlFor="reason" className="block text-sm font-medium">Lý do (nếu có)</label>
                            <textarea id="reason" ref={allRefs.reason} value={reason}
                                      onChange={e => setReason(e.target.value)}
                                      onKeyDown={(e) => handleKeyDown(e, 'submit')}
                                      rows="2" className="mt-1 w-full px-3 py-2 border rounded-md"/>
                        </div>
                         <button type="submit" ref={allRefs.submit} disabled={isLoading || totalHours <= 0}
                                className="w-full px-4 py-3 font-semibold text-white bg-green-600 rounded-md hover:bg-green-700 disabled:bg-gray-400 focus:ring-2 focus:ring-green-500">
                            {isLoading ? 'Đang lưu...' : `Lưu và Nhập Phép (${totalHours} giờ)`}
                        </button>
                    </form>
                )}

                {message.text && (
                   <p className={`mt-4 text-center text-sm font-semibold ${message.type === 'error' ? 'text-red-600' : 'text-green-600'}`}>
                        {message.text}
                    </p>
                )}
            </div>
        </div>
    );
}
