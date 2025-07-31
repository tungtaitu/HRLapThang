/*
 * File: components/timesheet/TimesheetLog.js
 * Mô tả: Giao diện nhập công và tính toán giờ làm cho nhân viên, với logic tính toán tự động.
 * Cập nhật: Sửa lại công thức tính phụ cấp ca đêm 0.3 (B5).
 */
import React, { useState } from 'react';
import { apiGetEmployeeInfo, apiUploadTimesheet } from '../../api'; // Sử dụng apiUploadTimesheet để lưu

// --- HÀM HỖ TRỢ TÍNH TOÁN ---

// Hàm định dạng input giờ (chỉ cho phép 4 chữ số)
const formatTimeInput = (value) => {
    return value.replace(/\D/g, '').slice(0, 4);
};

// Hàm chuyển đổi chuỗi HHmm sang định dạng HH:mm
const parseToHHMM = (hhmm) => {
    if (!hhmm || hhmm.length < 3) return null;
    const padded = hhmm.padStart(4, '0');
    const hours = padded.substring(0, 2);
    const minutes = padded.substring(2, 4);
    if (parseInt(hours) > 23 || parseInt(minutes) > 59) return null;
    return `${hours}:${minutes}`;
};

// SỬA LỖI: Cập nhật hàm tính phụ cấp ca đêm 0.3 (B5)
const calculateNightAllowance_0_3 = (officialStartTimeStr, workDate, standardHours) => {
    if (!officialStartTimeStr || standardHours <= 0) return 0;
    
    const workDateStr = workDate.toISOString().split('T')[0];
    const officialStartTime = new Date(`${workDateStr}T${officialStartTimeStr}`);
    const officialWorkEnd = new Date(officialStartTime.getTime() + standardHours * 60 * 60 * 1000);

    // Khung giờ đêm được tính từ 22:00 hôm nay đến 06:00 sáng hôm sau
    const nightWindowStart = new Date(officialStartTime);
    nightWindowStart.setHours(22, 0, 0, 0);

    const nightWindowEnd = new Date(officialStartTime);
    nightWindowEnd.setDate(nightWindowEnd.getDate() + 1);
    nightWindowEnd.setHours(6, 0, 0, 0);

    // Nếu ca làm việc bắt đầu sau nửa đêm (ví dụ 01:00), thì khung giờ đêm phải là của ngày hôm trước
    if (officialStartTime.getHours() < 6) {
        nightWindowStart.setDate(nightWindowStart.getDate() - 1);
        nightWindowEnd.setDate(nightWindowEnd.getDate() - 1);
    }

    const overlapStart = Math.max(officialStartTime.getTime(), nightWindowStart.getTime());
    const overlapEnd = Math.min(officialWorkEnd.getTime(), nightWindowEnd.getTime());
    
    const diffMs = overlapEnd - overlapStart;

    if (diffMs > 0) {
        return parseFloat((diffMs / (1000 * 60 * 60)).toFixed(2));
    }
    return 0;
};


export default function TimesheetLog() {
    const [formData, setFormData] = useState({
        empid: '',
        workDate: new Date().toISOString().split('T')[0],
        breakHours: 1.0, // Giờ nghỉ
        timeup: '0800', // Giờ vào thực tế
        timedown: '1700', // Giờ ra thực tế
        shiftStartTime: '0800' // Giờ bắt đầu ca chính thức
    });
    const [employeeInfo, setEmployeeInfo] = useState(null);
    const [calculations, setCalculations] = useState(null);
    const [isLoading, setIsLoading] = useState(false);
    const [message, setMessage] = useState({ type: '', text: '' });

    const SPECIAL_ALLOWANCE_GROUPS = ['A059', 'A073', 'A067'];

    const handleChange = (e) => {
        let { name, value } = e.target;
        if (name === 'timeup' || name === 'timedown' || name === 'shiftStartTime') {
            value = formatTimeInput(value);
        }
        setFormData(prev => ({ ...prev, [name]: value }));
        setCalculations(null);
    };

    const handleCheckEmployee = async () => {
        if (!formData.empid) return;
        try {
            const info = await apiGetEmployeeInfo(formData.empid);
            setEmployeeInfo(info);
            setMessage({ type: '', text: '' });
        } catch (error) {
            setEmployeeInfo(null);
            setMessage({ type: 'error', text: 'Không tìm thấy nhân viên.' });
        }
    };

    const handleCalculate = () => {
        const { timeup, timedown, workDate, breakHours, shiftStartTime } = formData;
        const standardHours = 8.0; // Giờ công chuẩn mặc định
        const timeupHHMM = parseToHHMM(timeup);
        const timedownHHMM = parseToHHMM(timedown);
        const shiftStartHHMM = parseToHHMM(shiftStartTime);

        if (!timeupHHMM || !timedownHHMM || !shiftStartHHMM) {
            setMessage({type: 'error', text: 'Vui lòng nhập đủ thông tin và đúng định dạng giờ (HHmm).'});
            return;
        }

        const start = new Date(`${workDate}T${timeupHHMM}`);
        const end = new Date(`${workDate}T${timedownHHMM}`);
        const isNightShift = timeup > timedown;
        if (isNightShift) {
            end.setDate(end.getDate() + 1);
        }

        const breakMinutes = (parseFloat(breakHours) || 0) * 60;
        const totalMinutes = (end - start) / (1000 * 60) - breakMinutes;
        const TOTH = parseFloat((totalMinutes / 60).toFixed(2));
        const officialHours = Math.min(TOTH, standardHours);

        let H1 = 0, H2 = 0, H3 = 0, B3 = 0;
        const overtimeHours = TOTH > standardHours ? TOTH - standardHours : 0;
        const dateObj = new Date(workDate);
        const dayOfWeek = dateObj.getDay(); // 0 = Sunday

        if (overtimeHours > 0) {
            if (isNightShift) {
                B3 = overtimeHours;
            } else if (dayOfWeek === 0) {
                H2 = overtimeHours;
            } else {
                H1 = overtimeHours;
            }
        }

        let B4 = 0;
        if (employeeInfo && SPECIAL_ALLOWANCE_GROUPS.includes(employeeInfo.GROUPID) && officialHours >= 8) {
            B4 = 0.5;
        }

        const B5 = calculateNightAllowance_0_3(shiftStartHHMM, dateObj, officialHours);

        setCalculations({ TOTH, officialHours, H1, H2, H3, B3, B4, B5 });
        setMessage({type: '', text: ''});
    };

    const handleSave = async () => {
        if (!calculations) {
            setMessage({type: 'error', text: 'Vui lòng nhấn "Tính toán" trước khi lưu.'});
            return;
        }
        setIsLoading(true);
        setMessage({ type: '', text: '' });
        
        const workdat_yyyymmdd = formData.workDate.replace(/-/g, '');
        const { officialHours, ...restOfCalculations } = calculations;

        const payload = {
            ...restOfCalculations,
            Kzhour: 0, // Kzhour (giờ vắng) luôn là 0 khi nhập thủ công
            empid: formData.empid,
            workdat: workdat_yyyymmdd,
            timeup: formData.timeup.padEnd(6, '0'),
            timedown: formData.timedown.padEnd(6, '0'),
            yymm: workdat_yyyymmdd.substring(0, 6),
            muser: 'MANUAL',
            flag: 'Manual',
            BC: null, FORGET: 0, Latefor: 0,
            EMPWHSNO: employeeInfo?.WHSNO || 'LT',
            indat: employeeInfo?.INDAT || null,
            outdat: null
        };

        try {
            const result = await apiUploadTimesheet([payload]); 
            setMessage({ type: 'success', text: result.message });
        } catch (error) {
            setMessage({ type: 'error', text: error.message });
        } finally {
            setIsLoading(false);
        }
    };

    return (
        <div>
            <h2 className="text-2xl font-bold text-gray-800 mb-4">Nhập & Tính Công</h2>
            <div className="space-y-4 max-w-2xl mx-auto">
                <div className="bg-gray-50 p-4 rounded-lg space-y-4">
                    <div>
                        <label className="block text-sm font-medium">Mã số nhân viên (*)</label>
                        <div className="mt-1">
                            <input type="text" name="empid" value={formData.empid} onChange={handleChange} onBlur={handleCheckEmployee} className="w-full px-3 py-2 border rounded-md" required />
                        </div>
                        {employeeInfo && <p className="text-sm text-green-600 mt-1">Tên nhân viên: {employeeInfo.EMPNAM_VN}</p>}
                    </div>
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                        <div>
                            <label className="block text-sm font-medium">Ngày làm việc (*)</label>
                            <input type="date" name="workDate" value={formData.workDate} onChange={handleChange} className="mt-1 w-full px-3 py-2 border rounded-md" required />
                        </div>
                         <div>
                            <label className="block text-sm font-medium">Giờ nghỉ (giờ) (*)</label>
                            <input type="number" step="0.5" name="breakHours" value={formData.breakHours} onChange={handleChange} className="mt-1 w-full px-3 py-2 border rounded-md" required />
                        </div>
                        <div>
                            <label className="block text-sm font-medium">Giờ bắt đầu ca (HHmm) (*)</label>
                            <input type="text" name="shiftStartTime" value={formData.shiftStartTime} onChange={handleChange} className="mt-1 w-full px-3 py-2 border rounded-md" required placeholder="Ví dụ: 0800"/>
                        </div>
                        <div></div> {/* Placeholder for grid alignment */}
                        <div>
                            <label className="block text-sm font-medium">Giờ vào thực tế (HHmm) (*)</label>
                            <input type="text" name="timeup" value={formData.timeup} onChange={handleChange} className="mt-1 w-full px-3 py-2 border rounded-md" required placeholder="Ví dụ: 1930"/>
                        </div>
                        <div>
                            <label className="block text-sm font-medium">Giờ ra thực tế (HHmm) (*)</label>
                            <input type="text" name="timedown" value={formData.timedown} onChange={handleChange} className="mt-1 w-full px-3 py-2 border rounded-md" required placeholder="Ví dụ: 0700"/>
                        </div>
                    </div>
                </div>

                <div className="flex justify-center gap-4">
                    <button onClick={handleCalculate} disabled={isLoading || !employeeInfo} className="px-6 py-2 font-semibold text-white bg-blue-600 rounded-md hover:bg-blue-700 disabled:bg-gray-400">
                        {isLoading ? '...' : 'Tính toán'}
                    </button>
                    <button onClick={handleSave} disabled={isLoading || !calculations} className="px-6 py-2 font-semibold text-white bg-green-600 rounded-md hover:bg-green-700 disabled:bg-gray-400">
                        {isLoading ? '...' : 'Lưu Công'}
                    </button>
                </div>

                {message.text && ( <p className={`text-center p-3 rounded-md mt-4 text-sm ${ message.type === 'error' ? 'bg-red-100 text-red-700' : message.type === 'success' ? 'bg-green-100 text-green-700' : 'bg-blue-100 text-blue-700' }`}> {message.text} </p> )}

                {calculations && (
                    <div className="mt-6 bg-white p-4 rounded-lg shadow-md border">
                        <h3 className="text-lg font-bold mb-2">Kết quả tính toán:</h3>
                        <div className="grid grid-cols-2 gap-x-4 gap-y-2 text-sm">
                            <p>Tổng giờ làm (đã trừ nghỉ):</p><p className="font-semibold">{calculations.TOTH.toFixed(2)} giờ</p>
                            <p>Giờ chính thức:</p><p className="font-semibold">{calculations.officialHours.toFixed(2)} giờ</p>
                            <p>Tăng ca ngày thường (H1):</p><p className="font-semibold text-green-600">{calculations.H1.toFixed(2)} giờ</p>
                            <p>Tăng ca Chủ Nhật (H2):</p><p className="font-semibold text-green-600">{calculations.H2.toFixed(2)} giờ</p>
                            <p>Tăng ca đêm (B3):</p><p className="font-semibold text-indigo-600">{calculations.B3.toFixed(2)} giờ</p>
                            <p>Phụ cấp 0.5 (B4):</p><p className="font-semibold text-yellow-600">{calculations.B4.toFixed(2)} giờ</p>
                            <p>Phụ cấp ca đêm 0.3 (B5):</p><p className="font-semibold text-purple-600">{calculations.B5.toFixed(2)} giờ</p>
                        </div>
                    </div>
                )}
            </div>
        </div>
    );
}
