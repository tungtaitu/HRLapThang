/*
 * File: components/timesheet/UploadTimesheetComponent.js
 * Mô tả: Chứa toàn bộ logic cho việc upload, xử lý, và chỉnh sửa dữ liệu chấm công từ file Excel.
 * Cập nhật: Loại bỏ tính năng làm mới, sửa lỗi và bỏ cột "Ngày nghỉ việc".
 */
import React, { useState, useEffect } from 'react';
import * as XLSX from 'xlsx';

// --- Import hàm từ file api/index.js của bạn ---
import { apiGetAllEmployeesDetails, apiUploadTimesheet } from '../../api';

// --- HÀM HỖ TRỢ TÍNH TOÁN ---
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

// --- Component Bảng Chỉnh Sửa Dữ Liệu ---
const EditableTimesheetTable = ({ data, onSave, onCancel, isSaving, saveSuccessMessage, onReset }) => {
    const [editedData, setEditedData] = useState(data);

    const handleInputChange = (index, field, value) => {
        const newData = [...editedData];
        newData[index][field] = value;
        setEditedData(newData);
    };

    const handleSave = () => {
        const dataToSave = editedData.map(row => {
            const newRow = {...row};
            delete newRow.empName;
            delete newRow.outdat; 
            for (const key in newRow) {
                if (['TOTH', 'Kzhour', 'H1', 'H2', 'H3', 'B3', 'B4', 'B5', 'FORGET', 'Latefor'].includes(key)) {
                    newRow[key] = parseFloat(newRow[key]) || 0;
                }
            }
            return newRow;
        });
        onSave(dataToSave);
    };

    return (
        <div className="mt-6">
            <h2 className="text-2xl font-bold text-gray-800 mb-4">Xem lại và Chỉnh sửa Dữ liệu Chấm công</h2>
            <div className="overflow-auto bg-white rounded-lg shadow" style={{ maxHeight: '70vh' }}>
                <table className="min-w-full text-sm">
                    <thead>
                        <tr>
                            {/* SỬA ĐỔI: Bỏ cột "Ngày nghỉ việc" */}
                            {['Mã NV', 'Tên NV', 'Ngày', 'Giờ vào', 'Giờ ra', 'Giờ vắng (Kz)', 'Tổng giờ', 'TC 1.5 (H1)', 'TC 2.0 (H2)', 'TC 3.0 (H3)', 'TC Đêm (B3)', 'PC 0.5 (B4)', 'PC 0.3 (B5)'].map(header => (
                                <th key={header} className="sticky top-0 z-10 p-2 text-left font-semibold text-gray-600 bg-gray-100 whitespace-nowrap border-b">
                                    {header}
                                </th>
                            ))}
                        </tr>
                    </thead>
                    <tbody className="divide-y divide-gray-200">
                        {editedData.map((row, index) => (
                            <tr key={row.empid + row.workdat} className="hover:bg-gray-50">
                                <td className="p-2 whitespace-nowrap font-semibold">{row.empid}</td>
                                <td className="p-2 whitespace-nowrap">{row.empName}</td>
                                <td className="p-2 whitespace-nowrap">{row.workdat}</td>
                                <td className="p-1"><input type="text" value={row.timeup || ''} onChange={e => handleInputChange(index, 'timeup', e.target.value)} className="w-20 p-1 border rounded-md bg-white"/></td>
                                <td className="p-1"><input type="text" value={row.timedown || ''} onChange={e => handleInputChange(index, 'timedown', e.target.value)} className="w-20 p-1 border rounded-md bg-white"/></td>
                                <td className="p-1"><input type="number" step="0.1" value={row.Kzhour} onChange={e => handleInputChange(index, 'Kzhour', e.target.value)} className="w-16 p-1 border rounded-md bg-red-50"/></td>
                                <td className="p-1"><input type="number" step="0.1" value={row.TOTH} onChange={e => handleInputChange(index, 'TOTH', e.target.value)} className="w-16 p-1 border rounded-md bg-gray-50"/></td>
                                <td className="p-1"><input type="number" step="0.1" value={row.H1} onChange={e => handleInputChange(index, 'H1', e.target.value)} className="w-16 p-1 border rounded-md bg-green-50"/></td>
                                <td className="p-1"><input type="number" step="0.1" value={row.H2} onChange={e => handleInputChange(index, 'H2', e.target.value)} className="w-16 p-1 border rounded-md bg-green-50"/></td>
                                <td className="p-1"><input type="number" step="0.1" value={row.H3} onChange={e => handleInputChange(index, 'H3', e.target.value)} className="w-16 p-1 border rounded-md bg-green-50"/></td>
                                <td className="p-1"><input type="number" step="0.1" value={row.B3} onChange={e => handleInputChange(index, 'B3', e.target.value)} className="w-16 p-1 border rounded-md bg-purple-50"/></td>
                                <td className="p-1"><input type="number" step="0.1" value={row.B4} onChange={e => handleInputChange(index, 'B4', e.target.value)} className="w-16 p-1 border rounded-md bg-yellow-50"/></td>
                                <td className="p-1"><input type="number" step="0.1" value={row.B5} onChange={e => handleInputChange(index, 'B5', e.target.value)} className="w-16 p-1 border rounded-md bg-blue-50"/></td>
                            </tr>
                        ))}
                    </tbody>
                </table>
            </div>
            <div className="mt-6 flex justify-end gap-4">
                {saveSuccessMessage ? (
                    <div className="w-full flex flex-col sm:flex-row items-center justify-between gap-4 p-4 bg-green-100 text-green-800 rounded-md">
                        <p className="font-semibold">{saveSuccessMessage}</p>
                        <button onClick={onReset} className="px-6 py-2 font-semibold text-white bg-green-600 rounded-md hover:bg-green-700 whitespace-nowrap">
                           Upload File Mới
                        </button>
                    </div>
                ) : (
                    <>
                        <button onClick={onCancel} disabled={isSaving} className="px-6 py-2 font-semibold text-gray-700 bg-gray-200 rounded-md hover:bg-gray-300 disabled:opacity-50">
                            Hủy bỏ & Upload Lại
                        </button>
                        <button onClick={handleSave} disabled={isSaving} className="px-6 py-2 font-semibold text-white bg-blue-600 rounded-md hover:bg-blue-700 disabled:bg-blue-400">
                            {isSaving ? 'Đang lưu...' : 'Lưu vào CSDL'}
                        </button>
                    </>
                )}
            </div>
        </div>
    );
};

// --- Component Upload ---
function TimesheetUpload({ onProcessComplete, employeeDetails }) {
    const [file, setFile] = useState(null);
    const [isProcessing, setIsProcessing] = useState(false);
    const [error, setError] = useState('');
    
    const SPECIAL_ALLOWANCE_GROUPS = ['A059', 'A073', 'A067'];

    const handleFileChange = (e) => setFile(e.target.files[0]);

    const convertExcelDate = (serial) => {
        if (typeof serial !== 'number' || isNaN(serial)) return null;
        const utc_days = Math.floor(serial - 25569);
        const date = new Date(utc_days * 86400 * 1000);
        const year = date.getUTCFullYear();
        const month = date.getUTCMonth();
        const day = date.getUTCDate();
        return new Date(Date.UTC(year, month, day));
    };
    
    const handleProcessFile = () => {
        if (!file || !employeeDetails) {
            setError('Vui lòng chọn tệp và chờ dữ liệu nhân viên được tải xong.');
            return;
        }
        setIsProcessing(true);
        setError('');

        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const workbook = XLSX.read(e.target.result, { type: 'binary' });
                const worksheet = workbook.Sheets[workbook.SheetNames[0]];
                const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, range: 2 });
                const headers = jsonData[0];
                const dataRows = jsonData.slice(1);
                const requiredColumns = ['Mã N.Viên', 'Tên nhân viên', 'Ngày', 'Vào 1', 'Ra 1', 'Giờ', 'TC1', 'TC2', 'TC3', 'Tổng giờ', 'Vào Trễ', 'Ra sớm'];
                
                const missingColumns = requiredColumns.filter(col => !headers.includes(col));
                if (missingColumns.length > 0) throw new Error(`Tệp thiếu cột: ${missingColumns.join(', ')}`);

                const dataWithHeaders = dataRows.map(row => headers.reduce((obj, header, index) => ({ ...obj, [header]: row[index] }), {}));
                
                const formattedDataForDB = dataWithHeaders.map(row => {
                    const empid = row['Mã N.Viên'];
                    if (!empid) return null;
                    
                    const workDate = convertExcelDate(parseFloat(row['Ngày']));
                    if (!workDate) return null;
                    
                    const employeeInfo = employeeDetails[empid];
                    
                    if (employeeInfo && employeeInfo.OUTDAT) {
                        const outDate = new Date(employeeInfo.OUTDAT);
                        const firstDayOfWorkMonth = new Date(Date.UTC(workDate.getUTCFullYear(), workDate.getUTCMonth(), 1));
                        if (outDate < firstDayOfWorkMonth) {
                            return null; 
                        }
                    }
                    
                    const workdat_yyyymmdd = `${workDate.getUTCFullYear()}${(workDate.getUTCMonth() + 1).toString().padStart(2, '0')}${workDate.getUTCDate().toString().padStart(2, '0')}`;
                    const timeup = row['Vào 1'] || '';
                    const timedown = row['Ra 1'] || '';

                    const isAbsentBySymbol = (row['Kí hiệu'] === 'V' || row['Kí hiệu 2'] === 'V');
                    const isAbsentByTime = !timeup && !timedown;
                    const isSunday = workDate.getUTCDay() === 0;

                    // Case 1: Absent (by symbol 'V' or by empty time on a workday)
                    if (isAbsentBySymbol || (isAbsentByTime && !isSunday)) {
                        return {
                            empid, empName: row['Tên nhân viên'], workdat: workdat_yyyymmdd,
                            timeup: '000000', timedown: '000000', TOTH: 0, Kzhour: 8,
                            H1: 0, H2: 0, H3: 0, B3: 0, B4: 0, B5: 0, BC: isAbsentBySymbol ? 'V' : null,
                            yymm: workdat_yyyymmdd.substring(0, 6), FORGET: 0, Latefor: 0,
                            EMPWHSNO: employeeInfo?.WHSNO || 'LT', indat: employeeInfo?.INDAT || null,
                            outdat: employeeInfo?.OUTDAT || null, flag: 'Auto', memo: 'Vắng', muser: 'Step2',
                        };
                    } 
                    // Case 2: Present or Sunday off
                    else {
                        const normalHoursFromFile = parseFloat(row['Giờ'] || 0);
                        let absentHours = 0;
                        if (!isSunday && normalHoursFromFile < 8) {
                            absentHours = 8 - normalHoursFromFile;
                        }

                        let phuCap_0_5 = 0;
                        if (employeeInfo && SPECIAL_ALLOWANCE_GROUPS.includes(employeeInfo.GROUPID) && normalHoursFromFile >= 8) {
                            phuCap_0_5 = 0.5;
                        }
                        
                        const isNightShift = timeup > timedown;
                        
                        let H1 = 0, H2 = parseFloat(row['TC2'] || 0), B3 = 0;
                        const tc1Value = parseFloat(row['TC1'] || 0);
                        if (isNightShift) { B3 = tc1Value; } 
                        else if (isSunday) { H2 += tc1Value; } 
                        else { H1 = tc1Value; }

                        const lateArrival = parseFloat(row['Vào Trễ'] || 0);
                        
                        return {
                            empid, empName: row['Tên nhân viên'], workdat: workdat_yyyymmdd,
                            timeup: timeup ? timeup.replace(':', '') + '00' : '000000',
                            timedown: timedown ? timedown.replace(':', '') + '00' : '000000',
                            TOTH: parseFloat(row['Tổng giờ'] || 0), Kzhour: absentHours,
                            H1, H2, H3: parseFloat(row['TC3'] || 0), B3, B4: phuCap_0_5,
                            B5: calculateNightAllowance_0_3(timeup, workDate, isNightShift, normalHoursFromFile),
                            BC: row['Tên ca'] || null, yymm: workdat_yyyymmdd.substring(0, 6),
                            FORGET: (timeup && !timedown) || (!timeup && timedown) ? 1 : 0,
                            Latefor: lateArrival,
                            EMPWHSNO: employeeInfo?.WHSNO || 'LT', indat: employeeInfo?.INDAT || null,
                            outdat: employeeInfo?.OUTDAT || null, flag: 'Auto', memo: null, muser: 'Step2',
                        };
                    }
                }).filter(Boolean);

                if (formattedDataForDB.length === 0) throw new Error("Không tìm thấy dữ liệu hợp lệ hoặc tất cả nhân viên trong file đã nghỉ việc.");
                
                onProcessComplete(formattedDataForDB);

            } catch (err) {
                setError(`Lỗi xử lý tệp: ${err.message}`);
            } finally {
                setIsProcessing(false);
            }
        };
        reader.readAsBinaryString(file);
    };

    return (
        <div className="max-w-2xl mx-auto">
            <h2 className="text-2xl font-bold text-gray-800 mb-4">Upload File Chấm Công (Excel/CSV)</h2>
            <div className="bg-gray-50 p-6 rounded-lg border">
                <div className="mb-4">
                    <label className="block text-sm font-medium text-gray-700 mb-2">Chọn tệp Excel:</label>
                    <input type="file" accept=".xls, .xlsx" onChange={handleFileChange} className="block w-full text-sm text-gray-500 file:mr-4 file:py-2 file:px-4 file:rounded-full file:border-0 file:text-sm file:font-semibold file:bg-blue-50 file:text-blue-700 hover:file:bg-blue-100" />
                </div>
                <button onClick={handleProcessFile} disabled={isProcessing || !file || !employeeDetails} className="w-full mt-4 px-6 py-3 font-semibold text-white bg-indigo-600 rounded-md hover:bg-indigo-700 disabled:bg-gray-400">
                    {isProcessing ? 'Đang xử lý...' : (employeeDetails ? 'Xem trước & Chỉnh sửa' : 'Đang tải dữ liệu nhân viên...')}
                </button>
            </div>
            {error && <div className="mt-4 p-3 rounded-md bg-red-100 text-red-700 text-sm"><strong>Lỗi:</strong> {error}</div>}
        </div>
    );
}


// --- Component Cha cho chức năng Upload ---
export default function UploadTimesheetComponent() {
    const [processedData, setProcessedData] = useState(null);
    const [isSaving, setIsSaving] = useState(false);
    const [saveSuccessMessage, setSaveSuccessMessage] = useState('');
    const [saveErrorMessage, setSaveErrorMessage] = useState('');
    
    const [employeeDetails, setEmployeeDetails] = useState(null);
    const [fetchError, setFetchError] = useState('');

    useEffect(() => {
        const fetchDetails = async () => {
            try {
                const details = await apiGetAllEmployeesDetails();
                const detailsMap = details.reduce((acc, emp) => ({...acc, [emp.EMPID]: emp}), {});
                setEmployeeDetails(detailsMap);
            } catch (err) {
                setFetchError(err.message);
            }
        };
        fetchDetails();
    }, []);


    const handleSaveToDB = async (data) => {
        setIsSaving(true);
        setSaveSuccessMessage('');
        setSaveErrorMessage('');
        try {
            const result = await apiUploadTimesheet(data);
            setSaveSuccessMessage(result.message || 'Lưu thành công!');
        } catch (error) {
            setSaveErrorMessage(`Lỗi khi lưu: ${error.message}`);
        } finally {
            setIsSaving(false);
        }
    };
    
    const handleResetComponent = () => {
        setProcessedData(null);
        setSaveSuccessMessage('');
        setSaveErrorMessage('');
        setIsSaving(false);
    };

    if (processedData) {
        return (
            <div>
                {saveErrorMessage && <div className="my-4 p-3 rounded-md bg-red-100 text-red-700 text-sm"><strong>Lỗi:</strong> {saveErrorMessage}</div>}
                <EditableTimesheetTable 
                    data={processedData} 
                    onSave={handleSaveToDB} 
                    onCancel={handleResetComponent}
                    isSaving={isSaving} 
                    saveSuccessMessage={saveSuccessMessage}
                    onReset={handleResetComponent}
                />
            </div>
        );
    }
    return (
        <div>
             {fetchError && <div className="my-4 p-3 rounded-md bg-red-100 text-red-700 text-sm"><strong>Lỗi tải dữ liệu NV:</strong> {fetchError}</div>}
            <TimesheetUpload 
                onProcessComplete={setProcessedData}
                employeeDetails={employeeDetails}
            />
        </div>
    );
}
