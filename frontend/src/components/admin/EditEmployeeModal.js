/*
 * File: components/admin/EditEmployeeModal.js
 * Mô tả: Component modal hiển thị form để sửa thông tin nhân viên.
 * Đã cập nhật để hiển thị và sửa các trường liên quan đến EMPFILEB và ZUNO.
 * Định dạng ngày tháng đầu vào/đầu ra là DD/MM/YYYY.
 * Cập nhật: PASSPORTNO là Ngày cấp căn cước, PIssueDate là rỗng.
 * Cập nhật thêm: Gửi AGE và đảm bảo các trường EMPFILE là rỗng nếu không nhập.
 * Đã sửa lỗi validation cho AGES bằng cách chuyển đổi tường minh sang chuỗi số hoặc null.
 * Cập nhật: Bỏ tên tiếng Hoa và thêm nút Hủy/Đóng.
 */
import React, { useState, useEffect } from 'react';
import { apiGetEmployeeInfo, apiUpdateEmployee, apiGetBasicCodeOptions } from '../../api';

// Hàm chuyển đổi định dạng ngày tháng sang DD/MM/YYYY
const formatDateToDDMMYYYY = (dateInput) => {
    if (!dateInput) return '';
    const date = new Date(dateInput);
    // Adjust for timezone offset to get the correct local date
    const userTimezoneOffset = date.getTimezoneOffset() * 60000;
    const adjustedDate = new Date(date.getTime() + userTimezoneOffset);
    if (isNaN(adjustedDate.getTime())) return '';
    const day = String(adjustedDate.getDate()).padStart(2, '0');
    const month = String(adjustedDate.getMonth() + 1).padStart(2, '0');
    const year = adjustedDate.getFullYear();
    return `${day}/${month}/${year}`;
};

// Hàm tính tuổi từ ngày sinh (DD/MM/YYYY)
const calculateAge = (birthDateString) => {
    if (!birthDateString) return '';
    const parts = birthDateString.split('/').map(Number);
    if (parts.length !== 3) return '';
    const [day, month, year] = parts;
    const birthDate = new Date(year, month - 1, day);
    const today = new Date();
    let age = today.getFullYear() - birthDate.getFullYear();
    const m = today.getMonth() - birthDate.getMonth();
    if (m < 0 || (m === 0 && today.getDate() < birthDate.getDate())) {
        age--;
    }
    return age >= 0 ? age : ''; // Trả về tuổi hoặc chuỗi rỗng nếu ngày sinh trong tương lai
};

export default function EditEmployeeModal({ employeeId, onClose, onSaveSuccess }) {
    const [formData, setFormData] = useState(null);
    const [isLoading, setIsLoading] = useState(true);
    const [error, setError] = useState('');
    const [jobOptions, setJobOptions] = useState([]);
    const [departmentOptions, setDepartmentOptions] = useState([]);
    const [whsnoOptions, setWhsnoOptions] = useState([]);
    const [allZunoOptions, setAllZunoOptions] = useState([]); 
    const [filteredZunoOptions, setFilteredZunoOptions] = useState([]); 

    useEffect(() => {
        const fetchEmployeeAndOptionsData = async () => {
            if (!employeeId) return;
            setIsLoading(true);
            setError('');
            try {
                const [employeeData, jobs, departments, whsnos, zunos] = await Promise.all([
                    apiGetEmployeeInfo(employeeId),
                    apiGetBasicCodeOptions('LEV'),
                    apiGetBasicCodeOptions('GROUPID'),
                    apiGetBasicCodeOptions('WHSNO'),
                    apiGetBasicCodeOptions('ZUNO')
                ]);
                
                setJobOptions(jobs);
                setDepartmentOptions(departments);
                setWhsnoOptions(whsnos);
                setAllZunoOptions(zunos); 

                // Ánh xạ PASSPORTNO từ DB thành idCardIssueDate cho frontend
                employeeData.idCardIssueDate = employeeData.PASSPORTNO || '';
                // Đảm bảo PIssueDate luôn là rỗng theo yêu cầu mới
                employeeData.PIssueDate = ''; 
                // Đảm bảo b_shift mặc định là 'ALL' nếu không có giá trị từ DB
                employeeData.b_shift = employeeData.b_shift || 'ALL';

                setFormData(employeeData);
            } catch (err) {
                setError(err.message);
            } finally {
                setIsLoading(false);
            }
        };
        fetchEmployeeAndOptionsData();
    }, [employeeId]);

    // Effect để lọc ZUNO options khi GROUPID thay đổi trong modal
    useEffect(() => {
        if (formData && formData.GROUPID && allZunoOptions.length > 0) {
            const filtered = allZunoOptions.filter(opt => opt.SYS_TYPE.startsWith(formData.GROUPID));
            setFilteredZunoOptions(filtered);
            if (!filtered.some(opt => opt.SYS_TYPE === formData.ZUNO)) {
                setFormData(prev => ({ ...prev, ZUNO: '' }));
            }
        } else {
            setFilteredZunoOptions([]); 
            if (formData) { 
                setFormData(prev => ({ ...prev, ZUNO: '' }));
            }
        }
    }, [formData?.GROUPID, allZunoOptions]); 

    const handleChange = (e) => {
        const { name, value } = e.target;
        setFormData(prev => ({ ...prev, [name]: value }));
    };

    const handleSubmit = async (e) => {
        e.preventDefault();
        setIsLoading(true);
        setError('');
        try {
            let BYY = null, BMM = null, BDD = null; 
            const ageCalculated = calculateAge(formData.birthDate);

            if (formData.birthDate) {
                const parts = formData.birthDate.split('/').map(Number);
                if (parts.length === 3 && !isNaN(parts[0]) && !isNaN(parts[1]) && !isNaN(parts[2])) {
                    BDD = parts[0];
                    BMM = parts[1];
                    BYY = parts[2];
                }
            }

            const payload = { 
                ...formData, 
                BYY: BYY !== null ? String(BYY) : null, 
                BMM: BMM !== null ? String(BMM) : null, 
                BDD: BDD !== null ? String(BDD) : null, 
                AGES: (typeof ageCalculated === 'number' && !isNaN(ageCalculated)) 
                      ? String(ageCalculated) 
                      : null,
                INDAT: formData.INDAT,
                OUTDAT: formData.OUTDAT,
                PIssueDate: '',
                PASSPORTNO: formData.idCardIssueDate || '',
                EMPNAM_CN: '', // Luôn gửi rỗng
                empType: formData.empType || '',
                TX: formData.TX || 0.0,
                BHDAT: formData.BHDAT || '',
                GTDAT: formData.GTDAT || '',
                MEMO: formData.MEMO || '',
                PDUEDATE: formData.PDUEDATE || '',
                VDUEDATE: formData.VDUEDATE || '',
                StudyJob: formData.StudyJob || '',
                Grps: formData.Grps || '',
                SEX: formData.SEX || '',
                PERSONID: formData.PERSONID || '',
                HOMEADDR: formData.HOMEADDR || '',
                PHONE: formData.PHONE || '',
                MOBILEPHONE: formData.MOBILEPHONE || '',
                EMAIL: formData.EMAIL || '',
                MARRYED: formData.MARRYED || '',
                SCHOOL: formData.SCHOOL || '',
                COUNTRY: formData.COUNTRY || '',
                BANKID: formData.BANKID || '',
                taxCode: formData.taxCode || '',
                VISANO: formData.VISANO || '',
                GROUPID: formData.GROUPID || '',
                JOB: formData.JOB || '',
                WHSNO: formData.WHSNO || '',
                ZUNO: formData.ZUNO || '',
                WKD_No: formData.WKD_No || '',
                WKD_dueDate: formData.WKD_dueDate || '',
                experience: formData.experience || '',
                urgent_person: formData.urgent_person || '',
                releation: formData.releation || '',
                urgent_addr: formData.urgent_addr || '',
                urgent_tel: formData.urgent_tel || '',
                urgent_mobile: formData.urgent_mobile || '',
                bh_person: formData.bh_person || '',
                bh_personID: formData.bh_personID || '',
                b_shift: formData.b_shift || 'ALL',
                soBH: formData.soBH || ''
            };
            
            await apiUpdateEmployee(employeeId, payload);
            onSaveSuccess(); 
        } catch (err) {
            setError(err.message);
        } finally {
            setIsLoading(false);
        }
    };

    if (isLoading && !formData) return <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50"><p className="text-white">Đang tải thông tin...</p></div>;
    if (error) return <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50"><div className="bg-white p-4 rounded">{error} <button onClick={onClose}>Đóng</button></div></div>;
    if (!formData) return null;

    const currentAge = calculateAge(formData.birthDate);

    return (
        <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50 p-4">
            <div className="bg-white rounded-lg shadow-xl w-full max-w-3xl max-h-[90vh] flex flex-col">
                <h2 className="text-xl font-bold p-6 pb-0">Chỉnh sửa Nhân viên: {formData.EMPNAM_VN} ({formData.EMPID})</h2>
                <form onSubmit={handleSubmit} className="p-6 space-y-4 overflow-y-auto">
                    <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                        <div>
                            <label className="block text-sm font-medium">Họ tên (Việt) (*)</label>
                            <input type="text" name="EMPNAM_VN" value={formData.EMPNAM_VN || ''} onChange={handleChange} className="mt-1 w-full px-3 py-2 border rounded-md" required />
                        </div>
                        <div>
                            <label className="block text-sm font-medium">Giới tính</label>
                            <select name="SEX" value={formData.SEX || 'M'} onChange={handleChange} className="mt-1 w-full px-3 py-2 border rounded-md bg-white">
                                <option value="M">Nam</option>
                                <option value="F">Nữ</option>
                            </select>
                        </div>
                        <div>
                            <label className="block text-sm font-medium">Ngày sinh</label>
                            <input type="text" name="birthDate" value={formData.birthDate || ''} onChange={handleChange} placeholder="dd/mm/yyyy" className="mt-1 w-full px-3 py-2 border rounded-md" />
                            {currentAge && <p className="text-xs text-gray-500 mt-1">Tuổi: {currentAge}</p>}
                        </div>
                        <div>
                            <label className="block text-sm font-medium">Ngày vào làm (*)</label>
                            <input type="text" name="INDAT" value={formData.INDAT || ''} onChange={handleChange} placeholder="dd/mm/yyyy" className="mt-1 w-full px-3 py-2 border rounded-md" required />
                        </div>
                         <div>
                            <label className="block text-sm font-medium">Ngày thôi việc</label>
                            <input type="text" name="OUTDAT" value={formData.OUTDAT || ''} onChange={handleChange} placeholder="dd/mm/yyyy" className="mt-1 w-full px-3 py-2 border rounded-md" />
                        </div>
                        <div>
                            <label className="block text-sm font-medium">Xưởng (*)</label>
                            <select name="WHSNO" value={formData.WHSNO || ''} onChange={handleChange} className="mt-1 w-full px-3 py-2 border rounded-md bg-white" required>
                                <option value="">-- Chọn xưởng --</option>
                                {whsnoOptions.map(opt => <option key={opt.SYS_TYPE} value={opt.SYS_TYPE}>{opt.SYS_TYPE} - {opt.SYS_VALUE}</option>)}
                            </select>
                        </div>
                        <div>
                            <label className="block text-sm font-medium">Bộ phận</label>
                            <select name="GROUPID" value={formData.GROUPID || ''} onChange={handleChange} className="mt-1 w-full px-3 py-2 border rounded-md bg-white">
                                <option value="">-- Chọn bộ phận --</option>
                                {departmentOptions.map(opt => <option key={opt.SYS_TYPE} value={opt.SYS_TYPE}>{opt.SYS_TYPE} - {opt.SYS_VALUE}</option>)}
                            </select>
                        </div>
                        <div>
                            <label className="block text-sm font-medium">Tổ</label>
                            <select name="ZUNO" value={formData.ZUNO || ''} onChange={handleChange} className="mt-1 w-full px-3 py-2 border rounded-md bg-white" disabled={!formData.GROUPID || filteredZunoOptions.length === 0}>
                                <option value="">-- Chọn Tổ --</option>
                                {filteredZunoOptions.map(opt => <option key={opt.SYS_TYPE} value={opt.SYS_TYPE}>{opt.SYS_TYPE} - {opt.SYS_VALUE}</option>)}
                            </select>
                        </div>
                        <div>
                            <label className="block text-sm font-medium">Chức vụ</label>
                            <select name="JOB" value={formData.JOB || ''} onChange={handleChange} className="mt-1 w-full px-3 py-2 border rounded-md bg-white">
                                <option value="">-- Chọn chức vụ --</option>
                                {jobOptions.map(opt => <option key={opt.SYS_TYPE} value={opt.SYS_TYPE}>{opt.SYS_TYPE} - {opt.SYS_VALUE}</option>)}
                            </select>
                        </div>
                        <div>
                            <label className="block text-sm font-medium">Ca làm việc</label>
                            <input type="text" name="b_shift" value={formData.b_shift || ''} onChange={handleChange} className="mt-1 w-full px-3 py-2 border rounded-md" />
                        </div>
                        <div>
                            <label className="block text-sm font-medium">Số CMND/CCCD</label>
                            <input type="text" name="PERSONID" value={formData.PERSONID || ''} onChange={handleChange} className="mt-1 w-full px-3 py-2 border rounded-md" />
                        </div>
                        <div>
                            <label className="block text-sm font-medium">Ngày cấp căn cước</label>
                            <input type="text" name="idCardIssueDate" value={formData.idCardIssueDate || ''} onChange={handleChange} placeholder="dd/mm/yyyy" className="mt-1 w-full px-3 py-2 border rounded-md" />
                        </div>
                         <div>
                        <label className="block text-sm font-medium">Nơi cấp (CMND/CCCD)</label>
                            <select name="VISANO" value={formData.VISANO} onChange={handleChange} className="mt-1 w-full px-3 py-2 border rounded-md  bg-white">
                                <option value="Cục Cảnh sát quản lý hành chính về trật tự xã hội">Cục Cảnh sát quản lý hành chính về trật tự xã hội</option>
                                <option value="Bộ Công an">Bộ Công an</option>
                                <option value="Cục Cảnh sát đăng ký quản lý cư trú và dữ liệu Quốc gia về dân cư">Cục Cảnh sát đăng ký quản lý cư trú và dữ liệu Quốc gia về dân cư</option>
                            </select>
                        </div>
                        <div>
                            <label className="block text-sm font-medium">Tình trạng hôn nhân</label>
                            <select name="MARRYED" value={formData.MARRYED} onChange={handleChange} className="mt-1 w-full px-3 py-2 border rounded-md bg-white">
                                <option value="">-----------</option>
                                <option value="F">Chưa kết hôn</option>
                                <option value="T">Đã kết hôn</option>
                            </select>
                        </div>
                        <div>
                            <label className="block text-sm font-medium">Địa chỉ</label>
                            <input type="text" name="HOMEADDR" value={formData.HOMEADDR || ''} onChange={handleChange} className="mt-1 w-full px-3 py-2 border rounded-md" />
                        </div>
                        <div>
                            <label className="block text-sm font-medium">Điện thoại</label>
                            <input type="text" name="PHONE" value={formData.PHONE || ''} onChange={handleChange} className="mt-1 w-full px-3 py-2 border rounded-md" />
                        </div>
                        <div>
                            <label className="block text-sm font-medium">Email</label>
                            <input type="email" name="EMAIL" value={formData.EMAIL || ''} onChange={handleChange} className="mt-1 w-full px-3 py-2 border rounded-md" />
                        </div>
                        <div>
                            <label className="block text-sm font-medium">Mã số thuế</label>
                            <input type="text" name="taxCode" value={formData.taxCode || ''} onChange={handleChange} className="mt-1 w-full px-3 py-2 border rounded-md" />
                        </div>
                        <div>
                            <label className="block text-sm font-medium">Số tài khoản ngân hàng</label>
                            <input type="text" name="BANKID" value={formData.BANKID || ''} onChange={handleChange} className="mt-1 w-full px-3 py-2 border rounded-md" />
                        </div>
                        <div>
                            <label className="block text-sm font-medium">Trình độ</label>
                            <input type="text" name="SCHOOL" value={formData.SCHOOL || ''} onChange={handleChange} className="mt-1 w-full px-3 py-2 border rounded-md" />
                        </div>
                        <div>
                            <label className="block text-sm font-medium">Quốc tịch</label>
                            <select name="COUNTRY" value={formData.COUNTRY || ''} onChange={handleChange} className="mt-1 w-full px-3 py-2 border rounded-md">
                                <option value="VN">Việt Nam</option>
                                <option value="CT">Nước ngoài</option>
                            </select>
                        </div>
                    </div>

                    <div className="flex justify-end pt-4 gap-x-2">
                        <button type="button" onClick={onClose} className="px-6 py-2 font-semibold text-gray-700 bg-gray-200 rounded-md hover:bg-gray-300">
                            Hủy
                        </button>
                        <button type="submit" disabled={isLoading} className="px-6 py-2 font-semibold text-white bg-blue-600 rounded-md hover:bg-blue-700 disabled:bg-gray-400">
                            {isLoading ? 'Đang lưu...' : 'Lưu thay đổi'}
                        </button>
                    </div>
                </form>
            </div>
        </div>
    );
}
