/*
 * File: components/admin/AddEmployee.js
 * Mô tả: Component form chi tiết để thêm nhân viên mới vào hệ thống.
 */
import React, { useState, useEffect } from 'react';
import { apiAddEmployee, apiGetBasicCodeOptions, apiGetNextEmployeeId } from '../../api';

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

export default function AddEmployee() {
    const initialFormState = {
        COUNTRY: 'VN', EMPID: '', birthDate: '', MARRYED: '', PERSONID: '',
        idCardIssueDate: '', 
        VISANO: '', masobh: '', BANKID: '', HOMEADDR: '',
        INDAT: formatDateToDDMMYYYY(new Date()), OUTDAT: '',
        EMPNAM_CN: '', EMPNAM_VN: '', SEX: 'M', JOB: '', SCHOOL: '', taxCode: '',
        PHONE: '', EMAIL: '', GROUPID: '', WHSNO: 'LT', ZUNO: '', // Mặc định xưởng là LT
        // Các trường bổ sung từ EMPFILE
        empType: '', TX: 0.0, BHDAT: '', GTDAT: '', MEMO: '', AGES: '', PDUEDATE: '', VDUEDATE: '', StudyJob: '', Grps: '',
        // Các trường từ EMPFILEB
        WKD_No: '', WKD_dueDate: '', experience: '', urgent_person: '', releation: '',
        urgent_addr: '', urgent_tel: '', urgent_mobile: '', bh_person: '', bh_personID: '', b_shift: 'ALL', soBH: ''
    };
    const [formData, setFormData] = useState(initialFormState);
    const [isLoading, setIsLoading] = useState(false);
    const [message, setMessage] = useState({ type: '', text: '' });
    const [jobOptions, setJobOptions] = useState([]);
    const [departmentOptions, setDepartmentOptions] = useState([]);
    const [whsnoOptions, setWhsnoOptions] = useState([]);
    const [allZunoOptions, setAllZunoOptions] = useState([]); 
    const [filteredZunoOptions, setFilteredZunoOptions] = useState([]); 
    const [isIdLoading, setIsIdLoading] = useState(true);

    useEffect(() => {
        const fetchInitialData = async () => {
            setIsIdLoading(true);
            try {
                const [jobs, departments, nextIdResponse, whsnos, zunos] = await Promise.all([
                    apiGetBasicCodeOptions('LEV'),
                    apiGetBasicCodeOptions('GROUPID'),
                    apiGetNextEmployeeId(),
                    apiGetBasicCodeOptions('WHSNO'),
                    apiGetBasicCodeOptions('ZUNO') 
                ]);
                setJobOptions(jobs);
                setDepartmentOptions(departments);
                setWhsnoOptions(whsnos);
                setAllZunoOptions(zunos); 
                setFormData(prev => ({ ...prev, EMPID: nextIdResponse.nextId }));
            } catch (error) {
                console.error("Không thể tải dữ liệu ban đầu:", error);
                setMessage({ type: 'error', text: 'Không thể tải dữ liệu cần thiết. Vui lòng thử lại.' });
            } finally {
                setIsIdLoading(false);
            }
        };
        fetchInitialData();
    }, []);

    // Effect để lọc ZUNO options khi GROUPID thay đổi
    useEffect(() => {
        if (formData.GROUPID && allZunoOptions.length > 0) {
            const filtered = allZunoOptions.filter(opt => opt.SYS_TYPE.startsWith(formData.GROUPID));
            setFilteredZunoOptions(filtered);
            if (!filtered.some(opt => opt.SYS_TYPE === formData.ZUNO)) {
                setFormData(prev => ({ ...prev, ZUNO: '' }));
            }
        } else {
            setFilteredZunoOptions([]); 
            setFormData(prev => ({ ...prev, ZUNO: '' })); 
        }
    }, [formData.GROUPID, allZunoOptions]);

    const handleChange = (e) => {
        const { name, value } = e.target;
        setFormData(prev => ({ ...prev, [name]: value }));
    };

    const handleCancel = () => {
        // Reset form to initial state, but keep the next employee ID
        setFormData(prev => ({
            ...initialFormState,
            EMPID: prev.EMPID 
        }));
        setMessage({ type: '', text: '' });
    };

    const handleSubmit = async (e) => {
        e.preventDefault();
        setIsLoading(true);
        setMessage({ type: '', text: '' });
        
        const { birthDate, idCardIssueDate, ...rest } = formData; 
        let BYY = null, BMM = null, BDD = null; 
        const ageCalculated = calculateAge(birthDate);

        if (birthDate) {
            const parts = birthDate.split('/').map(Number);
            if (parts.length === 3 && !isNaN(parts[0]) && !isNaN(parts[1]) && !isNaN(parts[2])) {
                BDD = parts[0];
                BMM = parts[1];
                BYY = parts[2];
            }
        }

        const payload = { 
            ...rest, 
            BYY: BYY !== null ? String(BYY) : null, 
            BMM: BMM !== null ? String(BMM) : null, 
            BDD: BDD !== null ? String(BDD) : null, 
            AGES: (typeof ageCalculated === 'number' && !isNaN(ageCalculated)) 
                  ? String(ageCalculated) 
                  : null,
            INDAT: formData.INDAT,
            OUTDAT: formData.OUTDAT,
            PIssueDate: '',
            PASSPORTNO: idCardIssueDate || '',
            EMPNAM_CN: '', // Luôn gửi rỗng
            empType: rest.empType || '',
            TX: rest.TX || 0.0,
            BHDAT: rest.BHDAT || '',
            GTDAT: rest.GTDAT || '',
            MEMO: rest.MEMO || '',
            PDUEDATE: rest.PDUEDATE || '',
            VDUEDATE: rest.VDUEDATE || '',
            StudyJob: rest.StudyJob || '',
            Grps: rest.Grps || '',
            SEX: rest.SEX || '',
            PERSONID: rest.PERSONID || '',
            HOMEADDR: rest.HOMEADDR || '',
            PHONE: rest.PHONE || '',
            MOBILEPHONE: rest.MOBILEPHONE || '',
            EMAIL: rest.EMAIL || '',
            MARRYED: rest.MARRYED || '',
            SCHOOL: rest.SCHOOL || '',
            COUNTRY: rest.COUNTRY || '',
            BANKID: rest.BANKID || '',
            taxCode: rest.taxCode || '',
            VISANO: rest.VISANO || '',
            GROUPID: rest.GROUPID || '',
            JOB: rest.JOB || '',
            WHSNO: rest.WHSNO || '',
            ZUNO: rest.ZUNO || '',
            WKD_No: rest.WKD_No || '',
            WKD_dueDate: rest.WKD_dueDate || '',
            experience: rest.experience || '',
            urgent_person: rest.urgent_person || '',
            releation: rest.releation || '',
            urgent_addr: rest.urgent_addr || '',
            urgent_tel: rest.urgent_tel || '',
            urgent_mobile: rest.urgent_mobile || '',
            bh_person: rest.bh_person || '',
            bh_personID: rest.bh_personID || '',
            b_shift: rest.b_shift || 'ALL',
            soBH: rest.soBH || ''
        };

        try {
            const result = await apiAddEmployee(payload);
            setMessage({ type: 'success', text: result.message });
            const nextIdResponse = await apiGetNextEmployeeId();
            setFormData({ ...initialFormState, EMPID: nextIdResponse.nextId });
        } catch (error) {
            setMessage({ type: 'error', text: error.message });
        } finally {
            setIsLoading(false);
        }
    };

    const currentAge = calculateAge(formData.birthDate);

    return (
        <div>
            <h2 className="text-2xl font-bold text-gray-800 mb-6">Thêm Nhân viên Mới</h2>
            <form onSubmit={handleSubmit} className="space-y-6">
                {message.text && (
                    <div className={`p-3 rounded-md ${message.type === 'success' ? 'bg-green-100 text-green-800' : 'bg-red-100 text-red-800'}`}>
                        {message.text}
                    </div>
                )}
                <div className="p-4 border rounded-md bg-gray-50">
                    <h3 className="text-lg font-semibold mb-4 text-gray-700">Thông tin Cơ bản</h3>
                    <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                        <div>
                            <label className="block text-sm font-medium">Xưởng (*)</label>
                            <select name="WHSNO" value={formData.WHSNO} onChange={handleChange} className="mt-1 w-full px-3 py-2 border rounded-md bg-white" required>
                                <option value="">-- Chọn xưởng --</option>
                                {whsnoOptions.map(opt => <option key={opt.SYS_TYPE} value={opt.SYS_TYPE}>{opt.SYS_TYPE} - {opt.SYS_VALUE}</option>)}
                            </select>
                        </div>
                        <div>
                            <label className="block text-sm font-medium">Mã số nhân viên (*)</label>
                            <input type="text" name="EMPID" value={isIdLoading ? 'Đang tải...' : formData.EMPID} className="mt-1 w-full px-3 py-2 border rounded-md bg-gray-100" readOnly />
                        </div>
                        <div>
                            <label className="block text-sm font-medium">Họ tên (Việt) (*)</label>
                            <input type="text" name="EMPNAM_VN" value={formData.EMPNAM_VN} onChange={handleChange} className="mt-1 w-full px-3 py-2 border rounded-md" required />
                        </div>
                        <div>
                            <label className="block text-sm font-medium">Giới tính</label>
                            <select name="SEX" value={formData.SEX} onChange={handleChange} className="mt-1 w-full px-3 py-2 border rounded-md bg-white">
                                <option value="M">Nam</option>
                                <option value="F">Nữ</option>
                            </select>
                        </div>
                        <div>
                            <label className="block text-sm font-medium">Ngày sinh</label>
                            <input type="text" name="birthDate" value={formData.birthDate} onChange={handleChange} placeholder="dd/mm/yyyy" className="mt-1 w-full px-3 py-2 border rounded-md" />
                            {currentAge && <p className="text-xs text-gray-500 mt-1">Tuổi: {currentAge}</p>}
                        </div>
                    </div>
                </div>

                <div className="p-4 border rounded-md bg-gray-50">
                    <h3 className="text-lg font-semibold mb-4 text-gray-700">Thông tin Công việc</h3>
                     <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                        <div>
                            <label className="block text-sm font-medium">Ngày vào làm (*)</label>
                            <input type="text" name="INDAT" value={formData.INDAT} onChange={handleChange} placeholder="dd/mm/yyyy" className="mt-1 w-full px-3 py-2 border rounded-md" required />
                        </div>
                        <div>
                            <label className="block text-sm font-medium">Ngày thôi việc</label>
                            <input type="text" name="OUTDAT" value={formData.OUTDAT} onChange={handleChange} placeholder="dd/mm/yyyy" className="mt-1 w-full px-3 py-2 border rounded-md" />
                        </div>
                         <div>
                            <label className="block text-sm font-medium">Bộ phận</label>
                            <select name="GROUPID" value={formData.GROUPID} onChange={handleChange} className="mt-1 w-full px-3 py-2 border rounded-md bg-white">
                                <option value="">-- Chọn bộ phận --</option>
                                {departmentOptions.map(opt => <option key={opt.SYS_TYPE} value={opt.SYS_TYPE}>{opt.SYS_TYPE} - {opt.SYS_VALUE}</option>)}
                            </select>
                        </div>
                         <div>
                            <label className="block text-sm font-medium">Tổ</label>
                            <select name="ZUNO" value={formData.ZUNO} onChange={handleChange} className="mt-1 w-full px-3 py-2 border rounded-md bg-white" disabled={!formData.GROUPID || filteredZunoOptions.length === 0}>
                                <option value="">-- Chọn Tổ --</option>
                                {filteredZunoOptions.map(opt => <option key={opt.SYS_TYPE} value={opt.SYS_TYPE}>{opt.SYS_TYPE} - {opt.SYS_VALUE}</option>)}
                            </select>
                        </div>
                         <div>
                            <label className="block text-sm font-medium">Chức vụ</label>
                            <select name="JOB" value={formData.JOB} onChange={handleChange} className="mt-1 w-full px-3 py-2 border rounded-md bg-white">
                                <option value="">-- Chọn chức vụ --</option>
                                {jobOptions.map(opt => <option key={opt.SYS_TYPE} value={opt.SYS_TYPE}>{opt.SYS_TYPE} - {opt.SYS_VALUE}</option>)}
                            </select>
                        </div>
                        <div>
                            <label className="block text-sm font-medium">Ca làm việc</label>
                            <input type="text" name="b_shift" value={formData.b_shift} onChange={handleChange} className="mt-1 w-full px-3 py-2 border rounded-md" />
                        </div>
                    </div>
                </div>
                
                <div className="p-4 border rounded-md bg-gray-50">
                    <h3 className="text-lg font-semibold mb-4 text-gray-700">Thông tin Cá nhân & Liên lạc</h3>
                     <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                        <div>
                            <label className="block text-sm font-medium">Số CMND/CCCD</label>
                            <input type="text" name="PERSONID" value={formData.PERSONID} onChange={handleChange} className="mt-1 w-full px-3 py-2 border rounded-md" />
                        </div>
                        <div>
                            <label className="block text-sm font-medium">Ngày cấp căn cước</label>
                            <input type="text" name="idCardIssueDate" value={formData.idCardIssueDate} onChange={handleChange} placeholder="dd/mm/yyyy" className="mt-1 w-full px-3 py-2 border rounded-md" />
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
                            <input type="text" name="HOMEADDR" value={formData.HOMEADDR} onChange={handleChange} className="mt-1 w-full px-3 py-2 border rounded-md" />
                        </div>
                        <div>
                            <label className="block text-sm font-medium">Điện thoại</label>
                            <input type="text" name="PHONE" value={formData.PHONE} onChange={handleChange} className="mt-1 w-full px-3 py-2 border rounded-md" />
                        </div>
                        <div>
                            <label className="block text-sm font-medium">Email</label>
                            <input type="email" name="EMAIL" value={formData.EMAIL} onChange={handleChange} className="mt-1 w-full px-3 py-2 border rounded-md" />
                        </div>
                        <div>
                            <label className="block text-sm font-medium">Mã số thuế</label>
                            <input type="text" name="taxCode" value={formData.taxCode} onChange={handleChange} className="mt-1 w-full px-3 py-2 border rounded-md" />
                        </div>
                        <div>
                            <label className="block text-sm font-medium">Số tài khoản ngân</label>
                            <input type="text" name="BANKID" value={formData.BANKID} onChange={handleChange} className="mt-1 w-full px-3 py-2 border rounded-md" />
                        </div>
                        <div>
                            <label className="block text-sm font-medium">Trình độ</label>
                            <input type="text" name="SCHOOL" value={formData.SCHOOL} onChange={handleChange} className="mt-1 w-full px-3 py-2 border rounded-md" />
                        </div>
                        <div>
                            <label className="block text-sm font-medium">Quốc tịch</label>
                            <select name="COUNTRY" value={formData.COUNTRY || ''} onChange={handleChange} className="mt-1 w-full px-3 py-2 border rounded-md">
                                <option value="VN">Việt Nam</option>
                                <option value="CT">Nước ngoài</option>
                            </select>
                        </div>
                    </div>
                </div>

                <div className="flex justify-end pt-4 gap-x-2">
                    <button type="button" onClick={handleCancel} className="px-6 py-2 font-semibold text-gray-700 bg-gray-200 rounded-md hover:bg-gray-300">
                        Hủy
                    </button>
                    <button type="submit" disabled={isLoading} className="px-6 py-2 font-semibold text-white bg-green-600 rounded-md hover:bg-green-700 disabled:bg-gray-400">
                        {isLoading ? 'Đang lưu...' : 'Xác nhận Thêm'}
                    </button>
                </div>
            </form>
        </div>
    );
}
