/*
 * File: components/admin/LuongT13Management.js
 * Mô tả: Chức năng upload dữ liệu lương tháng 13.
 */
import React, { useState } from 'react';
import { apiUploadLuongT13 } from '../../api';

export default function LuongT13ManagementComponent() {
    const [selectedFile, setSelectedFile] = useState(null);
    const [isUploading, setIsUploading] = useState(false);
    const [message, setMessage] = useState({ type: '', text: '' });
    const [uploadYear, setUploadYear] = useState(new Date().getFullYear());

    const handleFileChange = (event) => { setSelectedFile(event.target.files[0]); setMessage({ type: '', text: '' }); };

    const handleUpload = async () => {
        if (!selectedFile) { setMessage({ type: 'error', text: 'Vui lòng chọn một file Excel để upload.' }); return; }
        setIsUploading(true);
        setMessage({ type: '', text: '' });
        try {
            const result = await apiUploadLuongT13(selectedFile, uploadYear);
            setMessage({ type: 'success', text: result.message });
        } catch (error) {
            setMessage({ type: 'error', text: `Lỗi: ${error.message}` });
        } finally {
            setIsUploading(false);
            setSelectedFile(null);
            if (document.getElementById('luong-t13-file-input')) document.getElementById('luong-t13-file-input').value = null;
        }
    };

    const startYear = new Date().getFullYear() + 1;
    const years = Array.from({ length: 10 }, (_, i) => startYear - i);

    return (
        <div>
            <h2 className="text-2xl font-bold text-gray-800 mb-4">Upload Dữ liệu Lương Tháng 13</h2>
            <div className="border-2 border-dashed border-gray-300 rounded-lg p-6 text-center space-y-4">
                <p className="text-gray-500">Chọn năm áp dụng và file Excel chứa dữ liệu lương tháng 13. <br/><span className="font-semibold">Lưu ý: Tên cột trong file Excel phải khớp với mẫu chuẩn.</span></p>
                <div className="flex flex-col sm:flex-row justify-center items-center gap-4">
                    <div><label htmlFor="upload-year" className="sr-only">Chọn năm</label><select id="upload-year" value={uploadYear} onChange={(e) => setUploadYear(parseInt(e.target.value))} className="px-3 py-2 border border-gray-300 rounded-md shadow-sm">{years.map(year => <option key={year} value={year}>{year}</option>)}</select></div>
                    <input id="luong-t13-file-input" type="file" accept=".xlsx, .xls" onChange={handleFileChange} className="block w-full max-w-xs text-sm text-gray-500 file:mr-4 file:py-2 file:px-4 file:rounded-full file:border-0 file:text-sm file:font-semibold file:bg-indigo-50 file:text-indigo-700 hover:file:bg-indigo-100"/>
                </div>
                {selectedFile && <p className="text-sm text-gray-600">File đã chọn: {selectedFile.name}</p>}
                <button onClick={handleUpload} disabled={!selectedFile || isUploading} className="mt-4 px-6 py-2 font-semibold text-white bg-indigo-600 rounded-md hover:bg-indigo-700 disabled:bg-gray-400">{isUploading ? 'Đang xử lý...' : 'Upload và Cập nhật'}</button>
                {message.text && (<p className={`mt-4 text-sm font-semibold ${message.type === 'error' ? 'text-red-600' : 'text-green-600'}`}>{message.text}</p>)}
            </div>
        </div>
    );
}
