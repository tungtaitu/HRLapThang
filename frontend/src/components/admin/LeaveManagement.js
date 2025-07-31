/*
 * File: components/admin/LeaveManagement.js
 * Mô tả: Chức năng quản lý phép năm: upload file cấu hình và xuất file tổng hợp.
 */
import React, { useState, useEffect } from 'react';
import { apiUploadLeaveFile, apiExportLeaveFile, apiGetGroups } from '../../api';

export default function LeaveManagementComponent() {
    const [selectedFile, setSelectedFile] = useState(null);
    const [isUploading, setIsUploading] = useState(false);
    const [isExporting, setIsExporting] = useState(false);
    const [message, setMessage] = useState('');
    const [exportYear, setExportYear] = useState(new Date().getFullYear());
    const [groups, setGroups] = useState([]);
    const [selectedGroupId, setSelectedGroupId] = useState('ALL');

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

    const handleFileChange = (event) => { setSelectedFile(event.target.files[0]); setMessage(''); };
    
    const handleUpload = async () => {
        if (!selectedFile) { setMessage('Vui lòng chọn một file Excel để upload.'); return; }
        setIsUploading(true);
        setMessage('');
        try {
            const result = await apiUploadLeaveFile(selectedFile);
            setMessage(result.message);
        } catch (error) {
            setMessage(`Lỗi: ${error.message}`);
        } finally {
            setIsUploading(false);
            setSelectedFile(null);
            if(document.getElementById('file-input')) {
                document.getElementById('file-input').value = null;
            }
        }
    };
    
    const handleExport = async () => {
        setIsExporting(true);
        await apiExportLeaveFile(exportYear, selectedGroupId);
        setIsExporting(false);
    };
    
    const startYear = new Date().getFullYear() + 1;
    const years = Array.from({ length: 10 }, (_, i) => startYear - i);
    
    return (
        <div>
             <div className="mb-8">
                 <h2 className="text-2xl font-bold text-gray-800 mb-4">Upload Dữ liệu Cấu hình Phép Năm</h2>
                 <div className="border-2 border-dashed border-gray-300 rounded-lg p-6 text-center">
                    <p className="mb-4 text-gray-500">
                        Chọn file Excel (.xlsx) chứa dữ liệu phép năm.
                        <br/>File cần có các cột <code className="font-bold bg-gray-200 px-1 rounded">MSNV</code>, <code className="font-bold bg-gray-200 px-1 rounded">Month</code>, <code className="font-bold bg-gray-200 px-1 rounded">PHEP</code>.
                        <br/>Để xác định lao động nước ngoài, thêm cột <code className="font-bold bg-gray-200 px-1 rounded">NUOCNGOAI</code> và điền 'x' hoặc 'yes'.
                    </p>
                    <input id="file-input" type="file" accept=".xlsx, .xls" onChange={handleFileChange} className="block w-full max-w-xs mx-auto text-sm text-gray-500 file:mr-4 file:py-2 file:px-4 file:rounded-full file:border-0 file:text-sm file:font-semibold file:bg-indigo-50 file:text-indigo-700 hover:file:bg-indigo-100" />
                    {selectedFile && <p className="mt-4 text-sm text-gray-600">File đã chọn: {selectedFile.name}</p>}
                    <button onClick={handleUpload} disabled={!selectedFile || isUploading} className="mt-6 px-6 py-2 font-semibold text-white bg-indigo-600 rounded-md hover:bg-indigo-700 disabled:bg-gray-400">
                        {isUploading ? 'Đang xử lý...' : 'Upload và Cập nhật'}
                    </button>
                    {message && <p className={`mt-4 text-sm ${message.startsWith('Lỗi') ? 'text-red-600' : 'text-green-600'}`}>{message}</p>}
                 </div>
             </div>
             <div>
                 <h2 className="text-2xl font-bold text-gray-800 mb-4">Xuất File Tổng hợp Phép Năm</h2>
                 <div className="border-2 border-dashed border-gray-300 rounded-lg p-6 text-center">
                    <p className="mb-4 text-gray-500">Chọn năm và bộ phận để xuất báo cáo. File sẽ có thêm cột "Đối Tượng".</p>
                    <div className="flex flex-col sm:flex-row justify-center items-center gap-4">
                        <div>
                            <label htmlFor="leave-year" className="block text-sm font-medium text-gray-700">Chọn năm</label>
                            <select id="leave-year" value={exportYear} onChange={(e) => setExportYear(parseInt(e.target.value))} className="mt-1 px-3 py-2 border border-gray-300 rounded-md shadow-sm">
                                {years.map(year => <option key={year} value={year}>{year}</option>)}
                            </select>
                        </div>
                        <div>
                            <label htmlFor="leave-group" className="block text-sm font-medium text-gray-700">Chọn bộ phận</label>
                            <select id="leave-group" value={selectedGroupId} onChange={(e) => setSelectedGroupId(e.target.value)} className="mt-1 block w-full pl-3 pr-10 py-2 text-base border-gray-300 focus:outline-none focus:ring-indigo-500 focus:border-indigo-500 sm:text-sm rounded-md">
                                <option value="ALL">Tất cả bộ phận</option>
                                {groups.map(group => (<option key={group.groupId} value={group.groupId}>{group.groupId} - {group.groupName}</option>))}
                            </select>
                        </div>
                        <div className="self-end">
                            <button onClick={handleExport} disabled={isExporting} className="px-6 py-2 font-semibold text-white bg-green-600 rounded-md hover:bg-green-700 disabled:bg-gray-400">
                                {isExporting ? 'Đang xử lý...' : 'Xuất Excel'}
                            </button>
                        </div>
                    </div>
                 </div>
             </div>
        </div>
    );
}
