import React, { useState, useEffect } from 'react';

// --- HÀM GỌI API ---
// Chúng ta sẽ định nghĩa các hàm API ngay tại đây để component này được độc lập
const API_URL = 'https://api.nhansulapthang.io.vn/api/leave'; // Đường dẫn tới module API nghỉ phép mới

const apiFetchApprovers = async () => {
    const res = await fetch(`${API_URL}/approvers`);
    if (!res.ok) throw new Error('Không thể tải danh sách người duyệt.');
    return res.json();
};

const apiFetchRequests = async (user) => {
    const params = new URLSearchParams({ userId: user.id, isAdmin: user.isAdmin });
    const res = await fetch(`${API_URL}/requests?${params}`);
    if (!res.ok) throw new Error('Không thể tải danh sách yêu cầu.');
    return res.json();
};

const apiSubmitRequest = async (requestData) => {
    const res = await fetch(`${API_URL}/requests`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(requestData),
    });
    if (!res.ok) throw new Error('Gửi yêu cầu thất bại.');
    return res.json();
};

const apiUpdateRequestStatus = async (id, status) => {
    const res = await fetch(`${API_URL}/requests/${id}`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ status }),
    });
    if (!res.ok) throw new Error('Cập nhật trạng thái thất bại.');
    return res.json();
};

const apiDeleteRequest = async (id) => {
    const res = await fetch(`${API_URL}/requests/${id}`, { method: 'DELETE' });
    if (!res.ok) throw new Error('Xóa yêu cầu thất bại.');
    return res.json();
};

// --- COMPONENT CHÍNH ---
export default function LeaveManagement({ user }) {
    const [requests, setRequests] = useState([]);
    const [approvers, setApprovers] = useState([]);
    const [isLoading, setIsLoading] = useState(true);
    const [error, setError] = useState(null);

    // State cho form
    const [startDate, setStartDate] = useState('');
    const [endDate, setEndDate] = useState('');
    const [reason, setReason] = useState('');
    const [approverId, setApproverId] = useState('');

    // Hàm tải dữ liệu
    const fetchData = async () => {
        try {
            setIsLoading(true);
            const [requestsData, approversData] = await Promise.all([
                apiFetchRequests(user),
                apiFetchApprovers()
            ]);
            setRequests(requestsData);
            setApprovers(approversData);
            if (approversData.length > 0) {
                setApproverId(approversData[0].id); // Mặc định chọn người đầu tiên
            }
        } catch (err) {
            setError(err.message);
        } finally {
            setIsLoading(false);
        }
    };
    
    // Tải dữ liệu khi component được render
    useEffect(() => {
        fetchData();
    }, [user]);

    // Xử lý Gửi Form
    const handleSubmit = async (e) => {
        e.preventDefault();
        if (!startDate || !endDate || !reason || !approverId) {
            alert('Vui lòng điền đủ thông tin.'); return;
        }
        try {
            const requestData = {
                employeeId: user.id,
                employeeName: user.name,
                startDate, endDate, reason, approverId
            };
            await apiSubmitRequest(requestData);
            alert('Gửi yêu cầu thành công!');
            // Reset form và tải lại dữ liệu
            setStartDate(''); setEndDate(''); setReason('');
            fetchData();
        } catch (err) {
            alert(`Lỗi: ${err.message}`);
        }
    };
    
    // Xử lý Duyệt/Từ chối
    const handleUpdate = async (id, status) => {
        try {
            await apiUpdateRequestStatus(id, status);
            fetchData(); // Tải lại danh sách
        } catch (err) {
             alert(`Lỗi: ${err.message}`);
        }
    };

    // Xử lý Xóa
    const handleDelete = async (id) => {
        if(window.confirm('Bạn có chắc muốn xóa yêu cầu này?')) {
            try {
                await apiDeleteRequest(id);
                fetchData();
            } catch (err) {
                 alert(`Lỗi: ${err.message}`);
            }
        }
    }

    // Phân loại yêu cầu
    const pendingRequests = requests.filter(r => r.status === 'pending');
    const approvedRequests = requests.filter(r => r.status === 'approved');
    const rejectedRequests = requests.filter(r => r.status === 'rejected');

    // Các hàm render giao diện con
    const renderStatusBadge = (status) => {
        const styles = {
            pending: 'bg-yellow-100 text-yellow-800',
            approved: 'bg-green-100 text-green-800',
            rejected: 'bg-red-100 text-red-800',
        };
        const text = {
            pending: 'Chờ duyệt',
            approved: 'Đã duyệt',
            rejected: 'Đã từ chối',
        }
        return <span className={`px-2 py-1 text-xs font-medium rounded-full ${styles[status]}`}>{text[status]}</span>;
    };

    const RequestCard = ({ req }) => (
        <div className="bg-white p-4 rounded-lg shadow-sm border border-gray-200">
            <div className="flex justify-between items-start">
                <div>
                    <p className="font-bold text-gray-800">{req.employeeName} <span className="font-normal text-gray-500 text-sm">({req.employeeId})</span></p>
                    <p className="text-sm text-gray-600">
                        {new Date(req.startDate).toLocaleDateString('vi-VN')} - {new Date(req.endDate).toLocaleDateString('vi-VN')}
                    </p>
                    <p className="text-xs text-gray-500 mt-2">Gửi đến: <strong>{req.approverName}</strong></p>
                </div>
                {renderStatusBadge(req.status)}
            </div>
            <p className="text-sm text-gray-700 mt-2 p-2 bg-gray-50 rounded"><strong>Lý do:</strong> {req.reason}</p>
            <div className="flex justify-end items-center gap-2 mt-3">
                {user.isAdmin && req.status === 'pending' && (
                    <>
                        <button onClick={() => handleUpdate(req.id, 'approved')} className="px-3 py-1 text-xs font-semibold text-white bg-green-500 rounded hover:bg-green-600">Duyệt</button>
                        <button onClick={() => handleUpdate(req.id, 'rejected')} className="px-3 py-1 text-xs font-semibold text-white bg-red-500 rounded hover:bg-red-600">Từ chối</button>
                    </>
                )}
                 {(user.id === req.employeeId || user.isAdmin) && (
                     <button onClick={() => handleDelete(req.id)} className="px-3 py-1 text-xs font-semibold text-gray-600 bg-gray-200 rounded hover:bg-gray-300">Xóa</button>
                 )}
            </div>
        </div>
    );
    
    if (isLoading) return <p>Đang tải dữ liệu...</p>;
    if (error) return <p className="text-red-500">Lỗi: {error}</p>;

    return (
        <div className="grid grid-cols-1 lg:grid-cols-3 gap-8">
            {/* CỘT TẠO YÊU CẦU - Chỉ hiện cho nhân viên */}
            {!user.isAdmin && (
                <div className="lg:col-span-1">
                    <div className="bg-white p-6 rounded-lg shadow">
                         <h3 className="text-xl font-bold mb-4">Tạo Yêu Cầu Nghỉ Phép</h3>
                         <form onSubmit={handleSubmit} className="space-y-4">
                             <div>
                                 <label className="block text-sm font-medium">Người duyệt</label>
                                 <select value={approverId} onChange={e => setApproverId(e.target.value)} className="w-full mt-1 px-3 py-2 border rounded-md bg-white">
                                     {approvers.map(a => <option key={a.id} value={a.id}>{a.name} ({a.id})</option>)}
                                 </select>
                             </div>
                              <div>
                                 <label className="block text-sm font-medium">Ngày bắt đầu</label>
                                 <input type="date" value={startDate} onChange={e => setStartDate(e.target.value)} required className="w-full mt-1 px-3 py-2 border rounded-md" />
                             </div>
                              <div>
                                 <label className="block text-sm font-medium">Ngày kết thúc</label>
                                 <input type="date" value={endDate} onChange={e => setEndDate(e.target.value)} required className="w-full mt-1 px-3 py-2 border rounded-md" />
                             </div>
                              <div>
                                 <label className="block text-sm font-medium">Lý do</label>
                                 <textarea value={reason} onChange={e => setReason(e.target.value)} required rows="3" className="w-full mt-1 px-3 py-2 border rounded-md"></textarea>
                             </div>
                             <button type="submit" className="w-full py-2 px-4 bg-indigo-600 text-white font-semibold rounded-md hover:bg-indigo-700">Gửi Yêu Cầu</button>
                         </form>
                    </div>
                </div>
            )}
            
            {/* CỘT HIỂN THỊ DANH SÁCH */}
            <div className={user.isAdmin ? "lg:col-span-3" : "lg:col-span-2"}>
                <div className="space-y-6">
                    <div>
                        <h3 className="text-xl font-bold mb-3">Yêu Cầu Chờ Duyệt ({pendingRequests.length})</h3>
                        <div className="space-y-4">
                            {pendingRequests.length > 0 ? pendingRequests.map(req => <RequestCard key={req.id} req={req}/>) : <p className="text-gray-500">Không có yêu cầu nào.</p>}
                        </div>
                    </div>
                     <div>
                        <h3 className="text-xl font-bold mb-3">Yêu Cầu Đã Duyệt ({approvedRequests.length})</h3>
                         <div className="space-y-4">
                            {approvedRequests.length > 0 ? approvedRequests.map(req => <RequestCard key={req.id} req={req}/>) : <p className="text-gray-500">Không có yêu cầu nào.</p>}
                        </div>
                    </div>
                     <div>
                        <h3 className="text-xl font-bold mb-3">Yêu Cầu Bị Từ Chối ({rejectedRequests.length})</h3>
                         <div className="space-y-4">
                            {rejectedRequests.length > 0 ? rejectedRequests.map(req => <RequestCard key={req.id} req={req}/>) : <p className="text-gray-500">Không có yêu cầu nào.</p>}
                        </div>
                    </div>
                </div>
            </div>
        </div>
    );
}

