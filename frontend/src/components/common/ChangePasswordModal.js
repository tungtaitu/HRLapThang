/*
 * File: components/common/ChangePasswordModal.js
 * Mô tả: Component modal cho phép người dùng đổi mật khẩu.
 */
import React, { useState } from 'react';
import { apiChangePassword } from '../../api';

export default function ChangePasswordModal({ user, onClose }) {
    const [oldPassword, setOldPassword] = useState('');
    const [newPassword, setNewPassword] = useState('');
    const [confirmPassword, setConfirmPassword] = useState('');
    const [error, setError] = useState('');
    const [isLoading, setIsLoading] = useState(false);

    const handleSubmit = async (e) => {
        e.preventDefault();
        setError('');
        if (newPassword !== confirmPassword) {
            setError('Mật khẩu mới không khớp.');
            return;
        }
        setIsLoading(true);
        try {
            await apiChangePassword(user.id, oldPassword, newPassword);
            alert('Đổi mật khẩu thành công!');
            onClose();
        } catch (err) {
            setError(err.message);
        } finally {
            setIsLoading(false);
        }
    };

    return (
        <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50 p-4">
            <div className="bg-white rounded-lg shadow-xl p-6 w-full max-w-md">
                <h2 className="text-xl font-bold mb-4">Đổi Mật Khẩu</h2>
                <form onSubmit={handleSubmit} className="space-y-4">
                    <div>
                        <label className="block text-sm font-medium text-gray-700">Mật khẩu cũ</label>
                        <input type="password" value={oldPassword} onChange={(e) => setOldPassword(e.target.value)} required className="w-full px-3 py-2 mt-1 border rounded-md" />
                    </div>
                     <div>
                        <label className="block text-sm font-medium text-gray-700">Mật khẩu mới</label>
                        <input type="password" value={newPassword} onChange={(e) => setNewPassword(e.target.value)} required className="w-full px-3 py-2 mt-1 border rounded-md" />
                    </div>
                     <div>
                        <label className="block text-sm font-medium text-gray-700">Xác nhận mật khẩu mới</label>
                        <input type="password" value={confirmPassword} onChange={(e) => setConfirmPassword(e.target.value)} required className="w-full px-3 py-2 mt-1 border rounded-md" />
                    </div>
                    {error && <p className="text-sm text-red-600">{error}</p>}
                    <div className="flex justify-end gap-4 mt-6">
                        <button type="button" onClick={onClose} className="px-4 py-2 bg-gray-200 rounded-md">Hủy</button>
                        <button type="submit" disabled={isLoading} className="px-4 py-2 bg-indigo-600 text-white rounded-md disabled:bg-indigo-400">
                            {isLoading ? 'Đang lưu...' : 'Lưu thay đổi'}
                        </button>
                    </div>
                </form>
            </div>
        </div>
    );
}
