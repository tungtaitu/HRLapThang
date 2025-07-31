/*
 * File: pages/LoginPage.js
 * Mô tả: Component trang đăng nhập, chứa form và xử lý logic đăng nhập.
 */
import React, { useState } from 'react';
import { apiLogin } from '../api';

// --- Component LoginForm (tách từ App.js) ---
function LoginForm({ onLogin, error, isLoading }) {
    const [empid, setEmpid] = useState('');
    const [password, setPassword] = useState('');
    const [showPassword, setShowPassword] = useState(false);

    const handleSubmit = (e) => {
        e.preventDefault();
        onLogin(empid, password);
    };

    return (
        <div className="flex items-center justify-center min-h-screen bg-gray-100 px-4">
            <div className="w-full max-w-md p-8 space-y-4 bg-white rounded-lg shadow-md">
                <div className="flex flex-col items-center justify-center mb-6">
                    <img src="/logo.png" alt="Logo Công ty Lập Thắng" className="h-40 mb-4" />
                    <h1 className="text-xl font-bold text-center  text-indigo-600">CÔNG TY TNHH LẬP THẮNG</h1>
                    <h2 className="text-3xl font-bold text-center text-gray-700">Hệ Thống Nhân Sự</h2>
                </div>
                <h2 className="text-2xl font-bold text-center text-gray-800">Đăng nhập</h2>

                <form className="space-y-6" onSubmit={handleSubmit}>
                    <div>
                        <label htmlFor="empid" className="text-sm font-medium text-gray-700">Tên đăng nhập</label>
                        <input id="empid" type="text" value={empid} onChange={(e) => setEmpid(e.target.value)} required className="w-full px-3 py-2 mt-1 border border-gray-300 rounded-md shadow-sm focus:ring-indigo-500 focus:border-indigo-500" placeholder="Nhập mã nhân viên hoặc tài khoản admin" />
                    </div>
                    <div>
                        <label htmlFor="password" className="text-sm font-medium text-gray-700">Mật khẩu (Ngày sinh)</label>
                        <div className="relative mt-1">
                            <input id="password" type={showPassword ? 'text' : 'password'} value={password} onChange={(e) => setPassword(e.target.value)} required className="w-full px-3 py-2 pr-10 border border-gray-300 rounded-md shadow-sm focus:ring-indigo-500 focus:border-indigo-500" placeholder="Nhập theo định dạng ddmmyyyy" />
                            <button type="button" onClick={() => setShowPassword(!showPassword)} className="absolute inset-y-0 right-0 flex items-center px-3 text-gray-400 hover:text-gray-600" aria-label={showPassword ? "Ẩn mật khẩu" : "Hiện mật khẩu"}>
                                {showPassword ? (
                                    <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                                      <path strokeLinecap="round" strokeLinejoin="round" d="M13.875 18.825A10.05 10.05 0 0112 19c-4.478 0-8.268-2.943-9.543-7a9.97 9.97 0 011.563-3.029m5.858.908a3 3 0 114.243 4.243M9.878 9.878l4.242 4.242" />
                                    </svg>
                                ) : (
                                    <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                                      <path strokeLinecap="round" strokeLinejoin="round" d="M15 12a3 3 0 11-6 0 3 3 0 016 0z" />
                                      <path strokeLinecap="round" strokeLinejoin="round" d="M2.458 12C3.732 7.943 7.523 5 12 5c4.478 0 8.268 2.943 9.542 7-1.274 4.057-5.27 7-9.542 7S3.732 16.057 2.458 12z" />
                                    </svg>
                                )}
                            </button>
                        </div>
                    </div>

                    {error && <p className="text-sm text-center text-red-600">{error}</p>}
                    <div>
                        <button type="submit" disabled={isLoading} className="w-full px-4 py-2 font-medium text-white bg-indigo-600 rounded-md hover:bg-indigo-700 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-indigo-500 disabled:bg-gray-400">
                            {isLoading ? 'Đang xử lý...' : 'Đăng nhập'}
                        </button>
                    </div>
                </form>
            </div>
        </div>
    );
}


// --- Component Trang Đăng Nhập ---
export default function LoginPage({ onLoginSuccess }) {
    const [loginError, setLoginError] = useState('');
    const [isLoading, setIsLoading] = useState(false);

    const handleLogin = async (empid, password) => {
        setIsLoading(true);
        setLoginError('');
        try {
            const loggedInUser = await apiLogin(empid, password);
            onLoginSuccess(loggedInUser); // Gọi hàm callback từ App.js khi đăng nhập thành công
        } catch (error) {
            setLoginError(error.message);
        } finally {
            setIsLoading(false);
        }
    };

    return (
        <LoginForm
            onLogin={handleLogin}
            error={loginError}
            isLoading={isLoading}
        />
    );
}
