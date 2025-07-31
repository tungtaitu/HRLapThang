/*
 * File: App.js
 * Mô tả: Component gốc của ứng dụng React.
 * Chịu trách nhiệm quản lý trạng thái xác thực và điều hướng người dùng
 * đến trang đăng nhập hoặc trang dashboard tương ứng.
 * Cập nhật: Tích hợp BrowserRouter để quản lý định tuyến.
 */
import React, { useState, useEffect } from 'react';
import { BrowserRouter, Routes, Route } from 'react-router-dom'; // Thêm import cho BrowserRouter
import AdminDashboard from './pages/AdminDashboard';
import EmployeeDashboard from './pages/EmployeeDashboard';
import LoginPage from './pages/LoginPage';
import { apiCheckSession, apiLogout } from './api';

export default function App() {
    const [user, setUser] = useState(null);
    const [isLoading, setIsLoading] = useState(true);

    // Kiểm tra session khi ứng dụng được tải lần đầu
    useEffect(() => {
        const checkUserSession = async () => {
            try {
                const sessionUser = await apiCheckSession();
                if (sessionUser) {
                    setUser(sessionUser);
                }
            } catch (error) {
                // Không có session hợp lệ, không cần làm gì
            } finally {
                setIsLoading(false);
            }
        };
        checkUserSession();
    }, []);

    const handleLoginSuccess = (loggedInUser) => {
        setUser(loggedInUser);
    };

    const handleLogout = async () => {
        setIsLoading(true);
        try {
            await apiLogout();
        } catch (error) {
            console.error("Lỗi khi đăng xuất:", error);
        } finally {
            setUser(null);
            setIsLoading(false);
        }
    };

    // Hiển thị màn hình tải trong khi kiểm tra session
    if (isLoading) {
        return <div className="flex justify-center items-center min-h-screen"><p>Đang tải ứng dụng...</p></div>;
    }

    // Bao bọc toàn bộ ứng dụng trong BrowserRouter để kích hoạt định tuyến
    return (
        <BrowserRouter>
            {/* Logic hiển thị component không đổi, nhưng giờ đã nằm trong môi trường router */}
            {!user ? (
                <LoginPage onLoginSuccess={handleLoginSuccess} />
            ) : user.isAdmin ? (
                <AdminDashboard user={user} onLogout={handleLogout} />
            ) : (
                <EmployeeDashboard user={user} onLogout={handleLogout} />
            )}
        </BrowserRouter>
    );
}
