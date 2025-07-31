/*
 * File: utils/calculate.js
 * Mô tả: Chứa các hàm tính toán có thể tái sử dụng.
 */

function calculateWorkDuration(indat) {
    if (!indat) return 'Không rõ';
    const startDate = new Date(indat);
    const endDate = new Date();
    let years = endDate.getFullYear() - startDate.getFullYear();
    let months = endDate.getMonth() - startDate.getMonth();
    let days = endDate.getDate() - startDate.getDate();
    if (days < 0) {
        months--;
        const prevMonthLastDay = new Date(endDate.getFullYear(), endDate.getMonth(), 0).getDate();
        days += prevMonthLastDay;
    }
    if (months < 0) {
        years--;
        months += 12;
    }
    return `${years} năm ${months} tháng ${days} ngày`;
}

function calculateLeaveHours(startTimeStr, endTimeStr) {
    if (!startTimeStr || !endTimeStr || startTimeStr.length < 4 || endTimeStr.length < 4) return 0;

    const startH = parseInt(startTimeStr.substring(0, 2));
    const startM = parseInt(startTimeStr.substring(2, 4));
    const endH = parseInt(endTimeStr.substring(0, 2));
    const endM = parseInt(endTimeStr.substring(2, 4));

    if (isNaN(startH) || isNaN(startM) || isNaN(endH) || isNaN(endM)) return 0;

    const start = new Date(1970, 0, 1, startH, startM);
    const end = new Date(1970, 0, 1, endH, endM);
    if (end <= start) return 0;

    const morningStart = new Date(1970, 0, 1, 8, 0);
    const morningEnd = new Date(1970, 0, 1, 12, 0);
    const afternoonStart = new Date(1970, 0, 1, 13, 0);
    const afternoonEnd = new Date(1970, 0, 1, 17, 0);
    let totalMs = 0;

    const morningOverlapStart = Math.max(start, morningStart);
    const morningOverlapEnd = Math.min(end, morningEnd);
    if (morningOverlapEnd > morningOverlapStart) totalMs += morningOverlapEnd - morningOverlapStart;

    const afternoonOverlapStart = Math.max(start, afternoonStart);
    const afternoonOverlapEnd = Math.min(end, afternoonEnd);
    if (afternoonOverlapEnd > afternoonOverlapStart) totalMs += afternoonOverlapEnd - afternoonOverlapStart;

    return Math.round((totalMs / 3600000) * 10) / 10;
}

module.exports = {
    calculateWorkDuration,
    calculateLeaveHours
};
