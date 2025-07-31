/*
 * File: services/leave.service.js
 * Mô tả: Chứa logic nghiệp vụ phức tạp, ví dụ như tính toán phép năm.
 */
const sql = require('mssql');
const { jsonDataState } = require('./json.service');

async function getLeaveSummary(connectionPool, userId, year) {
    const selectedYear = parseInt(year);
    const today = new Date();
    const currentYear = today.getFullYear();
    let summary = { total: 0, used: 0, remaining: 0, isCurrentYear: false };

    const employeeConfig = jsonDataState.leaveData.find(emp => emp.MSNV === userId);

    if (!employeeConfig) {
        summary.remaining = 0;
        summary.total = 0;
        summary.isCurrentYear = (selectedYear === currentYear);
        try {
            const totalUsedInYearResult = await connectionPool.request()
                .input('userid_param', sql.VarChar, userId)
                .input('year_param', sql.Int, year)
                .query`SELECT SUM(ISNULL(HHour, 0)) as TotalUsed FROM EMPHOLIDAY WHERE empid = @userid_param AND YEAR(DateUP) = @year_param AND JiaType = 'E'`;
            summary.used = totalUsedInYearResult.recordset[0]?.TotalUsed || 0;
        } catch (dbError) {
            console.error(`Lỗi khi truy vấn số phép đã dùng cho NV ${userId} (không có trong file config):`, dbError);
            summary.used = 0;
        }
        return summary;
    }
    
    const isForeigner = employeeConfig.isForeigner || false;
    const monthlyLeaveHours = isForeigner ? 16 : 8;

    if (selectedYear === currentYear) {
        summary.isCurrentYear = true;

        const configMonth = parseInt(employeeConfig['Month']) || 1;
        const carriedOverHours = parseFloat(employeeConfig['PHEP']) || 0;
        const currentMonth = today.getMonth() + 1; // 1-12
        let entitledThisYear = 0;
        let entitledMonthsCount = 0;

        if (configMonth === 12) {
             entitledMonthsCount = currentMonth;
        } else {
            if(currentMonth > configMonth) {
                entitledMonthsCount = currentMonth - configMonth;
            }
        }
        
        entitledThisYear = entitledMonthsCount * monthlyLeaveHours;

        const firstUsageMonthToConsider = (configMonth === 12) ? 1 : configMonth + 1;

        const usedLeaveSinceEntitlementResult = await connectionPool.request()
            .input('userid_param', sql.VarChar, userId)
            .input('year_param', sql.Int, year)
            .input('start_month_param', sql.Int, firstUsageMonthToConsider)
            .query`
                SELECT SUM(ISNULL(HHour, 0)) as TotalUsed
                FROM EMPHOLIDAY
                WHERE empid = @userid_param
                  AND YEAR(DateUP) = @year_param
                  AND MONTH(DateUP) >= @start_month_param
                  AND JiaType = 'E'
            `;
        const usedHoursSinceEntitlement = usedLeaveSinceEntitlementResult.recordset[0]?.TotalUsed || 0;

        const totalUsedInYearResult = await connectionPool.request()
            .input('userid_param', sql.VarChar, userId)
            .input('year_param', sql.Int, year)
            .query`SELECT SUM(ISNULL(HHour, 0)) as TotalUsed FROM EMPHOLIDAY WHERE empid = @userid_param AND YEAR(DateUP) = @year_param AND JiaType = 'E'`;
        const totalUsedForDisplay = totalUsedInYearResult.recordset[0]?.TotalUsed || 0;

        summary.total = carriedOverHours + entitledThisYear;
        summary.used = totalUsedForDisplay;
        summary.remaining = summary.total - usedHoursSinceEntitlement;

    } else {
        const totalUsedInYearResult = await connectionPool.request()
            .input('userid_param', sql.VarChar, userId)
            .input('year_param', sql.Int, year)
            .query`SELECT SUM(ISNULL(HHour, 0)) as TotalUsed FROM EMPHOLIDAY WHERE empid = @userid_param AND YEAR(DateUP) = @year_param AND JiaType = 'E'`;

        summary.used = totalUsedInYearResult.recordset[0]?.TotalUsed || 0;
        summary.total = 0;
        summary.remaining = 0;
    }

    return summary;
}

module.exports = {
    getLeaveSummary
};
