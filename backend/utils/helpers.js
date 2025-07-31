/*
 * File: utils/helpers.js
 * Mô tả: Chứa các hàm tiện ích chung khác.
 */

const parseDate = (ddmmyyyy) => {
    if (!ddmmyyyy || ddmmyyyy.length !== 8) return null;
    const day = parseInt(ddmmyyyy.substring(0, 2));
    const month = parseInt(ddmmyyyy.substring(2, 4)) - 1;
    const year = parseInt(ddmmyyyy.substring(4, 8));
    const date = new Date(Date.UTC(year, month, day));
    if (isNaN(date.getTime()) || date.getUTCFullYear() !== year || date.getUTCMonth() !== month || date.getUTCDate() !== day) {
        return null;
    }
    return date;
};

module.exports = {
    parseDate
};
