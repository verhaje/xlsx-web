// date.ts - Date and time functions

import { ERRORS } from '../Errors';
import { DateUtils } from '../DateUtils';
import type { BuiltinFunction } from '../../types';

export function getDateFunctions(): Record<string, BuiltinFunction> {
  return {
    DATE: async (args) => {
      if (args.length < 3) return ERRORS.VALUE;
      const year = Number(args[0]);
      const month = Number(args[1]);
      const day = Number(args[2]);
      if (!Number.isFinite(year) || !Number.isFinite(month) || !Number.isFinite(day)) return ERRORS.VALUE;
      const serial = DateUtils.ymdToExcelSerial(year, month, day);
      return Number.isFinite(serial) ? serial : ERRORS.VALUE;
    },
    YEAR: async (args) => {
      const v = args[0];
      if (v instanceof Date) return v.getUTCFullYear();
      const n = Number(v);
      if (!Number.isFinite(n)) return ERRORS.VALUE;
      const parts = DateUtils.excelSerialToParts(n);
      return parts ? parts.year : ERRORS.VALUE;
    },
    MONTH: async (args) => {
      const v = args[0];
      if (v instanceof Date) return v.getUTCMonth() + 1;
      const n = Number(v);
      if (!Number.isFinite(n)) return ERRORS.VALUE;
      const parts = DateUtils.excelSerialToParts(n);
      return parts ? parts.month : ERRORS.VALUE;
    },
    DAY: async (args) => {
      const v = args[0];
      if (v instanceof Date) return v.getUTCDate();
      const n = Number(v);
      if (!Number.isFinite(n)) return ERRORS.VALUE;
      const parts = DateUtils.excelSerialToParts(n);
      return parts ? parts.day : ERRORS.VALUE;
    },
    TODAY: async () => DateUtils.todayToExcelSerial(),
    NOW: async () => {
      const now = new Date();
      const serial = DateUtils.ymdToExcelSerial(now.getFullYear(), now.getMonth() + 1, now.getDate());
      const timeFrac = (now.getHours() * 3600 + now.getMinutes() * 60 + now.getSeconds()) / 86400;
      return serial + timeFrac;
    },
    TIME: async (args) => {
      const h = Number(args[0]) || 0;
      const m = Number(args[1]) || 0;
      const s = Number(args[2]) || 0;
      const total = h * 3600 + m * 60 + s;
      if (total < 0) return ERRORS.NUM;
      return (total % 86400) / 86400;
    },
    HOUR: async (args) => {
      const v = Number(args[0]);
      if (!Number.isFinite(v)) return ERRORS.VALUE;
      const frac = v - Math.floor(v);
      return Math.floor(frac * 24) % 24;
    },
    MINUTE: async (args) => {
      const v = Number(args[0]);
      if (!Number.isFinite(v)) return ERRORS.VALUE;
      const frac = v - Math.floor(v);
      return Math.floor(frac * 1440) % 60;
    },
    SECOND: async (args) => {
      const v = Number(args[0]);
      if (!Number.isFinite(v)) return ERRORS.VALUE;
      const frac = v - Math.floor(v);
      return Math.floor(frac * 86400) % 60;
    },
    DATEDIF: async (args) => {
      const start = Number(args[0]);
      const end = Number(args[1]);
      const unit = String(args[2] ?? '').toUpperCase();
      if (!Number.isFinite(start) || !Number.isFinite(end)) return ERRORS.VALUE;
      if (start > end) return ERRORS.NUM;
      const sp = DateUtils.excelSerialToParts(start);
      const ep = DateUtils.excelSerialToParts(end);
      if (!sp || !ep) return ERRORS.VALUE;
      switch (unit) {
        case 'Y': {
          let years = ep.year - sp.year;
          if (ep.month < sp.month || (ep.month === sp.month && ep.day < sp.day)) years--;
          return years;
        }
        case 'M': {
          let months = (ep.year - sp.year) * 12 + (ep.month - sp.month);
          if (ep.day < sp.day) months--;
          return months;
        }
        case 'D': return Math.floor(end) - Math.floor(start);
        case 'MD': {
          let d = ep.day - sp.day;
          if (d < 0) {
            const prevMonth = new Date(Date.UTC(ep.year, ep.month - 1, 0));
            d = prevMonth.getUTCDate() - sp.day + ep.day;
          }
          return d;
        }
        case 'YM': {
          let m = ep.month - sp.month;
          if (m < 0) m += 12;
          if (ep.day < sp.day) m--;
          if (m < 0) m += 12;
          return m;
        }
        case 'YD': {
          let start2 = DateUtils.ymdToExcelSerial(ep.year, sp.month, sp.day);
          if (start2 > end) start2 = DateUtils.ymdToExcelSerial(ep.year - 1, sp.month, sp.day);
          return Math.floor(end) - Math.floor(start2);
        }
        default: return ERRORS.NUM;
      }
    },
    DATEVALUE: async (args) => {
      const s = String(args[0] ?? '').trim();
      const d = new Date(s);
      if (isNaN(d.getTime())) return ERRORS.VALUE;
      return DateUtils.ymdToExcelSerial(d.getFullYear(), d.getMonth() + 1, d.getDate());
    },
    TIMEVALUE: async (args) => {
      const s = String(args[0] ?? '').trim();
      const m = /^(\d{1,2}):(\d{2})(?::(\d{2}))?(?:\s*(AM|PM))?$/i.exec(s);
      if (!m) {
        const d = new Date(`1970-01-01 ${s}`);
        if (isNaN(d.getTime())) return ERRORS.VALUE;
        return (d.getHours() * 3600 + d.getMinutes() * 60 + d.getSeconds()) / 86400;
      }
      let h = Number(m[1]);
      const min = Number(m[2]);
      const sec = m[3] ? Number(m[3]) : 0;
      if (m[4]) {
        const pm = m[4].toUpperCase() === 'PM';
        if (pm && h < 12) h += 12;
        if (!pm && h === 12) h = 0;
      }
      return (h * 3600 + min * 60 + sec) / 86400;
    },
    DAYS: async (args) => {
      const end = Number(args[0]);
      const start = Number(args[1]);
      if (!Number.isFinite(end) || !Number.isFinite(start)) return ERRORS.VALUE;
      return Math.floor(end) - Math.floor(start);
    },
    DAYS360: async (args) => {
      const startSerial = Number(args[0]);
      const endSerial = Number(args[1]);
      const method = args[2] ? true : false;
      if (!Number.isFinite(startSerial) || !Number.isFinite(endSerial)) return ERRORS.VALUE;
      const sp = DateUtils.excelSerialToParts(startSerial);
      const ep = DateUtils.excelSerialToParts(endSerial);
      if (!sp || !ep) return ERRORS.VALUE;
      let sd = sp.day, ed = ep.day;
      if (method) {
        if (sd > 30) sd = 30;
        if (ed > 30) ed = 30;
      } else {
        if (sd === 31) sd = 30;
        if (ed === 31 && sd >= 30) ed = 30;
      }
      return (ep.year - sp.year) * 360 + (ep.month - sp.month) * 30 + (ed - sd);
    },
    EDATE: async (args) => {
      const startSerial = Number(args[0]);
      const months = Math.trunc(Number(args[1]));
      if (!Number.isFinite(startSerial) || !Number.isFinite(months)) return ERRORS.VALUE;
      const sp = DateUtils.excelSerialToParts(startSerial);
      if (!sp) return ERRORS.VALUE;
      let newMonth = sp.month + months;
      let newYear = sp.year;
      while (newMonth > 12) { newMonth -= 12; newYear++; }
      while (newMonth < 1) { newMonth += 12; newYear--; }
      const maxDay = new Date(Date.UTC(newYear, newMonth, 0)).getUTCDate();
      const newDay = Math.min(sp.day, maxDay);
      return DateUtils.ymdToExcelSerial(newYear, newMonth, newDay);
    },
    EOMONTH: async (args) => {
      const startSerial = Number(args[0]);
      const months = Math.trunc(Number(args[1]));
      if (!Number.isFinite(startSerial) || !Number.isFinite(months)) return ERRORS.VALUE;
      const sp = DateUtils.excelSerialToParts(startSerial);
      if (!sp) return ERRORS.VALUE;
      let newMonth = sp.month + months;
      let newYear = sp.year;
      while (newMonth > 12) { newMonth -= 12; newYear++; }
      while (newMonth < 1) { newMonth += 12; newYear--; }
      const lastDay = new Date(Date.UTC(newYear, newMonth, 0)).getUTCDate();
      return DateUtils.ymdToExcelSerial(newYear, newMonth, lastDay);
    },
    WEEKDAY: async (args) => {
      const serial = Number(args[0]);
      const returnType = args[1] !== undefined ? Number(args[1]) : 1;
      if (!Number.isFinite(serial)) return ERRORS.VALUE;
      const parts = DateUtils.excelSerialToParts(serial);
      if (!parts) return ERRORS.VALUE;
      const d = new Date(Date.UTC(parts.year, parts.month - 1, parts.day));
      const jsDay = d.getUTCDay();
      switch (returnType) {
        case 1: return jsDay + 1;
        case 2: return jsDay === 0 ? 7 : jsDay;
        case 3: return jsDay === 0 ? 6 : jsDay - 1;
        default: return ERRORS.NUM;
      }
    },
    WEEKNUM: async (args) => {
      const serial = Number(args[0]);
      const returnType = args[1] !== undefined ? Number(args[1]) : 1;
      if (!Number.isFinite(serial)) return ERRORS.VALUE;
      const parts = DateUtils.excelSerialToParts(serial);
      if (!parts) return ERRORS.VALUE;
      const d = new Date(Date.UTC(parts.year, parts.month - 1, parts.day));
      const jan1 = new Date(Date.UTC(parts.year, 0, 1));
      const startDay = returnType === 2 ? 1 : 0;
      const jan1Day = jan1.getUTCDay();
      const adjust = (jan1Day - startDay + 7) % 7;
      const dayOfYear = Math.floor((d.getTime() - jan1.getTime()) / 86400000);
      return Math.floor((dayOfYear + adjust) / 7) + 1;
    },
    ISOWEEKNUM: async (args) => {
      const serial = Number(args[0]);
      if (!Number.isFinite(serial)) return ERRORS.VALUE;
      const parts = DateUtils.excelSerialToParts(serial);
      if (!parts) return ERRORS.VALUE;
      const d = new Date(Date.UTC(parts.year, parts.month - 1, parts.day));
      d.setUTCDate(d.getUTCDate() + 3 - ((d.getUTCDay() + 6) % 7));
      const week1 = new Date(Date.UTC(d.getUTCFullYear(), 0, 4));
      return 1 + Math.round(((d.getTime() - week1.getTime()) / 86400000 - 3 + ((week1.getUTCDay() + 6) % 7)) / 7);
    },
    NETWORKDAYS: async (args) => {
      const startSerial = Number(args[0]);
      const endSerial = Number(args[1]);
      const holidays = Array.isArray(args[2]) ? args[2].map((v: any) => Math.floor(Number(v))).filter((n: number) => Number.isFinite(n)) : [];
      if (!Number.isFinite(startSerial) || !Number.isFinite(endSerial)) return ERRORS.VALUE;
      const holidaySet = new Set(holidays);
      let start = Math.floor(startSerial);
      let end = Math.floor(endSerial);
      const sign = start <= end ? 1 : -1;
      if (sign === -1) [start, end] = [end, start];
      let count = 0;
      for (let s = start; s <= end; s++) {
        const parts = DateUtils.excelSerialToParts(s);
        if (!parts) continue;
        const d = new Date(Date.UTC(parts.year, parts.month - 1, parts.day));
        const day = d.getUTCDay();
        if (day !== 0 && day !== 6 && !holidaySet.has(s)) count++;
      }
      return count * sign;
    },
    WORKDAY: async (args) => {
      const startSerial = Math.floor(Number(args[0]));
      let days = Math.trunc(Number(args[1]));
      const holidays = Array.isArray(args[2]) ? args[2].map((v: any) => Math.floor(Number(v))).filter((n: number) => Number.isFinite(n)) : [];
      if (!Number.isFinite(startSerial) || !Number.isFinite(days)) return ERRORS.VALUE;
      const holidaySet = new Set(holidays);
      let current = startSerial;
      const step = days > 0 ? 1 : -1;
      days = Math.abs(days);
      while (days > 0) {
        current += step;
        const parts = DateUtils.excelSerialToParts(current);
        if (!parts) return ERRORS.VALUE;
        const d = new Date(Date.UTC(parts.year, parts.month - 1, parts.day));
        const day = d.getUTCDay();
        if (day !== 0 && day !== 6 && !holidaySet.has(current)) days--;
      }
      return current;
    },
    YEARFRAC: async (args) => {
      const startSerial = Number(args[0]);
      const endSerial = Number(args[1]);
      const basis = args[2] !== undefined ? Number(args[2]) : 0;
      if (!Number.isFinite(startSerial) || !Number.isFinite(endSerial)) return ERRORS.VALUE;
      const sd = Math.min(startSerial, endSerial);
      const ed = Math.max(startSerial, endSerial);
      const sp = DateUtils.excelSerialToParts(sd);
      const ep = DateUtils.excelSerialToParts(ed);
      if (!sp || !ep) return ERRORS.VALUE;
      const diffDays = Math.floor(ed) - Math.floor(sd);
      switch (basis) {
        case 0: {
          let d1 = sp.day, d2 = ep.day;
          if (d1 === 31) d1 = 30;
          if (d2 === 31 && d1 >= 30) d2 = 30;
          const days360 = (ep.year - sp.year) * 360 + (ep.month - sp.month) * 30 + (d2 - d1);
          return days360 / 360;
        }
        case 1: {
          const sy = sp.year;
          const ey = ep.year;
          if (sy === ey) {
            const yearLen = (new Date(Date.UTC(sy, 11, 31)).getTime() - new Date(Date.UTC(sy, 0, 1)).getTime()) / 86400000 + 1;
            return diffDays / yearLen;
          }
          const totalYears = ey - sy + 1;
          let totalDaysInYears = 0;
          for (let y = sy; y <= ey; y++) {
            totalDaysInYears += (new Date(Date.UTC(y, 11, 31)).getTime() - new Date(Date.UTC(y, 0, 1)).getTime()) / 86400000 + 1;
          }
          return diffDays / (totalDaysInYears / totalYears);
        }
        case 2: return diffDays / 360;
        case 3: return diffDays / 365;
        case 4: {
          let d1 = Math.min(sp.day, 30);
          let d2 = Math.min(ep.day, 30);
          const days360 = (ep.year - sp.year) * 360 + (ep.month - sp.month) * 30 + (d2 - d1);
          return days360 / 360;
        }
        default: return ERRORS.NUM;
      }
    },
    'WORKDAY.INTL': async (args) => {
      const startSerial = Math.trunc(Number(args[0]));
      const dayCount = Math.trunc(Number(args[1]));
      const weekend = args[2] !== undefined ? args[2] : 1;
      if (!Number.isFinite(startSerial) || !Number.isFinite(dayCount)) return ERRORS.VALUE;

      let isWeekendDay: (dow: number) => boolean;
      if (typeof weekend === 'string' && weekend.length === 7) {
        isWeekendDay = (dow: number) => weekend[(dow + 6) % 7] === '1';
      } else {
        const code = Number(weekend);
        const wkMap: Record<number, number[]> = {
          1: [0, 6], 2: [0, 1], 3: [1, 2], 4: [2, 3], 5: [3, 4], 6: [4, 5], 7: [5, 6],
          11: [0], 12: [1], 13: [2], 14: [3], 15: [4], 16: [5], 17: [6],
        };
        const wkDays = wkMap[code] || [0, 6];
        isWeekendDay = (dow: number) => wkDays.includes(dow);
      }

      const holidays = new Set<number>();
      if (args[3]) {
        const hArr = Array.isArray(args[3]) ? args[3] : [args[3]];
        for (const h of hArr) { const hn = Math.trunc(Number(h)); if (Number.isFinite(hn)) holidays.add(hn); }
      }

      let current = startSerial;
      let remaining = Math.abs(dayCount);
      const step = dayCount >= 0 ? 1 : -1;
      while (remaining > 0) {
        current += step;
        const parts = DateUtils.excelSerialToParts(current);
        if (!parts) return ERRORS.VALUE;
        const dt = new Date(Date.UTC(parts.year, parts.month - 1, parts.day));
        const dow = dt.getUTCDay();
        if (!isWeekendDay(dow) && !holidays.has(current)) remaining--;
      }
      return current;
    },
    'NETWORKDAYS.INTL': async (args) => {
      const startSerial = Math.trunc(Number(args[0]));
      const endSerial = Math.trunc(Number(args[1]));
      const weekend = args[2] !== undefined ? args[2] : 1;
      if (!Number.isFinite(startSerial) || !Number.isFinite(endSerial)) return ERRORS.VALUE;

      let isWeekendDay: (dow: number) => boolean;
      if (typeof weekend === 'string' && weekend.length === 7) {
        isWeekendDay = (dow: number) => weekend[(dow + 6) % 7] === '1';
      } else {
        const code = Number(weekend);
        const wkMap: Record<number, number[]> = {
          1: [0, 6], 2: [0, 1], 3: [1, 2], 4: [2, 3], 5: [3, 4], 6: [4, 5], 7: [5, 6],
          11: [0], 12: [1], 13: [2], 14: [3], 15: [4], 16: [5], 17: [6],
        };
        const wkDays = wkMap[code] || [0, 6];
        isWeekendDay = (dow: number) => wkDays.includes(dow);
      }

      const holidays = new Set<number>();
      if (args[3]) {
        const hArr = Array.isArray(args[3]) ? args[3] : [args[3]];
        for (const h of hArr) { const hn = Math.trunc(Number(h)); if (Number.isFinite(hn)) holidays.add(hn); }
      }

      const start = Math.min(startSerial, endSerial);
      const end = Math.max(startSerial, endSerial);
      const sign = startSerial <= endSerial ? 1 : -1;
      let count = 0;
      for (let s = start; s <= end; s++) {
        const parts = DateUtils.excelSerialToParts(s);
        if (!parts) continue;
        const dt = new Date(Date.UTC(parts.year, parts.month - 1, parts.day));
        const dow = dt.getUTCDay();
        if (!isWeekendDay(dow) && !holidays.has(s)) count++;
      }
      return count * sign;
    },
  };
}
