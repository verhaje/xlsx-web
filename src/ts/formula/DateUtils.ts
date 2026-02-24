// DateUtils.ts - Excel serial date conversion utilities

const MS_PER_DAY = 24 * 60 * 60 * 1000;
const EXCEL_EPOCH_UTC = Date.UTC(1899, 11, 31);
const EXCEL_1900_MAR_1_UTC = Date.UTC(1900, 2, 1);

/**
 * Utility class for converting between Excel serial dates and calendar dates.
 * Handles the Excel 1900 leap-year bug (serial 60 = Feb 29, 1900).
 */
export class DateUtils {
  /**
   * Normalize a 2-digit year to a 4-digit year (0–1899 maps to 1900–3799).
   */
  static normalizeExcelYear(year: number): number {
    if (year >= 0 && year <= 1899) return 1900 + year;
    return year;
  }

  /**
   * Convert a year/month/day to an Excel serial number.
   */
  static ymdToExcelSerial(year: number, month: number, day: number): number {
    if (!Number.isFinite(year) || !Number.isFinite(month) || !Number.isFinite(day)) return Number.NaN;
    const y = DateUtils.normalizeExcelYear(Math.trunc(year));
    const m = Math.trunc(month);
    const d = Math.trunc(day);
    if (y < 0) return Number.NaN;

    // Excel treats 1900-02-29 as a valid date (serial 60)
    if (y === 1900 && m === 2 && d === 29) return 60;

    const utc = Date.UTC(y, m - 1, d);
    if (!Number.isFinite(utc)) return Number.NaN;

    let days = Math.floor((utc - EXCEL_EPOCH_UTC) / MS_PER_DAY);
    if (utc >= EXCEL_1900_MAR_1_UTC) days += 1;
    return days;
  }

  /**
   * Convert an Excel serial number to { year, month, day }.
   */
  static excelSerialToParts(serial: number): { year: number; month: number; day: number } | null {
    if (!Number.isFinite(serial)) return null;
    const s = Math.floor(serial);
    if (s < 0) return null;
    if (s === 60) return { year: 1900, month: 2, day: 29 };

    let days = s;
    if (s > 60) days -= 1;

    const utc = EXCEL_EPOCH_UTC + days * MS_PER_DAY;
    const date = new Date(utc);
    return {
      year: date.getUTCFullYear(),
      month: date.getUTCMonth() + 1,
      day: date.getUTCDate(),
    };
  }

  /**
   * Get today's date as an Excel serial number.
   */
  static todayToExcelSerial(): number {
    const now = new Date();
    return DateUtils.ymdToExcelSerial(now.getFullYear(), now.getMonth() + 1, now.getDate());
  }
}
