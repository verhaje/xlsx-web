// FormulaAutocomplete.ts - Autocomplete dropdown for formula functions
//
// Shows a dropdown with matching Excel function names as the user types
// in the formula bar. Supports keyboard navigation and selection.

import { BuiltinFunctions } from '../formula/BuiltinFunctions';

/**
 * A single autocomplete suggestion.
 */
export interface AutocompleteSuggestion {
  /** Function name (e.g. "SUM") */
  name: string;
  /** Brief description */
  description: string;
  /** Usage hint (e.g. "SUM(number1, [number2], ...)") */
  syntax: string;
}

interface ActiveFunctionContext {
  name: string;
  parameterIndex: number;
}

/**
 * All known function metadata for autocomplete.
 */
const FUNCTION_METADATA: Record<string, { description: string; syntax: string }> = {
  SUM: { description: 'Adds all numbers in a range', syntax: 'SUM(number1, [number2], ...)' },
  AVERAGE: { description: 'Returns the average of arguments', syntax: 'AVERAGE(number1, [number2], ...)' },
  MIN: { description: 'Returns the smallest value', syntax: 'MIN(number1, [number2], ...)' },
  MAX: { description: 'Returns the largest value', syntax: 'MAX(number1, [number2], ...)' },
  COUNT: { description: 'Counts cells that contain numbers', syntax: 'COUNT(value1, [value2], ...)' },
  COUNTA: { description: 'Counts non-empty cells', syntax: 'COUNTA(value1, [value2], ...)' },
  COUNTIFS: { description: 'Counts cells that match criteria', syntax: 'COUNTIFS(range, criteria, ...)' },
  SUMIF: { description: 'Sums cells that match a criterion', syntax: 'SUMIF(range, criteria, [sum_range])' },
  SUMIFS: { description: 'Sums cells that match multiple criteria', syntax: 'SUMIFS(sum_range, range1, criteria1, ...)' },
  AVERAGEIF: { description: 'Averages cells that match a criterion', syntax: 'AVERAGEIF(range, criteria, [avg_range])' },
  AVERAGEIFS: { description: 'Averages cells matching multiple criteria', syntax: 'AVERAGEIFS(avg_range, range1, criteria1, ...)' },
  IF: { description: 'Returns one value if true, another if false', syntax: 'IF(condition, value_if_true, [value_if_false])' },
  AND: { description: 'Returns TRUE if all arguments are true', syntax: 'AND(logical1, [logical2], ...)' },
  OR: { description: 'Returns TRUE if any argument is true', syntax: 'OR(logical1, [logical2], ...)' },
  NOT: { description: 'Reverses the logic of its argument', syntax: 'NOT(logical)' },
  TRUE: { description: 'Returns the logical value TRUE', syntax: 'TRUE()' },
  FALSE: { description: 'Returns the logical value FALSE', syntax: 'FALSE()' },
  IFERROR: { description: 'Returns value if not error, otherwise alternative', syntax: 'IFERROR(value, value_if_error)' },
  ISERROR: { description: 'Returns TRUE if value is an error', syntax: 'ISERROR(value)' },
  ISBLANK: { description: 'Returns TRUE if cell is empty', syntax: 'ISBLANK(value)' },
  ISNUMBER: { description: 'Returns TRUE if value is a number', syntax: 'ISNUMBER(value)' },
  ISTEXT: { description: 'Returns TRUE if value is text', syntax: 'ISTEXT(value)' },
  ABS: { description: 'Returns the absolute value', syntax: 'ABS(number)' },
  ROUND: { description: 'Rounds a number to specified digits', syntax: 'ROUND(number, num_digits)' },
  SQRT: { description: 'Returns the square root', syntax: 'SQRT(number)' },
  LN: { description: 'Returns the natural logarithm', syntax: 'LN(number)' },
  POWER: { description: 'Returns number raised to a power', syntax: 'POWER(number, power)' },
  MOD: { description: 'Returns the remainder from division', syntax: 'MOD(number, divisor)' },
  CONCAT: { description: 'Joins text strings', syntax: 'CONCAT(text1, [text2], ...)' },
  CONCATENATE: { description: 'Joins text strings', syntax: 'CONCATENATE(text1, [text2], ...)' },
  LEFT: { description: 'Returns leftmost characters', syntax: 'LEFT(text, [num_chars])' },
  RIGHT: { description: 'Returns rightmost characters', syntax: 'RIGHT(text, [num_chars])' },
  MID: { description: 'Returns characters from middle of text', syntax: 'MID(text, start_num, num_chars)' },
  LEN: { description: 'Returns the number of characters', syntax: 'LEN(text)' },
  LOWER: { description: 'Converts text to lowercase', syntax: 'LOWER(text)' },
  UPPER: { description: 'Converts text to uppercase', syntax: 'UPPER(text)' },
  TRIM: { description: 'Removes extra spaces from text', syntax: 'TRIM(text)' },
  TEXT: { description: 'Formats number as text', syntax: 'TEXT(value, format_text)' },
  VALUE: { description: 'Converts text to number', syntax: 'VALUE(text)' },
  SUBSTITUTE: { description: 'Replaces text in a string', syntax: 'SUBSTITUTE(text, old_text, new_text, [instance])' },
  FIND: { description: 'Finds text within another text', syntax: 'FIND(find_text, within_text, [start_num])' },
  DATE: { description: 'Creates a date value', syntax: 'DATE(year, month, day)' },
  YEAR: { description: 'Returns the year of a date', syntax: 'YEAR(serial_number)' },
  MONTH: { description: 'Returns the month of a date', syntax: 'MONTH(serial_number)' },
  DAY: { description: 'Returns the day of a date', syntax: 'DAY(serial_number)' },
  TODAY: { description: 'Returns the current date', syntax: 'TODAY()' },
  MEDIAN: { description: 'Returns the median value', syntax: 'MEDIAN(number1, [number2], ...)' },
  STDEV: { description: 'Estimates standard deviation', syntax: 'STDEV(number1, [number2], ...)' },
  MATCH: { description: 'Searches for a value in range', syntax: 'MATCH(lookup_value, lookup_array, [match_type])' },
  INDEX: { description: 'Returns a value from a range', syntax: 'INDEX(array, row_num, [col_num])' },
  VLOOKUP: { description: 'Looks up value in first column of table', syntax: 'VLOOKUP(lookup_value, table_array, col_index, [range_lookup])' },
  HLOOKUP: { description: 'Looks up value in first row of table', syntax: 'HLOOKUP(lookup_value, table_array, row_index, [range_lookup])' },
  // Medium-priority — Math/Trig
  SIN: { description: 'Returns the sine of an angle', syntax: 'SIN(number)' },
  COS: { description: 'Returns the cosine of an angle', syntax: 'COS(number)' },
  TAN: { description: 'Returns the tangent of an angle', syntax: 'TAN(number)' },
  ASIN: { description: 'Returns the arcsine of a number', syntax: 'ASIN(number)' },
  ACOS: { description: 'Returns the arccosine of a number', syntax: 'ACOS(number)' },
  ATAN: { description: 'Returns the arctangent of a number', syntax: 'ATAN(number)' },
  ATAN2: { description: 'Returns arctangent from x and y coordinates', syntax: 'ATAN2(x_num, y_num)' },
  SINH: { description: 'Returns the hyperbolic sine', syntax: 'SINH(number)' },
  COSH: { description: 'Returns the hyperbolic cosine', syntax: 'COSH(number)' },
  TANH: { description: 'Returns the hyperbolic tangent', syntax: 'TANH(number)' },
  ACOT: { description: 'Returns the arccotangent of a number', syntax: 'ACOT(number)' },
  COT: { description: 'Returns the cotangent of an angle', syntax: 'COT(number)' },
  CSC: { description: 'Returns the cosecant of an angle', syntax: 'CSC(number)' },
  SEC: { description: 'Returns the secant of an angle', syntax: 'SEC(number)' },
  SQRTPI: { description: 'Returns sqrt(number * pi)', syntax: 'SQRTPI(number)' },
  PERMUT: { description: 'Returns number of permutations', syntax: 'PERMUT(number, number_chosen)' },
  COMBINA: { description: 'Returns combinations with repetition', syntax: 'COMBINA(number, number_chosen)' },
  MULTINOMIAL: { description: 'Returns the multinomial of a set of numbers', syntax: 'MULTINOMIAL(number1, [number2], ...)' },
  SERIESSUM: { description: 'Returns the sum of a power series', syntax: 'SERIESSUM(x, n, m, coefficients)' },
  BASE: { description: 'Converts a number to text in given base', syntax: 'BASE(number, radix, [min_length])' },
  DECIMAL: { description: 'Converts a text number in given base to decimal', syntax: 'DECIMAL(text, radix)' },
  ROMAN: { description: 'Converts a number to a Roman numeral', syntax: 'ROMAN(number, [form])' },
  ARABIC: { description: 'Converts a Roman numeral to a number', syntax: 'ARABIC(text)' },
  AGGREGATE: { description: 'Returns an aggregate in a list or database', syntax: 'AGGREGATE(function_num, options, ref1, ...)' },
  'CEILING.PRECISE': { description: 'Rounds number up to nearest multiple', syntax: 'CEILING.PRECISE(number, [significance])' },
  'FLOOR.PRECISE': { description: 'Rounds number down to nearest multiple', syntax: 'FLOOR.PRECISE(number, [significance])' },
  SUMX2MY2: { description: 'Returns sum of difference of squares', syntax: 'SUMX2MY2(array_x, array_y)' },
  SUMX2PY2: { description: 'Returns sum of sum of squares', syntax: 'SUMX2PY2(array_x, array_y)' },
  SUMXMY2: { description: 'Returns sum of squared differences', syntax: 'SUMXMY2(array_x, array_y)' },
  FACTDOUBLE: { description: 'Returns the double factorial', syntax: 'FACTDOUBLE(number)' },
  // Medium-priority — Statistical
  CORREL: { description: 'Returns correlation coefficient', syntax: 'CORREL(array1, array2)' },
  'COVARIANCE.P': { description: 'Returns population covariance', syntax: 'COVARIANCE.P(array1, array2)' },
  'COVARIANCE.S': { description: 'Returns sample covariance', syntax: 'COVARIANCE.S(array1, array2)' },
  SLOPE: { description: 'Returns the slope of a linear regression', syntax: 'SLOPE(known_ys, known_xs)' },
  INTERCEPT: { description: 'Returns the intercept of a linear regression', syntax: 'INTERCEPT(known_ys, known_xs)' },
  RSQ: { description: 'Returns the R-squared value', syntax: 'RSQ(known_ys, known_xs)' },
  FORECAST: { description: 'Calculates a future value using linear regression', syntax: 'FORECAST(x, known_ys, known_xs)' },
  'FORECAST.LINEAR': { description: 'Calculates a future value using linear regression', syntax: 'FORECAST.LINEAR(x, known_ys, known_xs)' },
  FISHER: { description: 'Returns the Fisher transformation', syntax: 'FISHER(x)' },
  FISHERINV: { description: 'Returns the inverse Fisher transformation', syntax: 'FISHERINV(y)' },
  GEOMEAN: { description: 'Returns the geometric mean', syntax: 'GEOMEAN(number1, [number2], ...)' },
  HARMEAN: { description: 'Returns the harmonic mean', syntax: 'HARMEAN(number1, [number2], ...)' },
  TRIMMEAN: { description: 'Returns the trimmed mean', syntax: 'TRIMMEAN(array, percent)' },
  DEVSQ: { description: 'Returns sum of squared deviations', syntax: 'DEVSQ(number1, [number2], ...)' },
  AVEDEV: { description: 'Returns the average absolute deviation', syntax: 'AVEDEV(number1, [number2], ...)' },
  SKEW: { description: 'Returns the skewness of a distribution', syntax: 'SKEW(number1, [number2], ...)' },
  KURT: { description: 'Returns the kurtosis of a data set', syntax: 'KURT(number1, [number2], ...)' },
  STANDARDIZE: { description: 'Returns a normalized value', syntax: 'STANDARDIZE(x, mean, standard_dev)' },
  'CONFIDENCE.NORM': { description: 'Returns a confidence interval for population mean', syntax: 'CONFIDENCE.NORM(alpha, standard_dev, size)' },
  'NORM.S.DIST': { description: 'Returns the standard normal distribution', syntax: 'NORM.S.DIST(z, cumulative)' },
  'NORM.DIST': { description: 'Returns the normal distribution', syntax: 'NORM.DIST(x, mean, standard_dev, cumulative)' },
  'NORM.S.INV': { description: 'Returns the inverse standard normal distribution', syntax: 'NORM.S.INV(probability)' },
  'NORM.INV': { description: 'Returns the inverse normal distribution', syntax: 'NORM.INV(probability, mean, standard_dev)' },
  'POISSON.DIST': { description: 'Returns the Poisson distribution', syntax: 'POISSON.DIST(x, mean, cumulative)' },
  'BINOM.DIST': { description: 'Returns the binomial distribution probability', syntax: 'BINOM.DIST(number_s, trials, probability_s, cumulative)' },
  'WEIBULL.DIST': { description: 'Returns the Weibull distribution', syntax: 'WEIBULL.DIST(x, alpha, beta, cumulative)' },
  'EXPON.DIST': { description: 'Returns the exponential distribution', syntax: 'EXPON.DIST(x, lambda, cumulative)' },
  'LOGNORM.DIST': { description: 'Returns the lognormal distribution', syntax: 'LOGNORM.DIST(x, mean, standard_dev, cumulative)' },
  // Medium-priority — Lookup / Dynamic Arrays
  SORT: { description: 'Sorts the contents of a range or array', syntax: 'SORT(array, [sort_index], [sort_order])' },
  UNIQUE: { description: 'Returns unique values from a range', syntax: 'UNIQUE(array, [by_col], [exactly_once])' },
  FILTER: { description: 'Filters a range based on criteria', syntax: 'FILTER(array, include, [if_empty])' },
  SORTBY: { description: 'Sorts a range by another range', syntax: 'SORTBY(array, by_array1, [sort_order1], ...)' },
  FORMULATEXT: { description: 'Returns a formula as text', syntax: 'FORMULATEXT(reference)' },
  AREAS: { description: 'Returns the number of areas in a reference', syntax: 'AREAS(reference)' },
  // Medium-priority — Text
  T: { description: 'Converts a value to text', syntax: 'T(value)' },
  UNICHAR: { description: 'Returns the Unicode character for a number', syntax: 'UNICHAR(number)' },
  UNICODE: { description: 'Returns the Unicode code point for first character', syntax: 'UNICODE(text)' },
  TEXTSPLIT: { description: 'Splits text by delimiter', syntax: 'TEXTSPLIT(text, col_delimiter, [row_delimiter])' },
  VALUETOTEXT: { description: 'Converts any value to text', syntax: 'VALUETOTEXT(value, [format])' },
  ARRAYTOTEXT: { description: 'Converts an array to text', syntax: 'ARRAYTOTEXT(array, [format])' },
  ENCODEURL: { description: 'URL-encodes a text string', syntax: 'ENCODEURL(text)' },
  // Medium-priority — Date/Time
  'WORKDAY.INTL': { description: 'Returns a date offset by work days', syntax: 'WORKDAY.INTL(start_date, days, [weekend], [holidays])' },
  'NETWORKDAYS.INTL': { description: 'Returns number of working days between dates', syntax: 'NETWORKDAYS.INTL(start_date, end_date, [weekend], [holidays])' },
  // Medium-priority — Financial
  FV: { description: 'Returns the future value of an investment', syntax: 'FV(rate, nper, pmt, [pv], [type])' },
  PV: { description: 'Returns the present value of an investment', syntax: 'PV(rate, nper, pmt, [fv], [type])' },
  PMT: { description: 'Returns the periodic payment for a loan', syntax: 'PMT(rate, nper, pv, [fv], [type])' },
  NPER: { description: 'Returns the number of periods for an investment', syntax: 'NPER(rate, pmt, pv, [fv], [type])' },
  RATE: { description: 'Returns the interest rate per period', syntax: 'RATE(nper, pmt, pv, [fv], [type], [guess])' },
  NPV: { description: 'Returns the net present value', syntax: 'NPV(rate, value1, [value2], ...)' },
  IRR: { description: 'Returns the internal rate of return', syntax: 'IRR(values, [guess])' },
  MIRR: { description: 'Returns the modified internal rate of return', syntax: 'MIRR(values, finance_rate, reinvest_rate)' },
  SLN: { description: 'Returns straight-line depreciation', syntax: 'SLN(cost, salvage, life)' },
  SYD: { description: 'Returns sum-of-years-digits depreciation', syntax: 'SYD(cost, salvage, life, per)' },
  DB: { description: 'Returns declining balance depreciation', syntax: 'DB(cost, salvage, life, period, [month])' },
  DDB: { description: 'Returns double-declining balance depreciation', syntax: 'DDB(cost, salvage, life, period, [factor])' },
  PPMT: { description: 'Returns principal portion of a payment', syntax: 'PPMT(rate, per, nper, pv, [fv], [type])' },
  IPMT: { description: 'Returns interest portion of a payment', syntax: 'IPMT(rate, per, nper, pv, [fv], [type])' },
  CUMIPMT: { description: 'Returns cumulative interest paid', syntax: 'CUMIPMT(rate, nper, pv, start_period, end_period, type)' },
  CUMPRINC: { description: 'Returns cumulative principal paid', syntax: 'CUMPRINC(rate, nper, pv, start_period, end_period, type)' },
  EFFECT: { description: 'Returns the effective annual interest rate', syntax: 'EFFECT(nominal_rate, npery)' },
  NOMINAL: { description: 'Returns the nominal annual interest rate', syntax: 'NOMINAL(effect_rate, npery)' },
  DOLLARDE: { description: 'Converts fractional dollar to decimal', syntax: 'DOLLARDE(fractional_dollar, fraction)' },
  DOLLARFR: { description: 'Converts decimal dollar to fractional', syntax: 'DOLLARFR(decimal_dollar, fraction)' },
  // Medium-priority — Information
  CELL: { description: 'Returns information about a cell', syntax: 'CELL(info_type, [reference])' },
  INFO: { description: 'Returns information about the environment', syntax: 'INFO(type_text)' },
  // --- Additional implemented functions ---
  // Math/Trig
  CEILING: { description: 'Rounds a number up to the nearest multiple of significance', syntax: 'CEILING(number, [significance])' },
  'CEILING.MATH': { description: 'Rounds a number up to the nearest integer or multiple of significance', syntax: 'CEILING.MATH(number, [significance], [mode])' },
  COMBIN: { description: 'Returns the number of combinations for a given number of items', syntax: 'COMBIN(number, number_chosen)' },
  DEGREES: { description: 'Converts radians to degrees', syntax: 'DEGREES(angle)' },
  EVEN: { description: 'Rounds a number up to the nearest even integer', syntax: 'EVEN(number)' },
  EXP: { description: 'Returns e raised to the power of a given number', syntax: 'EXP(number)' },
  FACT: { description: 'Returns the factorial of a number', syntax: 'FACT(number)' },
  FLOOR: { description: 'Rounds a number down to the nearest multiple of significance', syntax: 'FLOOR(number, [significance])' },
  'FLOOR.MATH': { description: 'Rounds a number down to the nearest integer or multiple of significance', syntax: 'FLOOR.MATH(number, [significance], [mode])' },
  FREQUENCY: { description: 'Returns a frequency distribution as a vertical array', syntax: 'FREQUENCY(data_array, bins_array)' },
  GCD: { description: 'Returns the greatest common divisor', syntax: 'GCD(number1, [number2], ...)' },
  INT: { description: 'Rounds a number down to the nearest integer', syntax: 'INT(number)' },
  LCM: { description: 'Returns the least common multiple', syntax: 'LCM(number1, [number2], ...)' },
  LOG: { description: 'Returns the logarithm of a number to a specified base', syntax: 'LOG(number, [base])' },
  LOG10: { description: 'Returns the base-10 logarithm of a number', syntax: 'LOG10(number)' },
  MROUND: { description: 'Returns a number rounded to the desired multiple', syntax: 'MROUND(number, multiple)' },
  ODD: { description: 'Rounds a number up to the nearest odd integer', syntax: 'ODD(number)' },
  PI: { description: 'Returns the value of pi', syntax: 'PI()' },
  PRODUCT: { description: 'Multiplies all the numbers given as arguments', syntax: 'PRODUCT(number1, [number2], ...)' },
  QUOTIENT: { description: 'Returns the integer portion of a division', syntax: 'QUOTIENT(numerator, denominator)' },
  RADIANS: { description: 'Converts degrees to radians', syntax: 'RADIANS(angle)' },
  RAND: { description: 'Returns a random number between 0 and 1', syntax: 'RAND()' },
  RANDBETWEEN: { description: 'Returns a random integer between two specified numbers', syntax: 'RANDBETWEEN(bottom, top)' },
  ROUNDDOWN: { description: 'Rounds a number down, toward zero', syntax: 'ROUNDDOWN(number, num_digits)' },
  ROUNDUP: { description: 'Rounds a number up, away from zero', syntax: 'ROUNDUP(number, num_digits)' },
  SIGN: { description: 'Returns the sign of a number (1, 0, or -1)', syntax: 'SIGN(number)' },
  SUBTOTAL: { description: 'Returns a subtotal in a list or database', syntax: 'SUBTOTAL(function_num, ref1, [ref2], ...)' },
  SUMPRODUCT: { description: 'Returns the sum of the products of corresponding array components', syntax: 'SUMPRODUCT(array1, [array2], ...)' },
  SUMSQ: { description: 'Returns the sum of the squares of the arguments', syntax: 'SUMSQ(number1, [number2], ...)' },
  TRUNC: { description: 'Truncates a number to an integer or specified number of digits', syntax: 'TRUNC(number, [num_digits])' },
  // Statistical
  AVERAGEA: { description: 'Returns the average of values, including text and logicals', syntax: 'AVERAGEA(value1, [value2], ...)' },
  COUNTBLANK: { description: 'Counts the number of empty cells in a range', syntax: 'COUNTBLANK(range)' },
  COUNTIF: { description: 'Counts the number of cells that meet a criteria', syntax: 'COUNTIF(range, criteria)' },
  LARGE: { description: 'Returns the k-th largest value in a data set', syntax: 'LARGE(array, k)' },
  MAXA: { description: 'Returns the largest value, including text and logicals', syntax: 'MAXA(value1, [value2], ...)' },
  MAXIFS: { description: 'Returns the maximum value among cells specified by criteria', syntax: 'MAXIFS(max_range, criteria_range1, criteria1, ...)' },
  MINA: { description: 'Returns the smallest value, including text and logicals', syntax: 'MINA(value1, [value2], ...)' },
  MINIFS: { description: 'Returns the minimum value among cells specified by criteria', syntax: 'MINIFS(min_range, criteria_range1, criteria1, ...)' },
  MODE: { description: 'Returns the most frequently occurring value in a data set', syntax: 'MODE(number1, [number2], ...)' },
  'MODE.SNGL': { description: 'Returns the most frequently occurring value in a data set', syntax: 'MODE.SNGL(number1, [number2], ...)' },
  PERCENTILE: { description: 'Returns the k-th percentile of values in a range', syntax: 'PERCENTILE(array, k)' },
  'PERCENTILE.EXC': { description: 'Returns the k-th percentile of values (exclusive)', syntax: 'PERCENTILE.EXC(array, k)' },
  'PERCENTILE.INC': { description: 'Returns the k-th percentile of values (inclusive)', syntax: 'PERCENTILE.INC(array, k)' },
  QUARTILE: { description: 'Returns the quartile of a data set', syntax: 'QUARTILE(array, quart)' },
  'QUARTILE.INC': { description: 'Returns the quartile of a data set (inclusive)', syntax: 'QUARTILE.INC(array, quart)' },
  RANK: { description: 'Returns the rank of a number in a list of numbers', syntax: 'RANK(number, ref, [order])' },
  'RANK.AVG': { description: 'Returns the rank of a number in a list, averaging ties', syntax: 'RANK.AVG(number, ref, [order])' },
  'RANK.EQ': { description: 'Returns the rank of a number in a list of numbers', syntax: 'RANK.EQ(number, ref, [order])' },
  SMALL: { description: 'Returns the k-th smallest value in a data set', syntax: 'SMALL(array, k)' },
  'STDEV.P': { description: 'Calculates standard deviation based on the entire population', syntax: 'STDEV.P(number1, [number2], ...)' },
  'STDEV.S': { description: 'Estimates standard deviation based on a sample', syntax: 'STDEV.S(number1, [number2], ...)' },
  STDEVP: { description: 'Calculates standard deviation based on the entire population', syntax: 'STDEVP(number1, [number2], ...)' },
  VAR: { description: 'Estimates variance based on a sample', syntax: 'VAR(number1, [number2], ...)' },
  'VAR.P': { description: 'Calculates variance based on the entire population', syntax: 'VAR.P(number1, [number2], ...)' },
  'VAR.S': { description: 'Estimates variance based on a sample', syntax: 'VAR.S(number1, [number2], ...)' },
  VARP: { description: 'Calculates variance based on the entire population', syntax: 'VARP(number1, [number2], ...)' },
  // Lookup / Reference
  ADDRESS: { description: 'Returns a cell reference as text', syntax: 'ADDRESS(row_num, column_num, [abs_num], [a1], [sheet_text])' },
  CHOOSE: { description: 'Returns a value from a list based on an index number', syntax: 'CHOOSE(index_num, value1, [value2], ...)' },
  COLUMN: { description: 'Returns the column number of a reference', syntax: 'COLUMN([reference])' },
  COLUMNS: { description: 'Returns the number of columns in a reference or array', syntax: 'COLUMNS(array)' },
  LOOKUP: { description: 'Looks up a value in a vector and returns from the same position in another', syntax: 'LOOKUP(lookup_value, lookup_vector, [result_vector])' },
  ROW: { description: 'Returns the row number of a reference', syntax: 'ROW([reference])' },
  ROWS: { description: 'Returns the number of rows in a reference or array', syntax: 'ROWS(array)' },
  TRANSPOSE: { description: 'Transposes the rows and columns of an array', syntax: 'TRANSPOSE(array)' },
  XLOOKUP: { description: 'Searches a range or array and returns a matching item', syntax: 'XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])' },
  XMATCH: { description: 'Returns the relative position of a value in an array', syntax: 'XMATCH(lookup_value, lookup_array, [match_mode], [search_mode])' },
  // Text
  CHAR: { description: 'Returns the character specified by a code number', syntax: 'CHAR(number)' },
  CLEAN: { description: 'Removes all non-printable characters from text', syntax: 'CLEAN(text)' },
  CODE: { description: 'Returns the numeric code for the first character in a text string', syntax: 'CODE(text)' },
  DOLLAR: { description: 'Converts a number to text in currency format', syntax: 'DOLLAR(number, [decimals])' },
  EXACT: { description: 'Checks whether two text strings are exactly the same', syntax: 'EXACT(text1, text2)' },
  FIXED: { description: 'Formats a number as text with a fixed number of decimals', syntax: 'FIXED(number, [decimals], [no_commas])' },
  NUMBERVALUE: { description: 'Converts text to a number in a locale-independent way', syntax: 'NUMBERVALUE(text, [decimal_separator], [group_separator])' },
  PROPER: { description: 'Capitalizes the first letter of each word in a text string', syntax: 'PROPER(text)' },
  REPLACE: { description: 'Replaces part of a text string with a different text string', syntax: 'REPLACE(old_text, start_num, num_chars, new_text)' },
  REPT: { description: 'Repeats text a given number of times', syntax: 'REPT(text, number_times)' },
  SEARCH: { description: 'Finds text within another text (case-insensitive)', syntax: 'SEARCH(find_text, within_text, [start_num])' },
  TEXTAFTER: { description: 'Returns text that occurs after a given delimiter', syntax: 'TEXTAFTER(text, delimiter, [instance_num])' },
  TEXTBEFORE: { description: 'Returns text that occurs before a given delimiter', syntax: 'TEXTBEFORE(text, delimiter, [instance_num])' },
  TEXTJOIN: { description: 'Combines text from multiple ranges with a delimiter', syntax: 'TEXTJOIN(delimiter, ignore_empty, text1, [text2], ...)' },
  // Date/Time
  DATEDIF: { description: 'Returns the difference between two dates', syntax: 'DATEDIF(start_date, end_date, unit)' },
  DATEVALUE: { description: 'Converts a date in text format to a serial number', syntax: 'DATEVALUE(date_text)' },
  DAYS: { description: 'Returns the number of days between two dates', syntax: 'DAYS(end_date, start_date)' },
  DAYS360: { description: 'Returns days between dates based on a 360-day year', syntax: 'DAYS360(start_date, end_date, [method])' },
  EDATE: { description: 'Returns a date a given number of months before or after a start date', syntax: 'EDATE(start_date, months)' },
  EOMONTH: { description: 'Returns the last day of the month a number of months before or after', syntax: 'EOMONTH(start_date, months)' },
  HOUR: { description: 'Returns the hour component of a serial number', syntax: 'HOUR(serial_number)' },
  ISOWEEKNUM: { description: 'Returns the ISO week number of the year for a given date', syntax: 'ISOWEEKNUM(date)' },
  MINUTE: { description: 'Returns the minute component of a serial number', syntax: 'MINUTE(serial_number)' },
  NETWORKDAYS: { description: 'Returns the number of whole working days between two dates', syntax: 'NETWORKDAYS(start_date, end_date, [holidays])' },
  NOW: { description: 'Returns the serial number of the current date and time', syntax: 'NOW()' },
  SECOND: { description: 'Returns the second component of a serial number', syntax: 'SECOND(serial_number)' },
  TIME: { description: 'Returns the serial number of a particular time', syntax: 'TIME(hour, minute, second)' },
  TIMEVALUE: { description: 'Converts a time in text format to a serial number', syntax: 'TIMEVALUE(time_text)' },
  WEEKDAY: { description: 'Returns the day of the week for a date', syntax: 'WEEKDAY(serial_number, [return_type])' },
  WEEKNUM: { description: 'Returns the week number in the year', syntax: 'WEEKNUM(serial_number, [return_type])' },
  WORKDAY: { description: 'Returns a date a specified number of working days ahead', syntax: 'WORKDAY(start_date, days, [holidays])' },
  YEARFRAC: { description: 'Returns the fraction of the year between two dates', syntax: 'YEARFRAC(start_date, end_date, [basis])' },
  // Logical
  IFNA: { description: 'Returns a value if formula evaluates to #N/A, otherwise the result', syntax: 'IFNA(value, value_if_na)' },
  IFS: { description: 'Checks multiple conditions and returns the first TRUE result', syntax: 'IFS(logical_test1, value_if_true1, ...)' },
  SWITCH: { description: 'Evaluates an expression against a list of values', syntax: 'SWITCH(expression, value1, result1, ..., [default])' },
  XOR: { description: 'Returns TRUE if an odd number of arguments are TRUE', syntax: 'XOR(logical1, [logical2], ...)' },
  // Information
  'ERROR.TYPE': { description: 'Returns a number corresponding to an error type', syntax: 'ERROR.TYPE(error_val)' },
  ISERR: { description: 'Returns TRUE if the value is any error except #N/A', syntax: 'ISERR(value)' },
  ISEVEN: { description: 'Returns TRUE if the number is even', syntax: 'ISEVEN(number)' },
  ISFORMULA: { description: 'Returns TRUE if the referenced cell contains a formula', syntax: 'ISFORMULA(reference)' },
  ISLOGICAL: { description: 'Returns TRUE if the value is a logical value', syntax: 'ISLOGICAL(value)' },
  ISNA: { description: 'Returns TRUE if the value is the #N/A error', syntax: 'ISNA(value)' },
  ISNONTEXT: { description: 'Returns TRUE if the value is not text', syntax: 'ISNONTEXT(value)' },
  ISODD: { description: 'Returns TRUE if the number is odd', syntax: 'ISODD(number)' },
  ISREF: { description: 'Returns TRUE if the value is a reference', syntax: 'ISREF(value)' },
  NA: { description: 'Returns the error value #N/A', syntax: 'NA()' },
  SHEET: { description: 'Returns the sheet number of the current sheet', syntax: 'SHEET([value])' },
  SHEETS: { description: 'Returns the number of sheets in a reference', syntax: 'SHEETS([reference])' },
  TYPE: { description: 'Returns a number indicating the data type of a value', syntax: 'TYPE(value)' },
};

/**
 * FormulaAutocomplete manages the autocomplete dropdown for the formula bar.
 */
export class FormulaAutocomplete {
  private container: HTMLElement;
  private dropdown: HTMLElement;
  private suggestions: AutocompleteSuggestion[] = [];
  private displayedSuggestions: AutocompleteSuggestion[] = [];
  private selectedIndex = -1;
  private visible = false;
  private allFunctions: AutocompleteSuggestion[];
  private activeFunctionContext: ActiveFunctionContext | null = null;

  /** Callback when a function is selected */
  onSelect?: (name: string) => void;

  constructor(container: HTMLElement) {
    this.container = container;

    // Build full function list
    const builtinNames = Object.keys(BuiltinFunctions.create());
    this.allFunctions = builtinNames.map((name) => {
      const meta = FUNCTION_METADATA[name];
      return {
        name,
        description: meta?.description || `Excel function ${name}`,
        syntax: meta?.syntax || `${name}(...)`,
      };
    }).sort((a, b) => a.name.localeCompare(b.name));

    // Create dropdown element
    this.dropdown = document.createElement('div');
    this.dropdown.className = 'formula-autocomplete';
    this.dropdown.setAttribute('role', 'listbox');
    this.dropdown.style.display = 'none';
    this.container.appendChild(this.dropdown);
  }

  /**
   * Update suggestions based on the current input text and cursor position.
   */
  update(formulaText: string, cursorPos: number): void {
    // Extract the token being typed at cursor position
    const token = this.extractCurrentToken(formulaText, cursorPos);
    const upperToken = token ? token.toUpperCase() : null;
    this.activeFunctionContext = this.getActiveFunctionContext(formulaText, cursorPos);

    if (!token || token.length === 0) {
      if (this.activeFunctionContext) {
        const activeFn = this.findFunctionByName(this.activeFunctionContext.name);
        this.suggestions = activeFn ? [activeFn] : [];
      } else {
        this.suggestions = [];
      }
    } else {
      this.suggestions = this.allFunctions.filter((f) => f.name.startsWith(upperToken!));

      // Keep function hint visible while editing arguments/cell refs
      if (this.suggestions.length === 0 && this.activeFunctionContext) {
        const activeFn = this.findFunctionByName(this.activeFunctionContext.name);
        this.suggestions = activeFn ? [activeFn] : [];
      }

      // Place exact match first when full function name is typed.
      this.suggestions.sort((a, b) => {
        const aExact = a.name === upperToken ? 1 : 0;
        const bExact = b.name === upperToken ? 1 : 0;
        if (aExact !== bExact) return bExact - aExact;
        return a.name.localeCompare(b.name);
      });
    }

    if (this.suggestions.length === 0) {
      this.hide();
      return;
    }

    // Limit displayed results
    const maxVisible = 8;
    this.displayedSuggestions = this.suggestions.slice(0, maxVisible);

    this.selectedIndex = 0;
    this.renderDropdown(this.displayedSuggestions, upperToken);
    this.show();
  }

  /**
   * Handle keyboard events for the autocomplete.
   * Returns true if the event was consumed.
   */
  handleKey(e: KeyboardEvent): boolean {
    if (!this.visible) return false;

    switch (e.key) {
      case 'ArrowDown':
        e.preventDefault();
        this.selectedIndex = Math.min(this.selectedIndex + 1, this.displayedSuggestions.length - 1);
        this.highlightSelected();
        return true;

      case 'ArrowUp':
        e.preventDefault();
        this.selectedIndex = Math.max(this.selectedIndex - 1, 0);
        this.highlightSelected();
        return true;

      case 'Tab':
      case 'Enter':
        if (this.selectedIndex >= 0 && this.selectedIndex < this.displayedSuggestions.length) {
          e.preventDefault();
          this.selectCurrent();
          return true;
        }
        return false;

      case 'Escape':
        e.preventDefault();
        this.hide();
        return true;

      default:
        return false;
    }
  }

  /**
   * Hide the dropdown.
   */
  hide(): void {
    this.visible = false;
    this.dropdown.style.display = 'none';
    this.selectedIndex = -1;
    this.displayedSuggestions = [];
    this.activeFunctionContext = null;
  }

  /**
   * Show the dropdown.
   */
  private show(): void {
    this.visible = true;
    this.dropdown.style.display = 'block';
  }

  /**
   * Check if dropdown is visible.
   */
  isVisible(): boolean {
    return this.visible;
  }

  /**
   * Destroy the autocomplete (cleanup DOM).
   */
  dispose(): void {
    this.dropdown.remove();
  }

  // ---- Private helpers ----

  /**
   * Extract the function name token at the current cursor position.
   * Returns the partial function name being typed, or null.
   */
  private extractCurrentToken(text: string, cursorPos: number): string | null {
    // Work backwards from cursor to find the start of the current token
    const beforeCursor = text.slice(0, cursorPos);

    // Find the last non-identifier character
    const match = /([A-Za-z_][A-Za-z0-9_]*)$/.exec(beforeCursor);
    if (!match) return null;

    const token = match[1];

    // Only suggest if we're after '=', '(', ',', '+', '-', '*', '/', or at start
    const charBefore = beforeCursor.slice(0, beforeCursor.length - token.length).slice(-1);
    if (charBefore && !['=', '(', ',', '+', '-', '*', '/', '^', '<', '>', ' ', '&'].includes(charBefore)) {
      return null;
    }

    return token.length >= 1 ? token : null;
  }

  /**
   * Render the dropdown items.
   */
  private renderDropdown(items: AutocompleteSuggestion[], upperToken: string | null): void {
    this.dropdown.innerHTML = '';

    items.forEach((item, index) => {
      const div = document.createElement('div');
      div.className = 'formula-autocomplete-item' + (index === this.selectedIndex ? ' selected' : '');
      div.setAttribute('role', 'option');
      div.dataset.index = String(index);

      const nameSpan = document.createElement('span');
      nameSpan.className = 'autocomplete-name';
      nameSpan.textContent = item.name;

      const descSpan = document.createElement('span');
      descSpan.className = 'autocomplete-desc';
      descSpan.textContent = item.description;

      const syntaxSpan = document.createElement('div');
      syntaxSpan.className = 'autocomplete-syntax';
      const highlightedIndex = this.getHighlightedParameterIndex(item.name, upperToken);
      if (highlightedIndex === null) {
        syntaxSpan.textContent = item.syntax;
      } else {
        syntaxSpan.innerHTML = this.formatSyntaxWithHighlightedParameter(item.syntax, highlightedIndex);
      }

      div.appendChild(nameSpan);
      div.appendChild(descSpan);
      div.appendChild(syntaxSpan);

      div.addEventListener('mousedown', (e) => {
        e.preventDefault(); // Don't blur the input
        this.selectedIndex = index;
        this.selectCurrent();
      });

      div.addEventListener('mouseenter', () => {
        this.selectedIndex = index;
        this.highlightSelected();
      });

      this.dropdown.appendChild(div);
    });
  }

  /**
   * Highlight the currently selected item.
   */
  private highlightSelected(): void {
    const items = this.dropdown.querySelectorAll('.formula-autocomplete-item');
    items.forEach((item, i) => {
      item.classList.toggle('selected', i === this.selectedIndex);
    });
  }

  /**
   * Select the currently highlighted suggestion.
   */
  private selectCurrent(): void {
    if (this.selectedIndex >= 0 && this.selectedIndex < this.displayedSuggestions.length) {
      const selected = this.displayedSuggestions[this.selectedIndex];
      this.onSelect?.(selected.name);
      this.hide();
    }
  }

  private findFunctionByName(name: string): AutocompleteSuggestion | null {
    return this.allFunctions.find((f) => f.name === name) || null;
  }

  private getHighlightedParameterIndex(functionName: string, upperToken: string | null): number | null {
    if (this.activeFunctionContext?.name === functionName) {
      return this.activeFunctionContext.parameterIndex;
    }

    // When a full function name is typed (without opening parenthesis),
    // keep exact function visible with a first-parameter hint.
    if (upperToken && functionName === upperToken) {
      return 0;
    }

    return null;
  }

  private formatSyntaxWithHighlightedParameter(syntax: string, parameterIndex: number): string {
    const openParen = syntax.indexOf('(');
    const closeParen = syntax.lastIndexOf(')');
    if (openParen < 0 || closeParen <= openParen) {
      return this.escapeHtml(syntax);
    }

    const fnName = syntax.slice(0, openParen);
    const rawParams = syntax.slice(openParen + 1, closeParen).trim();
    if (!rawParams) {
      return `${this.escapeHtml(fnName)}()`;
    }

    const params = rawParams.split(/\s*,\s*/);
    const safeIndex = Math.max(0, Math.min(parameterIndex, params.length - 1));
    const highlighted = params.map((param, index) => {
      const escaped = this.escapeHtml(param);
      return index === safeIndex ? `<strong>${escaped}</strong>` : escaped;
    }).join(', ');

    return `${this.escapeHtml(fnName)}(${highlighted})`;
  }

  private escapeHtml(text: string): string {
    return text
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  private getActiveFunctionContext(text: string, cursorPos: number): ActiveFunctionContext | null {
    const beforeCursor = text.slice(0, cursorPos);
    const stack: ActiveFunctionContext[] = [];
    let inString = false;

    for (let i = 0; i < beforeCursor.length; i++) {
      const ch = beforeCursor[i];

      if (ch === '"') {
        // Excel escapes quote inside strings as "".
        if (inString && beforeCursor[i + 1] === '"') {
          i++;
          continue;
        }
        inString = !inString;
        continue;
      }

      if (inString) {
        continue;
      }

      if (/[A-Za-z_]/.test(ch)) {
        let end = i + 1;
        while (end < beforeCursor.length && /[A-Za-z0-9_]/.test(beforeCursor[end])) {
          end++;
        }

        const identifier = beforeCursor.slice(i, end);
        let lookAhead = end;
        while (lookAhead < beforeCursor.length && /\s/.test(beforeCursor[lookAhead])) {
          lookAhead++;
        }

        if (beforeCursor[lookAhead] === '(') {
          stack.push({ name: identifier.toUpperCase(), parameterIndex: 0 });
          i = lookAhead;
          continue;
        }

        i = end - 1;
        continue;
      }

      if (ch === ',' && stack.length > 0) {
        stack[stack.length - 1].parameterIndex += 1;
        continue;
      }

      if (ch === ')' && stack.length > 0) {
        stack.pop();
      }
    }

    return stack.length > 0 ? stack[stack.length - 1] : null;
  }
}
