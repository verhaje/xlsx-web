// helpers.ts - Shared utilities used by multiple function categories

/**
 * Criteria matcher used by COUNTIFS, SUMIF, SUMIFS, AVERAGEIF, AVERAGEIFS, etc.
 */
export function matchCriteria(val: any, crit: any): boolean {
  if (crit === null || crit === undefined || crit === '') {
    return val === '' || val === null || val === undefined;
  }
  if (typeof crit === 'string') {
    // Operator prefix
    const m = /^(<>|[<>]=?|=)\s*(.*)$/s.exec(crit);
    if (m) {
      const op = m[1];
      const rhs = m[2];
      const rn = Number(rhs);
      const isRhsNum = !Number.isNaN(rn) && rhs !== '';
      const vn = Number(val);
      const isValNum = !Number.isNaN(vn) && val !== '' && val !== null;
      if (isRhsNum) {
        if (!isValNum) return false;
        switch (op) {
          case '>': return vn > rn;
          case '<': return vn < rn;
          case '>=': return vn >= rn;
          case '<=': return vn <= rn;
          case '=': return vn === rn;
          case '<>': return vn !== rn;
        }
      }
      const vs = String(val ?? '');
      switch (op) {
        case '>': return vs > rhs;
        case '<': return vs < rhs;
        case '>=': return vs >= rhs;
        case '<=': return vs <= rhs;
        case '=': return vs.toLowerCase() === rhs.toLowerCase();
        case '<>': return vs.toLowerCase() !== rhs.toLowerCase();
      }
    }
    // Wildcard support (* and ?)
    if (/[*?]/.test(crit)) {
      const esc = crit.replace(/[-\\^$+.()|[\]{}]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.');
      try {
        const re = new RegExp(`^${esc}$`, 'i');
        return re.test(String(val ?? ''));
      } catch {
        return false;
      }
    }
    // Equality
    return String(val ?? '').toLowerCase() === crit.toLowerCase();
  }
  return val == crit;
}

/**
 * VLOOKUP / HLOOKUP / MATCH comparison helper.
 */
export function compareValues(a: any, b: any): number {
  const aNum = Number(a);
  const bNum = Number(b);
  const aIsNum = !Number.isNaN(aNum) && a !== '' && a !== null;
  const bIsNum = !Number.isNaN(bNum) && b !== '' && b !== null;
  if (aIsNum && bIsNum) return aNum - bNum;
  const aStr = String(a ?? '').toLowerCase();
  const bStr = String(b ?? '').toLowerCase();
  return aStr.localeCompare(bStr);
}
