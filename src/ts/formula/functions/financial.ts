// financial.ts - Financial functions (TVM, depreciation, etc.)

import { ERRORS } from '../Errors';
import type { BuiltinFunction } from '../../types';

export function getFinancialFunctions(): Record<string, BuiltinFunction> {
  return {
    FV: async (args) => {
      const rate = Number(args[0]);
      const nper = Number(args[1]);
      const pmt = Number(args[2]) || 0;
      const pv = Number(args[3]) || 0;
      const type = Number(args[4]) || 0;
      if (!Number.isFinite(rate) || !Number.isFinite(nper)) return ERRORS.VALUE;
      if (rate === 0) return -(pv + pmt * nper);
      const factor = Math.pow(1 + rate, nper);
      return -(pv * factor + pmt * (1 + rate * type) * (factor - 1) / rate);
    },
    PV: async (args) => {
      const rate = Number(args[0]);
      const nper = Number(args[1]);
      const pmt = Number(args[2]) || 0;
      const fv = Number(args[3]) || 0;
      const type = Number(args[4]) || 0;
      if (!Number.isFinite(rate) || !Number.isFinite(nper)) return ERRORS.VALUE;
      if (rate === 0) return -(fv + pmt * nper);
      const factor = Math.pow(1 + rate, nper);
      return -(fv + pmt * (1 + rate * type) * (factor - 1) / rate) / factor;
    },
    PMT: async (args) => {
      const rate = Number(args[0]);
      const nper = Number(args[1]);
      const pv = Number(args[2]);
      const fv = Number(args[3]) || 0;
      const type = Number(args[4]) || 0;
      if (!Number.isFinite(rate) || !Number.isFinite(nper) || !Number.isFinite(pv)) return ERRORS.VALUE;
      if (nper === 0) return ERRORS.NUM;
      if (rate === 0) return -(pv + fv) / nper;
      const factor = Math.pow(1 + rate, nper);
      return -(rate * (pv * factor + fv)) / ((1 + rate * type) * (factor - 1));
    },
    NPER: async (args) => {
      const rate = Number(args[0]);
      const pmt = Number(args[1]);
      const pv = Number(args[2]);
      const fv = Number(args[3]) || 0;
      const type = Number(args[4]) || 0;
      if (!Number.isFinite(rate) || !Number.isFinite(pmt) || !Number.isFinite(pv)) return ERRORS.VALUE;
      if (rate === 0) {
        if (pmt === 0) return ERRORS.NUM;
        return -(pv + fv) / pmt;
      }
      const pmtAdj = pmt * (1 + rate * type);
      const num = pmtAdj - fv * rate;
      const den = pv * rate + pmtAdj;
      if (num / den <= 0) return ERRORS.NUM;
      return Math.log(num / den) / Math.log(1 + rate);
    },
    RATE: async (args) => {
      const nper = Number(args[0]);
      const pmt = Number(args[1]);
      const pv = Number(args[2]);
      const fv = Number(args[3]) || 0;
      const type = Number(args[4]) || 0;
      const guessArg = args[5] !== undefined ? Number(args[5]) : 0.1;
      if (!Number.isFinite(nper) || !Number.isFinite(pmt) || !Number.isFinite(pv)) return ERRORS.VALUE;
      let rate = guessArg;
      for (let iter = 0; iter < 100; iter++) {
        if (rate <= -1) return ERRORS.NUM;
        const factor = Math.pow(1 + rate, nper);
        const fVal = pv * factor + pmt * (1 + rate * type) * (factor - 1) / rate + fv;
        const dfVal = pv * nper * Math.pow(1 + rate, nper - 1) +
          pmt * (1 + rate * type) * ((nper * Math.pow(1 + rate, nper - 1) * rate - (factor - 1)) / (rate * rate)) +
          pmt * type * (factor - 1) / rate;
        if (Math.abs(dfVal) < 1e-15) return ERRORS.NUM;
        const newRate = rate - fVal / dfVal;
        if (Math.abs(newRate - rate) < 1e-10) return newRate;
        rate = newRate;
      }
      return ERRORS.NUM;
    },
    NPV: async (args) => {
      const rate = Number(args[0]);
      if (!Number.isFinite(rate)) return ERRORS.VALUE;
      let npv = 0;
      for (let i = 1; i < args.length; i++) {
        const cf = Number(args[i]);
        if (Number.isNaN(cf)) continue;
        npv += cf / Math.pow(1 + rate, i);
      }
      return npv;
    },
    IRR: async (args) => {
      const values = args[0];
      const guess = args[1] !== undefined ? Number(args[1]) : 0.1;
      if (!Array.isArray(values) || values.length < 2) return ERRORS.NUM;
      const cfs = values.map((v: any) => Number(v));
      if (cfs.some((nn3: number) => Number.isNaN(nn3))) return ERRORS.VALUE;
      let rate = guess;
      for (let iter = 0; iter < 200; iter++) {
        let npv = 0, dnpv = 0;
        for (let i = 0; i < cfs.length; i++) {
          const f = Math.pow(1 + rate, i);
          npv += cfs[i] / f;
          dnpv -= i * cfs[i] / (f * (1 + rate));
        }
        if (Math.abs(dnpv) < 1e-15) return ERRORS.NUM;
        const newRate = rate - npv / dnpv;
        if (Math.abs(newRate - rate) < 1e-10) return newRate;
        rate = newRate;
      }
      return ERRORS.NUM;
    },
    MIRR: async (args) => {
      const values = args[0];
      const financeRate = Number(args[1]);
      const reinvestRate = Number(args[2]);
      if (!Array.isArray(values) || values.length < 2) return ERRORS.VALUE;
      if (!Number.isFinite(financeRate) || !Number.isFinite(reinvestRate)) return ERRORS.VALUE;
      const cfs = values.map((v: any) => Number(v));
      const n = cfs.length;
      let pvNeg = 0, fvPos = 0;
      for (let i = 0; i < n; i++) {
        if (cfs[i] < 0) pvNeg += cfs[i] / Math.pow(1 + financeRate, i);
        else fvPos += cfs[i] * Math.pow(1 + reinvestRate, n - 1 - i);
      }
      if (pvNeg === 0) return ERRORS.DIV0;
      return Math.pow(-fvPos / pvNeg, 1 / (n - 1)) - 1;
    },
    SLN: async (args) => {
      const cost = Number(args[0]);
      const salvage = Number(args[1]);
      const life = Number(args[2]);
      if (!Number.isFinite(cost) || !Number.isFinite(salvage) || !Number.isFinite(life)) return ERRORS.VALUE;
      if (life === 0) return ERRORS.DIV0;
      return (cost - salvage) / life;
    },
    SYD: async (args) => {
      const cost = Number(args[0]);
      const salvage = Number(args[1]);
      const life = Number(args[2]);
      const per = Number(args[3]);
      if (!Number.isFinite(cost) || !Number.isFinite(salvage) || !Number.isFinite(life) || !Number.isFinite(per)) return ERRORS.VALUE;
      if (life <= 0 || per <= 0 || per > life) return ERRORS.NUM;
      return (cost - salvage) * (life - per + 1) * 2 / (life * (life + 1));
    },
    DB: async (args) => {
      const cost = Number(args[0]);
      const salvage = Number(args[1]);
      const life = Number(args[2]);
      const period = Number(args[3]);
      const month = args[4] !== undefined ? Number(args[4]) : 12;
      if (!Number.isFinite(cost) || !Number.isFinite(salvage) || !Number.isFinite(life) || !Number.isFinite(period)) return ERRORS.VALUE;
      if (life <= 0 || period <= 0 || period > life + 1 || cost < 0 || salvage < 0) return ERRORS.NUM;
      const rate = 1 - Math.pow(salvage / cost, 1 / life);
      const roundedRate = Math.round(rate * 1000) / 1000;
      let totalDepreciation = 0;
      for (let p = 1; p <= period; p++) {
        let dep: number;
        if (p === 1) {
          dep = cost * roundedRate * month / 12;
        } else if (p === Math.ceil(life) + 1) {
          dep = (cost - totalDepreciation) * roundedRate * (12 - month) / 12;
        } else {
          dep = (cost - totalDepreciation) * roundedRate;
        }
        if (p === period) return dep;
        totalDepreciation += dep;
      }
      return ERRORS.VALUE;
    },
    DDB: async (args) => {
      const cost = Number(args[0]);
      const salvage = Number(args[1]);
      const life = Number(args[2]);
      const period = Number(args[3]);
      const factor = args[4] !== undefined ? Number(args[4]) : 2;
      if (!Number.isFinite(cost) || !Number.isFinite(salvage) || !Number.isFinite(life) || !Number.isFinite(period)) return ERRORS.VALUE;
      if (life <= 0 || period <= 0 || period > life || cost < 0 || salvage < 0) return ERRORS.NUM;
      let bookValue = cost;
      for (let p = 1; p <= period; p++) {
        const dep = Math.min(bookValue * factor / life, bookValue - salvage);
        if (p === period) return Math.max(dep, 0);
        bookValue -= dep;
      }
      return 0;
    },
    PPMT: async (args) => {
      const rate = Number(args[0]);
      const per = Number(args[1]);
      const nper = Number(args[2]);
      const pv = Number(args[3]);
      const fv = Number(args[4]) || 0;
      const type = Number(args[5]) || 0;
      if (!Number.isFinite(rate) || !Number.isFinite(per) || !Number.isFinite(nper) || !Number.isFinite(pv)) return ERRORS.VALUE;
      if (per < 1 || per > nper) return ERRORS.NUM;
      let pmt: number;
      if (rate === 0) {
        pmt = -(pv + fv) / nper;
      } else {
        const factor = Math.pow(1 + rate, nper);
        pmt = -(rate * (pv * factor + fv)) / ((1 + rate * type) * (factor - 1));
      }
      let ipmt: number;
      if (rate === 0) {
        ipmt = 0;
      } else {
        const balance = per === 1 ? pv : pv * Math.pow(1 + rate, per - 1) + pmt * (1 + rate * type) * (Math.pow(1 + rate, per - 1) - 1) / rate;
        ipmt = type === 1 && per === 1 ? 0 : balance * rate;
      }
      return pmt - ipmt;
    },
    IPMT: async (args) => {
      const rate = Number(args[0]);
      const per = Number(args[1]);
      const nper = Number(args[2]);
      const pv = Number(args[3]);
      const fv = Number(args[4]) || 0;
      const type = Number(args[5]) || 0;
      if (!Number.isFinite(rate) || !Number.isFinite(per) || !Number.isFinite(nper) || !Number.isFinite(pv)) return ERRORS.VALUE;
      if (per < 1 || per > nper) return ERRORS.NUM;
      if (rate === 0) return 0;
      let pmt: number;
      const factor = Math.pow(1 + rate, nper);
      pmt = -(rate * (pv * factor + fv)) / ((1 + rate * type) * (factor - 1));
      const balance = per === 1 ? pv : pv * Math.pow(1 + rate, per - 1) + pmt * (1 + rate * type) * (Math.pow(1 + rate, per - 1) - 1) / rate;
      return type === 1 && per === 1 ? 0 : balance * rate;
    },
    CUMIPMT: async (args) => {
      const rate = Number(args[0]);
      const nper = Number(args[1]);
      const pv = Number(args[2]);
      const startPeriod = Math.trunc(Number(args[3]));
      const endPeriod = Math.trunc(Number(args[4]));
      const type = Number(args[5]) || 0;
      if (!Number.isFinite(rate) || !Number.isFinite(nper) || !Number.isFinite(pv)) return ERRORS.VALUE;
      if (rate <= 0 || nper <= 0 || pv <= 0 || startPeriod < 1 || endPeriod < startPeriod || endPeriod > nper) return ERRORS.NUM;
      const factor = Math.pow(1 + rate, nper);
      const pmt = -(rate * (pv * factor)) / ((1 + rate * type) * (factor - 1));
      let totalInterest = 0;
      for (let per = startPeriod; per <= endPeriod; per++) {
        const balance = per === 1 ? pv : pv * Math.pow(1 + rate, per - 1) + pmt * (1 + rate * type) * (Math.pow(1 + rate, per - 1) - 1) / rate;
        const ipmt = type === 1 && per === 1 ? 0 : balance * rate;
        totalInterest += ipmt;
      }
      return totalInterest;
    },
    CUMPRINC: async (args) => {
      const rate = Number(args[0]);
      const nper = Number(args[1]);
      const pv = Number(args[2]);
      const startPeriod = Math.trunc(Number(args[3]));
      const endPeriod = Math.trunc(Number(args[4]));
      const type = Number(args[5]) || 0;
      if (!Number.isFinite(rate) || !Number.isFinite(nper) || !Number.isFinite(pv)) return ERRORS.VALUE;
      if (rate <= 0 || nper <= 0 || pv <= 0 || startPeriod < 1 || endPeriod < startPeriod || endPeriod > nper) return ERRORS.NUM;
      const factor = Math.pow(1 + rate, nper);
      const pmt = -(rate * (pv * factor)) / ((1 + rate * type) * (factor - 1));
      let totalPrincipal = 0;
      for (let per = startPeriod; per <= endPeriod; per++) {
        const balance = per === 1 ? pv : pv * Math.pow(1 + rate, per - 1) + pmt * (1 + rate * type) * (Math.pow(1 + rate, per - 1) - 1) / rate;
        const ipmt = type === 1 && per === 1 ? 0 : balance * rate;
        totalPrincipal += pmt - ipmt;
      }
      return totalPrincipal;
    },
    EFFECT: async (args) => {
      const nominal = Number(args[0]);
      const npery = Math.trunc(Number(args[1]));
      if (!Number.isFinite(nominal) || !Number.isFinite(npery)) return ERRORS.VALUE;
      if (nominal <= 0 || npery < 1) return ERRORS.NUM;
      return Math.pow(1 + nominal / npery, npery) - 1;
    },
    NOMINAL: async (args) => {
      const effectRate = Number(args[0]);
      const npery = Math.trunc(Number(args[1]));
      if (!Number.isFinite(effectRate) || !Number.isFinite(npery)) return ERRORS.VALUE;
      if (effectRate <= 0 || npery < 1) return ERRORS.NUM;
      return npery * (Math.pow(effectRate + 1, 1 / npery) - 1);
    },
    DOLLARDE: async (args) => {
      const fractionalDollar = Number(args[0]);
      const fraction = Math.trunc(Number(args[1]));
      if (!Number.isFinite(fractionalDollar) || !Number.isFinite(fraction)) return ERRORS.VALUE;
      if (fraction < 1) return ERRORS.DIV0;
      const intPart = Math.trunc(fractionalDollar);
      const fracPart = fractionalDollar - intPart;
      return intPart + fracPart * 10 / fraction;
    },
    DOLLARFR: async (args) => {
      const decimalDollar = Number(args[0]);
      const fraction = Math.trunc(Number(args[1]));
      if (!Number.isFinite(decimalDollar) || !Number.isFinite(fraction)) return ERRORS.VALUE;
      if (fraction < 1) return ERRORS.DIV0;
      const intPart = Math.trunc(decimalDollar);
      const fracPart = decimalDollar - intPart;
      return intPart + fracPart * fraction / 10;
    },
  };
}
