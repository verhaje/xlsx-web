// Errors.ts - Shared error values for formula evaluation

export const ERRORS = {
  VALUE: '#VALUE!',
  REF: '#REF!',
  NAME: '#NAME?',
  DIV0: '#DIV/0!',
  NULL: '#NULL!',
  NUM: '#NUM!',
  NA: '#N/A',
  CYCLE: '#CYCLE!',
} as const;

export type ErrorValue = typeof ERRORS[keyof typeof ERRORS];

export function isError(value: any): value is ErrorValue {
  return typeof value === 'string' && value.startsWith('#');
}
