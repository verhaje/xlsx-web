// Shared error values for formula evaluation
export const ERRORS = {
  VALUE: '#VALUE!',
  REF: '#REF!',
  NAME: '#NAME?',
  DIV0: '#DIV/0!',
  NULL: '#NULL!',
  NUM: '#NUM!',
  NA: '#N/A',
  CYCLE: '#CYCLE!',
};

export function isError(value) {
  return typeof value === 'string' && value.startsWith('#');
}
