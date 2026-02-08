// Conditional formatting style tests
// Run with: node --experimental-vm-modules tests/conditional-formatting-tests.js

import { dxfToCss } from '../src/js/styles.js';

function assertEqual(name, actual, expected) {
  const pass = JSON.stringify(actual) === JSON.stringify(expected);
  console.log(pass ? '✓' : '✗', name, pass ? '' : `\n  Expected: ${JSON.stringify(expected)}\n  Actual:   ${JSON.stringify(actual)}`);
  return pass;
}

async function runTests() {
  console.log('--- Conditional Formatting Style Tests ---\n');
  let passed = 0;
  let total = 0;

  total += 1;
  const dxf1 = {
    font: { bold: true, rgb: 'FFFF0000' },
    fill: { rgb: 'FF00FF00' },
    border: { top: { style: 'thin', rgb: 'FF0000FF' } }
  };
  const css1 = dxfToCss(dxf1, {});
  if (assertEqual('DXF to CSS with border', css1, {
    fontWeight: 'bold',
    color: '#FF0000',
    backgroundColor: '#00FF00',
    borderTop: '1px solid #0000FF'
  })) passed += 1;

  total += 1;
  const themeColors = { 2: '#123456', 5: '#abcdef' };
  const dxf2 = {
    font: { theme: '2' },
    fill: { theme: '5' },
    border: { bottom: { style: 'medium', theme: '2' } }
  };
  const css2 = dxfToCss(dxf2, themeColors);
  if (assertEqual('DXF to CSS with theme colors', css2, {
    color: '#123456',
    backgroundColor: '#abcdef',
    borderBottom: '2px solid #123456'
  })) passed += 1;

  console.log(`\n--- Results: ${passed}/${total} passed ---`);
}

runTests();
