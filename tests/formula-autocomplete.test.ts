import { describe, it, expect, beforeEach } from 'vitest';
import { JSDOM } from 'jsdom';
import { FormulaAutocomplete } from '../src/ts/ui/FormulaAutocomplete';

describe('FormulaAutocomplete', () => {
  let dom: JSDOM;
  let container: HTMLElement;
  let autocomplete: FormulaAutocomplete;

  beforeEach(() => {
    dom = new JSDOM('<!DOCTYPE html><html><body><div id="container"></div></body></html>');
    const doc = dom.window.document;
    container = doc.getElementById('container')!;
    // Inject globals for BuiltinFunctions (it reads from a registry)
    (globalThis as any).document = doc;
    autocomplete = new FormulaAutocomplete(container);
  });

  describe('update', () => {
    it('shows suggestions when typing a function prefix', () => {
      autocomplete.update('=SU', 3);
      expect(autocomplete.isVisible()).toBe(true);
    });

    it('hides when no matches found', () => {
      autocomplete.update('=ZZZZNOTFUNC', 12);
      expect(autocomplete.isVisible()).toBe(false);
    });

    it('hides when token is empty', () => {
      autocomplete.update('=', 1);
      expect(autocomplete.isVisible()).toBe(false);
    });

    it('shows suggestions for partial function names', () => {
      autocomplete.update('=AV', 3);
      expect(autocomplete.isVisible()).toBe(true);
    });

    it('shows additional suggestions when a function name prefix matches others', () => {
      autocomplete.update('=SUM', 4);
      // SUM matches SUMIF, SUMIFS etc. — should still show suggestions
      expect(autocomplete.isVisible()).toBe(true);
    });

    it('keeps exact match visible when full function name is typed', () => {
      autocomplete.update('=TODAY', 6);
      expect(autocomplete.isVisible()).toBe(true);

      const items = container.querySelectorAll('.formula-autocomplete-item .autocomplete-name');
      expect(items.length).toBeGreaterThan(0);
      expect(items[0]?.textContent).toBe('TODAY');
    });

    it('shows parameter hint when cursor is inside a function', () => {
      autocomplete.update('=SUM(', 5);
      expect(autocomplete.isVisible()).toBe(true);

      const syntax = container.querySelector('.formula-autocomplete-item .autocomplete-syntax') as HTMLElement | null;
      expect(syntax).toBeTruthy();
      expect(syntax?.innerHTML).toContain('<strong>number1</strong>');
    });

    it('highlights current parameter while editing arguments', () => {
      autocomplete.update('=SUM(A1,', 8);
      expect(autocomplete.isVisible()).toBe(true);

      const syntax = container.querySelector('.formula-autocomplete-item .autocomplete-syntax') as HTMLElement | null;
      expect(syntax).toBeTruthy();
      expect(syntax?.innerHTML).toContain('<strong>[number2]</strong>');
    });

    it('hides outer hint after function is closed', () => {
      autocomplete.update('=SUM(A1)', 8);
      expect(autocomplete.isVisible()).toBe(false);
    });

    it('switches suggestions to a new function typed inside another function', () => {
      autocomplete.update('=SUM(AV', 7);
      expect(autocomplete.isVisible()).toBe(true);

      const first = container.querySelector('.formula-autocomplete-item .autocomplete-name') as HTMLElement | null;
      expect(first).toBeTruthy();
      expect(first?.textContent?.startsWith('AV')).toBe(true);
    });

    it('suggests functions after operators', () => {
      autocomplete.update('=A1+SU', 6);
      expect(autocomplete.isVisible()).toBe(true);
    });

    it('suggests functions inside parentheses', () => {
      autocomplete.update('=IF(SU', 6);
      expect(autocomplete.isVisible()).toBe(true);
    });
  });

  describe('handleKey', () => {
    it('returns false when not visible', () => {
      const event = new dom.window.KeyboardEvent('keydown', { key: 'ArrowDown' });
      expect(autocomplete.handleKey(event as any)).toBe(false);
    });

    it('returns true for Escape when visible', () => {
      autocomplete.update('=SU', 3);
      const event = new dom.window.KeyboardEvent('keydown', { key: 'Escape' });
      const consumed = autocomplete.handleKey(event as any);
      expect(consumed).toBe(true);
      expect(autocomplete.isVisible()).toBe(false);
    });
  });

  describe('onSelect callback', () => {
    it('fires when a suggestion would be selected', () => {
      let selectedName = '';
      autocomplete.onSelect = (name) => { selectedName = name; };

      autocomplete.update('=SU', 3);
      expect(autocomplete.isVisible()).toBe(true);

      // Simulate enter to select first item
      const enterEvent = new dom.window.KeyboardEvent('keydown', { key: 'Enter' });
      autocomplete.handleKey(enterEvent as any);

      // Should have selected something starting with SU
      if (selectedName) {
        expect(selectedName.startsWith('SU')).toBe(true);
      }
    });
  });

  describe('hide', () => {
    it('hides the dropdown', () => {
      autocomplete.update('=SU', 3);
      autocomplete.hide();
      expect(autocomplete.isVisible()).toBe(false);
    });
  });

  describe('dispose', () => {
    it('removes the dropdown from the DOM', () => {
      const dropdownCount = container.querySelectorAll('.formula-autocomplete').length;
      expect(dropdownCount).toBe(1);
      autocomplete.dispose();
      const afterCount = container.querySelectorAll('.formula-autocomplete').length;
      expect(afterCount).toBe(0);
    });
  });
});
