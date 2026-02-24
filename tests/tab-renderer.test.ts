// @vitest-environment jsdom
import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { TabRenderer } from '../src/ts/renderer/TabRenderer';

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function mockMatchMedia(matches: boolean) {
  Object.defineProperty(window, 'matchMedia', {
    writable: true,
    value: vi.fn().mockImplementation((query: string) => ({
      matches,
      media: query,
      onchange: null,
      addListener: vi.fn(),
      removeListener: vi.fn(),
      addEventListener: vi.fn(),
      removeEventListener: vi.fn(),
      dispatchEvent: vi.fn(),
    })),
  });
}

const sheets = [
  { name: 'Sheet1' },
  { name: 'Sheet2' },
  { name: 'Sheet3' },
  { name: 'Financials' },
  { name: 'Summary' },
];

// ---------------------------------------------------------------------------
// Desktop tests
// ---------------------------------------------------------------------------

describe('TabRenderer – Desktop', () => {
  let container: HTMLElement;

  beforeEach(() => {
    mockMatchMedia(false); // not mobile
    container = document.createElement('nav');
    container.className = 'tabs';
    document.body.appendChild(container);
  });

  afterEach(() => {
    container.remove();
    vi.restoreAllMocks();
  });

  it('renders ALL tab buttons in the track (no pagination)', () => {
    TabRenderer.renderTabs(container, {
      sheets,
      activeIndex: 0,
      onSelect: vi.fn().mockResolvedValue(undefined),
    });

    const track = container.querySelector('.tabs-track')!;
    expect(track).toBeTruthy();
    const buttons = track.querySelectorAll('.tab-button');
    expect(buttons.length).toBe(5);
    expect(buttons[0].textContent).toBe('Sheet1');
    expect(buttons[4].textContent).toBe('Summary');
  });

  it('marks only the active tab with .active', () => {
    TabRenderer.renderTabs(container, {
      sheets,
      activeIndex: 2,
      onSelect: vi.fn().mockResolvedValue(undefined),
    });

    const buttons = container.querySelectorAll('.tab-button');
    for (let i = 0; i < buttons.length; i++) {
      if (i === 2) expect(buttons[i].classList.contains('active')).toBe(true);
      else expect(buttons[i].classList.contains('active')).toBe(false);
    }
  });

  it('creates prev and next nav buttons', () => {
    TabRenderer.renderTabs(container, {
      sheets,
      activeIndex: 0,
      onSelect: vi.fn().mockResolvedValue(undefined),
    });

    const prevBtn = container.querySelector('.tab-nav-prev');
    const nextBtn = container.querySelector('.tab-nav-next');
    expect(prevBtn).toBeTruthy();
    expect(nextBtn).toBeTruthy();
  });

  it('renders add button when onAdd is provided', () => {
    TabRenderer.renderTabs(container, {
      sheets,
      activeIndex: 0,
      onSelect: vi.fn().mockResolvedValue(undefined),
      onAdd: vi.fn(),
    });

    const addBtn = container.querySelector('.tab-add-button');
    expect(addBtn).toBeTruthy();
    expect(addBtn!.textContent).toBe('+');
  });

  it('does not render add button when onAdd is omitted', () => {
    TabRenderer.renderTabs(container, {
      sheets,
      activeIndex: 0,
      onSelect: vi.fn().mockResolvedValue(undefined),
    });

    expect(container.querySelector('.tab-add-button')).toBeNull();
  });

  it('calls onSelect when a tab button is clicked', async () => {
    const onSelect = vi.fn().mockResolvedValue(undefined);
    TabRenderer.renderTabs(container, { sheets, activeIndex: 0, onSelect });

    const buttons = container.querySelectorAll('.tab-button');
    (buttons[1] as HTMLButtonElement).click();

    await vi.waitFor(() => expect(onSelect).toHaveBeenCalledWith(sheets[1], 1));
  });

  it('calls onAdd when the add button is clicked', async () => {
    const onAdd = vi.fn();
    TabRenderer.renderTabs(container, {
      sheets,
      activeIndex: 0,
      onSelect: vi.fn().mockResolvedValue(undefined),
      onAdd,
    });

    (container.querySelector('.tab-add-button') as HTMLButtonElement).click();
    await vi.waitFor(() => expect(onAdd).toHaveBeenCalled());
  });

  it('does NOT render mobile bar or hamburger menu on desktop', () => {
    TabRenderer.renderTabs(container, {
      sheets,
      activeIndex: 0,
      onSelect: vi.fn().mockResolvedValue(undefined),
    });

    expect(container.querySelector('.tabs-mobile-bar')).toBeNull();
    expect(container.querySelector('.tab-hamburger-button')).toBeNull();
    expect(container.querySelector('.tab-hamburger-menu')).toBeNull();
  });

  it('stores data-tab-index on each button', () => {
    TabRenderer.renderTabs(container, {
      sheets,
      activeIndex: 0,
      onSelect: vi.fn().mockResolvedValue(undefined),
    });

    const buttons = container.querySelectorAll('.tab-button');
    buttons.forEach((btn, i) => {
      expect((btn as HTMLElement).dataset.tabIndex).toBe(String(i));
    });
  });
});

// ---------------------------------------------------------------------------
// Mobile tests
// ---------------------------------------------------------------------------

describe('TabRenderer – Mobile', () => {
  let container: HTMLElement;

  beforeEach(() => {
    mockMatchMedia(true); // mobile
    container = document.createElement('nav');
    container.className = 'tabs';
    document.body.appendChild(container);
  });

  afterEach(() => {
    container.remove();
    vi.restoreAllMocks();
  });

  it('renders hamburger button with ☰ icon', () => {
    TabRenderer.renderTabs(container, {
      sheets,
      activeIndex: 0,
      onSelect: vi.fn().mockResolvedValue(undefined),
    });

    const hamburger = container.querySelector('.tab-hamburger-button');
    expect(hamburger).toBeTruthy();
    expect(hamburger!.textContent).toBe('☰');
  });

  it('renders active sheet name in the mobile bar', () => {
    TabRenderer.renderTabs(container, {
      sheets,
      activeIndex: 3,
      onSelect: vi.fn().mockResolvedValue(undefined),
    });

    const activeBtn = container.querySelector('.tab-mobile-active');
    expect(activeBtn).toBeTruthy();
    expect(activeBtn!.textContent).toBe('Financials');
  });

  it('hamburger menu is initially closed (no .open class)', () => {
    TabRenderer.renderTabs(container, {
      sheets,
      activeIndex: 0,
      onSelect: vi.fn().mockResolvedValue(undefined),
    });

    const menu = container.querySelector('.tab-hamburger-menu');
    expect(menu).toBeTruthy();
    expect(menu!.classList.contains('open')).toBe(false);
  });

  it('clicking hamburger button opens the menu', () => {
    TabRenderer.renderTabs(container, {
      sheets,
      activeIndex: 0,
      onSelect: vi.fn().mockResolvedValue(undefined),
    });

    const hamburger = container.querySelector('.tab-hamburger-button') as HTMLButtonElement;
    hamburger.click();

    const menu = container.querySelector('.tab-hamburger-menu')!;
    expect(menu.classList.contains('open')).toBe(true);
  });

  it('clicking hamburger again closes the menu', () => {
    TabRenderer.renderTabs(container, {
      sheets,
      activeIndex: 0,
      onSelect: vi.fn().mockResolvedValue(undefined),
    });

    const hamburger = container.querySelector('.tab-hamburger-button') as HTMLButtonElement;
    hamburger.click(); // open
    hamburger.click(); // close

    const menu = container.querySelector('.tab-hamburger-menu')!;
    expect(menu.classList.contains('open')).toBe(false);
  });

  it('clicking active-sheet button also toggles the menu', () => {
    TabRenderer.renderTabs(container, {
      sheets,
      activeIndex: 0,
      onSelect: vi.fn().mockResolvedValue(undefined),
    });

    const activeBtn = container.querySelector('.tab-mobile-active') as HTMLButtonElement;
    activeBtn.click();

    const menu = container.querySelector('.tab-hamburger-menu')!;
    expect(menu.classList.contains('open')).toBe(true);

    activeBtn.click();
    expect(menu.classList.contains('open')).toBe(false);
  });

  it('menu lists all sheet names', () => {
    TabRenderer.renderTabs(container, {
      sheets,
      activeIndex: 0,
      onSelect: vi.fn().mockResolvedValue(undefined),
    });

    const items = container.querySelectorAll('.tab-hamburger-item');
    expect(items.length).toBe(5);
    expect(items[0].textContent).toBe('Sheet1');
    expect(items[4].textContent).toBe('Summary');
  });

  it('active sheet item has .active class', () => {
    TabRenderer.renderTabs(container, {
      sheets,
      activeIndex: 2,
      onSelect: vi.fn().mockResolvedValue(undefined),
    });

    const items = container.querySelectorAll('.tab-hamburger-item');
    expect(items[2].classList.contains('active')).toBe(true);
    expect(items[0].classList.contains('active')).toBe(false);
  });

  it('clicking a menu item calls onSelect and closes menu', async () => {
    const onSelect = vi.fn().mockResolvedValue(undefined);
    TabRenderer.renderTabs(container, { sheets, activeIndex: 0, onSelect });

    // open menu first
    (container.querySelector('.tab-hamburger-button') as HTMLButtonElement).click();

    const items = container.querySelectorAll('.tab-hamburger-item');
    (items[2] as HTMLButtonElement).click();

    await vi.waitFor(() => expect(onSelect).toHaveBeenCalledWith(sheets[2], 2));

    const menu = container.querySelector('.tab-hamburger-menu')!;
    expect(menu.classList.contains('open')).toBe(false);
  });

  it('does NOT render desktop tabs-track on mobile', () => {
    TabRenderer.renderTabs(container, {
      sheets,
      activeIndex: 0,
      onSelect: vi.fn().mockResolvedValue(undefined),
    });

    expect(container.querySelector('.tabs-track')).toBeNull();
    expect(container.querySelectorAll('.tab-button').length).toBe(0);
  });

  it('renders add button in mobile bar when onAdd is provided', () => {
    TabRenderer.renderTabs(container, {
      sheets,
      activeIndex: 0,
      onSelect: vi.fn().mockResolvedValue(undefined),
      onAdd: vi.fn(),
    });

    const addBtn = container.querySelector('.tab-add-button');
    expect(addBtn).toBeTruthy();
  });

  it('aria-expanded reflects menu state', () => {
    TabRenderer.renderTabs(container, {
      sheets,
      activeIndex: 0,
      onSelect: vi.fn().mockResolvedValue(undefined),
    });

    const hamburger = container.querySelector('.tab-hamburger-button') as HTMLButtonElement;
    expect(hamburger.getAttribute('aria-expanded')).toBe('false');

    hamburger.click();
    expect(hamburger.getAttribute('aria-expanded')).toBe('true');

    hamburger.click();
    expect(hamburger.getAttribute('aria-expanded')).toBe('false');
  });
});

// ---------------------------------------------------------------------------
// setActiveTab
// ---------------------------------------------------------------------------

describe('TabRenderer.setActiveTab', () => {
  it('updates the .active class on the correct button', () => {
    mockMatchMedia(false);
    const container = document.createElement('nav');
    document.body.appendChild(container);

    TabRenderer.renderTabs(container, {
      sheets,
      activeIndex: 0,
      onSelect: vi.fn().mockResolvedValue(undefined),
    });

    TabRenderer.setActiveTab(container, 3);

    const buttons = container.querySelectorAll('.tab-button');
    expect(buttons[0].classList.contains('active')).toBe(false);
    expect(buttons[3].classList.contains('active')).toBe(true);

    container.remove();
  });
});

// ---------------------------------------------------------------------------
// isMobileView helper
// ---------------------------------------------------------------------------

describe('TabRenderer.isMobileView', () => {
  afterEach(() => vi.restoreAllMocks());

  it('returns true when matchMedia matches', () => {
    mockMatchMedia(true);
    expect(TabRenderer.isMobileView()).toBe(true);
  });

  it('returns false when matchMedia does not match', () => {
    mockMatchMedia(false);
    expect(TabRenderer.isMobileView()).toBe(false);
  });
});
