// TabRenderer.ts - Sheet tab rendering

export interface RenderTabsOptions<TSheet extends { name: string }> {
  sheets: TSheet[];
  activeIndex: number;
  /** @deprecated No longer used — scrolling is handled internally. */
  startIndex?: number;
  onSelect: (sheet: TSheet, index: number) => Promise<void>;
  /** @deprecated No longer used — scrolling is handled internally. */
  onNext?: () => void;
  /** @deprecated No longer used — scrolling is handled internally. */
  onPrev?: () => void;
  onAdd?: () => Promise<void> | void;
  onRename?: (sheet: TSheet, index: number, newName: string) => Promise<void> | void;
  onRemove?: (sheet: TSheet, index: number) => Promise<void> | void;
}

/**
 * Renders and manages sheet tab buttons.
 *
 * Desktop: all tabs are placed in a scrollable track (overflow hidden).
 * Nav buttons appear when the track overflows and scroll it left/right.
 *
 * Mobile (≤767 px): tabs collapse into a hamburger menu.
 */
export class TabRenderer {
  private static readonly SCROLL_AMOUNT = 200;
  private static readonly MOBILE_MEDIA_QUERY = '(max-width: 767px)';

  /**
   * Render tab buttons into a container element.
   */
  static renderTabs<TSheet extends { name: string }>(
    container: HTMLElement,
    options: RenderTabsOptions<TSheet>
  ): void {
    const {
      sheets,
      activeIndex,
      onSelect,
      onAdd,
      onRename,
      onRemove,
    } = options;

    container.replaceChildren();

    if (TabRenderer.isMobileView()) {
      TabRenderer.renderMobileTabs(container, options);
      return;
    }

    // ── Desktop: scrollable track containing ALL tab buttons ──

    const track = document.createElement('div');
    track.className = 'tabs-track';

    for (let index = 0; index < sheets.length; index += 1) {
      const sheet = sheets[index];
      const button = document.createElement('button');
      button.className = 'tab-button';
      if (index === activeIndex) button.classList.add('active');
      button.type = 'button';
      button.textContent = sheet.name;
      button.dataset.tabIndex = String(index);

      button.addEventListener('click', async () => {
        await onSelect(sheet, index);
      });

      if (onRemove || onRename) {
        button.addEventListener('contextmenu', (event) => {
          event.preventDefault();
          TabRenderer.showContextMenu(
            event.clientX,
            event.clientY,
            onRename
              ? async () => {
                  await TabRenderer.startInlineRename(button, sheet.name, async (newName) => {
                    await onRename(sheet, index, newName);
                  });
                }
              : undefined,
            onRemove
              ? async () => {
                  await onRemove(sheet, index);
                }
              : undefined
          );
        });
      }

      track.appendChild(button);
    }

    container.appendChild(track);

    // ── Controls: prev / next scroll buttons + add ──

    const controls = document.createElement('div');
    controls.className = 'tabs-controls';

    const prevBtn = document.createElement('button');
    prevBtn.type = 'button';
    prevBtn.className = 'tab-nav-button tab-nav-prev';
    prevBtn.textContent = '◀';
    prevBtn.title = 'Scroll tabs left';
    prevBtn.style.display = 'none';

    const nextBtn = document.createElement('button');
    nextBtn.type = 'button';
    nextBtn.className = 'tab-nav-button tab-nav-next';
    nextBtn.textContent = '▶';
    nextBtn.title = 'Scroll tabs right';
    nextBtn.style.display = 'none';

    prevBtn.addEventListener('click', () => {
      track.scrollBy({ left: -TabRenderer.SCROLL_AMOUNT, behavior: 'smooth' });
    });

    nextBtn.addEventListener('click', () => {
      track.scrollBy({ left: TabRenderer.SCROLL_AMOUNT, behavior: 'smooth' });
    });

    controls.appendChild(prevBtn);
    controls.appendChild(nextBtn);

    if (onAdd) {
      const addBtn = document.createElement('button');
      addBtn.type = 'button';
      addBtn.className = 'tab-add-button';
      addBtn.textContent = '+';
      addBtn.title = 'Add worksheet';
      addBtn.addEventListener('click', async () => {
        await onAdd();
      });
      controls.appendChild(addBtn);
    }

    container.appendChild(controls);

    // ── Show / hide nav buttons based on track overflow ──

    const updateNavButtons = (): void => {
      const hasOverflow = track.scrollWidth > track.clientWidth + 1;
      const canScrollLeft = track.scrollLeft > 0;
      const canScrollRight =
        track.scrollLeft + track.clientWidth < track.scrollWidth - 1;
      prevBtn.style.display = hasOverflow && canScrollLeft ? '' : 'none';
      nextBtn.style.display = hasOverflow && canScrollRight ? '' : 'none';
    };

    track.addEventListener('scroll', updateNavButtons);

    // After paint: scroll the active tab into view and refresh button state.
    if (typeof requestAnimationFrame === 'function') {
      requestAnimationFrame(() => {
        TabRenderer.scrollActiveTabIntoView(track);
        updateNavButtons();
      });
    }
  }

  /**
   * Scroll the `.tab-button.active` element into the visible area of the track
   * without causing the whole page to jump.
   */
  private static scrollActiveTabIntoView(track: HTMLElement): void {
    const activeButton = track.querySelector('.tab-button.active') as HTMLElement | null;
    if (!activeButton) return;
    const trackRect = track.getBoundingClientRect();
    const btnRect = activeButton.getBoundingClientRect();
    if (btnRect.left < trackRect.left) {
      track.scrollLeft -= trackRect.left - btnRect.left;
    } else if (btnRect.right > trackRect.right) {
      track.scrollLeft += btnRect.right - trackRect.right;
    }
  }

  static isMobileView(): boolean {
    return window.matchMedia(TabRenderer.MOBILE_MEDIA_QUERY).matches;
  }

  private static renderMobileTabs<TSheet extends { name: string }>(
    container: HTMLElement,
    options: RenderTabsOptions<TSheet>
  ): void {
    const {
      sheets,
      activeIndex,
      onSelect,
      onAdd,
      onRename,
      onRemove,
    } = options;

    const activeSheet = sheets[activeIndex];

    const mobileBar = document.createElement('div');
    mobileBar.className = 'tabs-mobile-bar';

    const activeBtn = document.createElement('button');
    activeBtn.type = 'button';
    activeBtn.className = 'tab-mobile-active';
    activeBtn.textContent = activeSheet?.name || 'Sheets';
    activeBtn.title = activeSheet?.name || 'Sheets';

    const menuToggle = document.createElement('button');
    menuToggle.type = 'button';
    menuToggle.className = 'tab-hamburger-button';
    menuToggle.textContent = '☰';
    menuToggle.title = 'Open sheets menu';
    menuToggle.setAttribute('aria-label', 'Open sheets menu');
    menuToggle.setAttribute('aria-expanded', 'false');

    mobileBar.appendChild(activeBtn);
    mobileBar.appendChild(menuToggle);

    if (onAdd) {
      const addBtn = document.createElement('button');
      addBtn.type = 'button';
      addBtn.className = 'tab-add-button';
      addBtn.textContent = '+';
      addBtn.title = 'Add worksheet';
      addBtn.addEventListener('click', async () => {
        await onAdd();
      });
      mobileBar.appendChild(addBtn);
    }

    const menu = document.createElement('div');
    menu.className = 'tab-hamburger-menu';

    const closeMenu = () => {
      menu.classList.remove('open');
      menuToggle.setAttribute('aria-expanded', 'false');
      document.removeEventListener('click', closeOnOutsideClick);
    };

    const openMenu = () => {
      menu.classList.add('open');
      menuToggle.setAttribute('aria-expanded', 'true');
      setTimeout(() => document.addEventListener('click', closeOnOutsideClick), 0);
    };

    const toggleMenu = () => {
      if (menu.classList.contains('open')) closeMenu();
      else openMenu();
    };

    menuToggle.addEventListener('click', toggleMenu);
    activeBtn.addEventListener('click', toggleMenu);

    const closeOnOutsideClick = (event: MouseEvent) => {
      if (container.contains(event.target as Node)) return;
      closeMenu();
    };

    for (let index = 0; index < sheets.length; index += 1) {
      const sheet = sheets[index];
      const item = document.createElement('button');
      item.type = 'button';
      item.className = 'tab-hamburger-item';
      if (index === activeIndex) item.classList.add('active');
      item.textContent = sheet.name;

      item.addEventListener('click', async () => {
        await onSelect(sheet, index);
        closeMenu();
      });

      if (onRemove || onRename) {
        item.addEventListener('contextmenu', (event) => {
          event.preventDefault();
          TabRenderer.showContextMenu(
            event.clientX,
            event.clientY,
            onRename
              ? async () => {
                  await TabRenderer.startInlineRename(item, sheet.name, async (newName) => {
                    await onRename(sheet, index, newName);
                  });
                }
              : undefined,
            onRemove
              ? async () => {
                  await onRemove(sheet, index);
                }
              : undefined
          );
        });
      }

      menu.appendChild(item);
    }

    container.appendChild(mobileBar);
    container.appendChild(menu);
  }

  /**
   * Set the active tab by index.
   */
  static setActiveTab(container: HTMLElement, index: number): void {
    const buttons = Array.from(container.querySelectorAll('.tab-button'));
    buttons.forEach((button) => {
      const tabIndexAttr = (button as HTMLElement).dataset.tabIndex;
      const tabIndex = tabIndexAttr ? Number(tabIndexAttr) : -1;
      if (tabIndex === index) button.classList.add('active');
      else button.classList.remove('active');
    });
  }

  private static showContextMenu(
    x: number,
    y: number,
    onRename?: () => Promise<void>,
    onRemove?: () => Promise<void>
  ): void {
    TabRenderer.hideContextMenu();

    const menu = document.createElement('div');
    menu.className = 'tab-context-menu';
    menu.style.left = `${x}px`;
    menu.style.top = `${y}px`;

    if (onRename) {
      const renameBtn = document.createElement('button');
      renameBtn.type = 'button';
      renameBtn.className = 'tab-context-menu-item';
      renameBtn.textContent = 'Rename sheet';
      renameBtn.addEventListener('click', async () => {
        await onRename();
        TabRenderer.hideContextMenu();
      });
      menu.appendChild(renameBtn);
    }

    if (onRemove) {
      const removeBtn = document.createElement('button');
      removeBtn.type = 'button';
      removeBtn.className = 'tab-context-menu-item danger';
      removeBtn.textContent = 'Remove sheet';
      removeBtn.addEventListener('click', async () => {
        await onRemove();
        TabRenderer.hideContextMenu();
      });
      menu.appendChild(removeBtn);
    }

    if (!onRename && !onRemove) return;
    document.body.appendChild(menu);

    const close = (event: Event) => {
      const target = event.target as Node | null;
      if (target && menu.contains(target)) return;
      TabRenderer.hideContextMenu();
      document.removeEventListener('click', close);
      document.removeEventListener('contextmenu', close);
      window.removeEventListener('resize', close);
      window.removeEventListener('scroll', close, true);
    };

    setTimeout(() => {
      document.addEventListener('click', close);
      document.addEventListener('contextmenu', close);
      window.addEventListener('resize', close);
      window.addEventListener('scroll', close, true);
    }, 0);
  }

  private static hideContextMenu(): void {
    document.querySelectorAll('.tab-context-menu').forEach((menu) => menu.remove());
  }

  private static async startInlineRename(
    button: HTMLButtonElement,
    currentName: string,
    onCommit: (newName: string) => Promise<void>
  ): Promise<void> {
    if (button.querySelector('.tab-rename-input')) return;

    const originalText = currentName;
    button.textContent = '';

    const input = document.createElement('input');
    input.type = 'text';
    input.className = 'tab-rename-input';
    input.value = originalText;
    input.setAttribute('aria-label', 'Rename worksheet');
    button.appendChild(input);

    const cancel = () => {
      button.textContent = originalText;
    };

    let finished = false;
    const commit = async () => {
      if (finished) return;
      finished = true;
      const newName = input.value.trim();
      if (!newName || newName === originalText) {
        cancel();
        return;
      }
      try {
        await onCommit(newName);
      } catch {
        cancel();
      }
    };

    input.addEventListener('keydown', async (event) => {
      if (event.key === 'Enter') {
        event.preventDefault();
        await commit();
      } else if (event.key === 'Escape') {
        event.preventDefault();
        finished = true;
        cancel();
      }
    });

    input.addEventListener('blur', async () => {
      await commit();
    });

    setTimeout(() => {
      input.focus();
      input.select();
    }, 0);
  }
}
