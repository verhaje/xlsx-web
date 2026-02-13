// TabRenderer.ts - Sheet tab rendering

/**
 * Renders and manages sheet tab buttons.
 */
export class TabRenderer {
  /**
   * Render tab buttons into a container element.
   */
  static renderTabs(
    container: HTMLElement,
    sheets: Array<{ name: string }>,
    onSelect: (sheet: any, index: number) => Promise<void>
  ): void {
    container.innerHTML = '';
    sheets.forEach((sheet, index) => {
      const button = document.createElement('button');
      button.className = 'tab-button';
      button.type = 'button';
      button.textContent = sheet.name;
      button.addEventListener('click', async () => {
        TabRenderer.setActiveTab(container, index);
        await onSelect(sheet, index);
      });
      container.appendChild(button);
    });
  }

  /**
   * Set the active tab by index.
   */
  static setActiveTab(container: HTMLElement, index: number): void {
    const buttons = Array.from(container.querySelectorAll('.tab-button'));
    buttons.forEach((button, idx) => {
      if (idx === index) button.classList.add('active');
      else button.classList.remove('active');
    });
  }
}
