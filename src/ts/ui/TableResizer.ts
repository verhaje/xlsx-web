// TableResizer.ts - Attach column and row resize handles to a rendered table

/**
 * Attaches column and row resize handles to a sheet table.
 * Returns a disposer function to clean up handles and listeners.
 */
export class TableResizer {
  /**
   * Attach resizers to the given table element.
   * @returns A disposer function that removes all handles and listeners.
   */
  static attach(table: HTMLTableElement): () => void {
    if (!table) return () => {};

    const disposers: Array<() => void> = [];
    const colgroup = table.querySelector('colgroup');

    // Column resizers
    const headerThs = Array.from(table.querySelectorAll('thead th')) as HTMLTableCellElement[];
    const dataHeaderThs = headerThs.slice(1);

    // Ensure a <colgroup> with enough <col> elements exists
    let ensuredColgroup = colgroup;
    const headerCount = dataHeaderThs.length;
    if (!ensuredColgroup) {
      ensuredColgroup = document.createElement('colgroup');
      table.insertBefore(ensuredColgroup, table.firstChild);
    }
    const existingCols = Array.from(ensuredColgroup.querySelectorAll('col')) as HTMLElement[];
    for (let i = existingCols.length; i < headerCount; i++) {
      const c = document.createElement('col');
      ensuredColgroup.appendChild(c);
      existingCols.push(c);
    }

    const addHandle = (parent: HTMLElement, handle: HTMLElement): (() => void) => {
      parent.appendChild(handle);
      return () => {
        if (handle.parentElement) handle.parentElement.removeChild(handle);
      };
    };

    const setResizing = (enabled: boolean): void => {
      if (enabled) document.body.classList.add('resizing');
      else document.body.classList.remove('resizing');
    };

    const cols = colgroup ? Array.from(colgroup.querySelectorAll('col')) as HTMLElement[] : [];

    // Column resize handles
    // cols may include a leading <col> for the row-header column, so data
    // columns start at offset 1 when such a col is present.
    const dataColOffset = (cols.length > 0 && !cols[0].dataset.colIndex) ? 1 : 0;

    dataHeaderThs.forEach((th, idx) => {
      const colIdx = idx + dataColOffset;
      const colEl = (cols[colIdx] || ensuredColgroup!.querySelectorAll('col')[colIdx] || null) as HTMLElement | null;
      const handle = document.createElement('div');
      handle.className = 'col-resize-handle';
      handle.title = 'Resize column';

      let pointerMove: ((ev: PointerEvent) => void) | null = null;
      let pointerUp: ((ev: PointerEvent) => void) | null = null;

      const onPointerDown = (e: PointerEvent): void => {
        e.stopPropagation();
        e.preventDefault();
        const startX = e.clientX;

        // Lock all column widths to pixels to prevent reflow
        const allCols = Array.from(ensuredColgroup!.querySelectorAll('col')) as HTMLElement[];
        const startWidths = allCols.map((c, i) => {
          const w = c.style.width ? parseInt(c.style.width, 10) : null;
          if (w && !Number.isNaN(w)) return w;
          const thForIdx = dataHeaderThs[i - dataColOffset];
          const measured = thForIdx ? Math.round(thForIdx.getBoundingClientRect().width) : 60;
          c.style.width = measured + 'px';
          return measured;
        });

        const startWidth = colEl ? startWidths[colIdx] : (th.getBoundingClientRect().width || 60);
        setResizing(true);
        if (handle.setPointerCapture) handle.setPointerCapture(e.pointerId);

        pointerMove = (ev: PointerEvent) => {
          ev.preventDefault();
          const dx = ev.clientX - startX;
          const newW = Math.max(24, Math.round(startWidth + dx));
          if (colEl) colEl.style.width = newW + 'px';
        };

        pointerUp = (ev: PointerEvent) => {
          ev.preventDefault();
          setResizing(false);
          if (handle.releasePointerCapture) handle.releasePointerCapture(ev.pointerId);
          document.removeEventListener('pointermove', pointerMove as EventListener);
          document.removeEventListener('pointerup', pointerUp as EventListener);
          pointerMove = null;
          pointerUp = null;
        };

        document.addEventListener('pointermove', pointerMove as EventListener);
        document.addEventListener('pointerup', pointerUp as EventListener);
      };

      handle.addEventListener('pointerdown', onPointerDown as EventListener);
      disposers.push(() => handle.removeEventListener('pointerdown', onPointerDown as EventListener));
      disposers.push(addHandle(th, handle));
    });

    // Row resize handles
    const bodyRows = Array.from(table.querySelectorAll('tbody tr')) as HTMLTableRowElement[];
    bodyRows.forEach((tr) => {
      const rowHeader = tr.querySelector('th') as HTMLElement | null;
      if (!rowHeader) return;
      const handle = document.createElement('div');
      handle.className = 'row-resize-handle';
      handle.title = 'Resize row';

      let pointerMove: ((ev: PointerEvent) => void) | null = null;
      let pointerUp: ((ev: PointerEvent) => void) | null = null;

      const onPointerDown = (e: PointerEvent): void => {
        e.stopPropagation();
        e.preventDefault();
        const startY = e.clientY;
        const startH = tr.getBoundingClientRect().height || 24;
        setResizing(true);
        if (handle.setPointerCapture) handle.setPointerCapture(e.pointerId);

        pointerMove = (ev: PointerEvent) => {
          ev.preventDefault();
          const dy = ev.clientY - startY;
          const newH = Math.max(18, Math.round(startH + dy));
          tr.style.height = newH + 'px';
        };

        pointerUp = (ev: PointerEvent) => {
          ev.preventDefault();
          setResizing(false);
          if (handle.releasePointerCapture) handle.releasePointerCapture(ev.pointerId);
          document.removeEventListener('pointermove', pointerMove as EventListener);
          document.removeEventListener('pointerup', pointerUp as EventListener);
          pointerMove = null;
          pointerUp = null;
        };

        document.addEventListener('pointermove', pointerMove as EventListener);
        document.addEventListener('pointerup', pointerUp as EventListener);
      };

      handle.addEventListener('pointerdown', onPointerDown as EventListener);
      disposers.push(() => handle.removeEventListener('pointerdown', onPointerDown as EventListener));
      disposers.push(addHandle(rowHeader, handle));
    });

    // Return disposer
    return () => {
      disposers.forEach((d) => {
        try { d(); } catch {}
      });
      setResizing(false);
    };
  }
}
