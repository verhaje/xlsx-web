// resize.js - attach column and row resizers to a table
export function attachResizers(table) {
  if (!table) return () => {};
  const disposers = [];
  const colgroup = table.querySelector('colgroup');

  // Column resizers
  const headerThs = Array.from(table.querySelectorAll('thead th'));
  // map header <th>s (skip corner) to <col> elements by index
  const dataHeaderThs = headerThs.slice(1);

  // Ensure a <colgroup> with a <col> for each data column exists so we can
  // set explicit pixel widths per-column. If missing, create one.
  let ensuredColgroup = colgroup;
  const headerCount = dataHeaderThs.length;
  if (!ensuredColgroup) {
    ensuredColgroup = document.createElement('colgroup');
    // insert as first child of table so browser uses it
    table.insertBefore(ensuredColgroup, table.firstChild);
  }
  // Ensure there are enough <col> elements
  const existingCols = Array.from(ensuredColgroup.querySelectorAll('col'));
  for (let i = existingCols.length; i < headerCount; i++) {
    const c = document.createElement('col');
    ensuredColgroup.appendChild(c);
    existingCols.push(c);
  }

  // Helper to add a DOM element and record cleanup
  function addHandle(parent, handle) {
    parent.appendChild(handle);
    return () => { if (handle.parentElement) handle.parentElement.removeChild(handle); };
  }

  // Prevent selection while dragging
  function setResizing(enabled) {
    if (enabled) document.body.classList.add('resizing'); else document.body.classList.remove('resizing');
  }

  const cols = colgroup ? Array.from(colgroup.querySelectorAll('col')) : [];

  dataHeaderThs.forEach((th, idx) => {
    const colEl = (cols[idx] || ensuredColgroup.querySelectorAll('col')[idx]) || null;
    const colIndex = th.dataset.colIndex ? Number(th.dataset.colIndex) : (idx + 1);
    const handle = document.createElement('div');
    handle.className = 'col-resize-handle';
    handle.title = 'Resize column';

    let pointerMove = null;
    let pointerUp = null;

    const onPointerDown = (e) => {
      e.stopPropagation();
      e.preventDefault();
      const startX = e.clientX;

      // Lock current widths for all columns to pixel values so changing one
      // column's width doesn't let the browser reflow and resize others.
      const allCols = Array.from(ensuredColgroup.querySelectorAll('col'));
      const startWidths = allCols.map((c, i) => {
        // prefer existing inline style, else computed width from header TH
        const w = c.style.width ? parseInt(c.style.width, 10) : null;
        if (w && !Number.isNaN(w)) return w;
        const thForIdx = dataHeaderThs[i];
        const measured = thForIdx ? Math.round(thForIdx.getBoundingClientRect().width) : 60;
        c.style.width = measured + 'px';
        return measured;
      });

      const startWidth = colEl ? startWidths[idx] : (th.getBoundingClientRect().width || 60);
      setResizing(true);
      handle.setPointerCapture && handle.setPointerCapture(e.pointerId);

      pointerMove = (ev) => {
        ev.preventDefault();
        const dx = ev.clientX - startX;
        const newW = Math.max(24, Math.round(startWidth + dx));
        if (colEl) colEl.style.width = newW + 'px';
      };

      pointerUp = (ev) => {
        ev.preventDefault();
        setResizing(false);
        handle.releasePointerCapture && handle.releasePointerCapture(ev.pointerId);
        document.removeEventListener('pointermove', pointerMove);
        document.removeEventListener('pointerup', pointerUp);
        pointerMove = null; pointerUp = null;
      };

      document.addEventListener('pointermove', pointerMove);
      document.addEventListener('pointerup', pointerUp);
    };

    handle.addEventListener('pointerdown', onPointerDown);
    disposers.push(() => handle.removeEventListener('pointerdown', onPointerDown));
    disposers.push(addHandle(th, handle));
  });

  // Row resizers: attach small handle to each row header
  const bodyRows = Array.from(table.querySelectorAll('tbody tr'));
  bodyRows.forEach((tr) => {
    const rowHeader = tr.querySelector('th');
    if (!rowHeader) return;
    const handle = document.createElement('div');
    handle.className = 'row-resize-handle';
    handle.title = 'Resize row';

    let pointerMove = null;
    let pointerUp = null;

    const onPointerDown = (e) => {
      e.stopPropagation();
      e.preventDefault();
      const startY = e.clientY;
      const startH = tr.getBoundingClientRect().height || 24;
      setResizing(true);
      handle.setPointerCapture && handle.setPointerCapture(e.pointerId);

      pointerMove = (ev) => {
        ev.preventDefault();
        const dy = ev.clientY - startY;
        const newH = Math.max(18, Math.round(startH + dy));
        tr.style.height = newH + 'px';
      };

      pointerUp = (ev) => {
        ev.preventDefault();
        setResizing(false);
        handle.releasePointerCapture && handle.releasePointerCapture(ev.pointerId);
        document.removeEventListener('pointermove', pointerMove);
        document.removeEventListener('pointerup', pointerUp);
        pointerMove = null; pointerUp = null;
      };

      document.addEventListener('pointermove', pointerMove);
      document.addEventListener('pointerup', pointerUp);
    };

    handle.addEventListener('pointerdown', onPointerDown);
    disposers.push(() => handle.removeEventListener('pointerdown', onPointerDown));
    disposers.push(addHandle(rowHeader, handle));
  });

  // Return disposer to clean up handles/listeners
  return () => {
    disposers.forEach((d) => { try { d(); } catch (e) {} });
    setResizing(false);
  };
}

export default attachResizers;
