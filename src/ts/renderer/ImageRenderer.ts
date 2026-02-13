// ImageRenderer.ts - Renders image overlays on the sheet table

import type { ImageAnchor } from '../types';

/**
 * Renders image overlays positioned over the table based on cell coordinates.
 */
export class ImageRenderer {
  /**
   * Create an absolutely-positioned image overlay container and append it to the table container.
   */
  static renderImages(tableContainer: HTMLElement, images: ImageAnchor[]): void {
    if (!images.length) return;

    const imageContainer = document.createElement('div');
    imageContainer.className = 'image-overlay';
    imageContainer.style.position = 'absolute';
    imageContainer.style.top = '0';
    imageContainer.style.left = '0';
    imageContainer.style.pointerEvents = 'none';
    imageContainer.style.zIndex = '10';

    images.forEach((image) => {
      const img = document.createElement('img');
      img.src = image.dataUrl || '';
      img.style.position = 'absolute';
      img.style.pointerEvents = 'auto';

      const startCol = image.from.col;
      const startRow = image.from.row;
      const endCol = image.to.col;
      const endRow = image.to.row;

      // Simplified positioning using default cell dimensions
      const colWidth = 64;
      const rowHeight = 20;

      const left = (startCol - 1) * colWidth;
      const top = (startRow - 1) * rowHeight;
      const width = (endCol - startCol + 1) * colWidth;
      const height = (endRow - startRow + 1) * rowHeight;

      img.style.left = `${left}px`;
      img.style.top = `${top}px`;
      img.style.width = `${width}px`;
      img.style.height = `${height}px`;
      img.style.objectFit = 'contain';

      imageContainer.appendChild(img);
    });

    tableContainer.style.position = 'relative';
    tableContainer.appendChild(imageContainer);
  }
}
