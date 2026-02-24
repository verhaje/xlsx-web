// XmlParser.ts - XML and ZIP text reading utilities

import type JSZip from 'jszip';

/**
 * Utility class for XML parsing and ZIP file reading operations.
 */
export class XmlParser {
  /**
   * Cache of decompressed text to avoid repeated jszip decompression of the
   * same entry (jszip creates many microtasks per decompression call).
   */
  private static textCache = new Map<string, string>();

  /**
   * Cache of parsed XML Documents to avoid redundant DOM parsing of the
   * same XML string.  Keyed by the ZIP path.
   */
  private static docCache = new Map<string, Document>();

  /**
   * In-flight decompression promises, coalescing concurrent requests for the
   * same path so jszip only decompresses once.
   */
  private static pendingReads = new Map<string, Promise<string>>();

  /**
   * Clear the decompressed-text cache (call when loading a new workbook).
   */
  static clearCache(): void {
    XmlParser.textCache.clear();
    XmlParser.docCache.clear();
    XmlParser.pendingReads.clear();
  }

  /**
   * Read a text file from a ZIP archive (cached).
   */
  static async readZipText(zip: JSZip, path: string): Promise<string> {
    // Cache hit – no jszip work needed
    if (XmlParser.textCache.has(path)) return XmlParser.textCache.get(path)!;

    // Coalesce concurrent reads for the same path
    if (XmlParser.pendingReads.has(path)) return XmlParser.pendingReads.get(path)!;

    const promise = (async () => {
      const entry = zip.file(path);
      if (!entry) throw new Error(`Missing ${path}`);
      const text = await entry.async('text');
      XmlParser.textCache.set(path, text);
      XmlParser.pendingReads.delete(path);
      return text;
    })();

    XmlParser.pendingReads.set(path, promise);
    return promise;
  }

  /**
   * Parse an XML string into a DOM Document.
   */
  static parseXml(xmlString: string): Document {
    const parser = new DOMParser();
    const doc = parser.parseFromString(xmlString, 'application/xml');
    const parserError = doc.getElementsByTagName('parsererror');
    if (parserError && parserError.length > 0) {
      throw new Error('Failed to parse XML.');
    }
    return doc;
  }

  /**
   * Read a ZIP entry and parse it as XML in one step, with Document caching.
   * Avoids re-parsing the same XML string into a new DOM tree.
   */
  static async readZipXml(zip: JSZip, path: string): Promise<Document> {
    const cached = XmlParser.docCache.get(path);
    if (cached) return cached;
    const text = await XmlParser.readZipText(zip, path);
    const doc = XmlParser.parseXml(text);
    XmlParser.docCache.set(path, doc);
    return doc;
  }

  /**
   * Evict a single path from the Document cache (e.g. after in-memory edits
   * invalidate the cached DOM).
   */
  static evictDocCache(path: string): void {
    XmlParser.docCache.delete(path);
  }

  /**
   * Build a Map of relationship IDs to target paths from a .rels document.
   */
  static buildRelationshipMap(relsDoc: Document): Map<string, string> {
    const rels = new Map<string, string>();
    const relationshipNodes = Array.from(relsDoc.getElementsByTagName('Relationship'));
    relationshipNodes.forEach((rel) => {
      const id = rel.getAttribute('Id');
      const target = rel.getAttribute('Target');
      if (id && target) rels.set(id, target);
    });
    return rels;
  }

  /**
   * Normalize a relationship target path to always start with `xl/`.
   * Handles relative targets like `../media/image1.png` from inside
   * `xl/` subdirectories (drawings, worksheets, etc.).
   */
  static normalizeTargetPath(target: string): string {
    if (target.startsWith('/')) return target.slice(1);
    if (target.startsWith('xl/')) return target;
    // Relative paths from xl/ sub-folders: strip ../ segments and prepend xl/
    let t = target;
    while (t.startsWith('../')) t = t.slice(3);
    return `xl/${t}`;
  }
}
