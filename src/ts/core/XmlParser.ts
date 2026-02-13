// XmlParser.ts - XML and ZIP text reading utilities

import type JSZip from 'jszip';

/**
 * Utility class for XML parsing and ZIP file reading operations.
 */
export class XmlParser {
  /**
   * Read a text file from a ZIP archive.
   */
  static async readZipText(zip: JSZip, path: string): Promise<string> {
    const entry = zip.file(path);
    if (!entry) {
      throw new Error(`Missing ${path}`);
    }
    return entry.async('text');
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
   */
  static normalizeTargetPath(target: string): string {
    if (target.startsWith('/')) return target.slice(1);
    if (target.startsWith('xl/')) return target;
    return `xl/${target}`;
  }
}
