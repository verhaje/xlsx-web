// WorkbookLoader.ts - Loads workbook structure from a ZIP archive

import type JSZip from 'jszip';
import type { SheetInfo } from '../types';
import { XmlParser } from '../core/XmlParser';

/**
 * Loads workbook metadata: relationships, sheets, and shared strings.
 */
export class WorkbookLoader {
  private zip: JSZip;

  constructor(zip: JSZip) {
    this.zip = zip;
  }

  /**
   * Load the workbook relationship map.
   */
  async loadRelationshipMap(): Promise<Map<string, string>> {
    const relsXml = await XmlParser.readZipText(this.zip, 'xl/_rels/workbook.xml.rels');
    const relsDoc = XmlParser.parseXml(relsXml);
    return XmlParser.buildRelationshipMap(relsDoc);
  }

  /**
   * Load the list of sheets from the workbook.
   */
  async loadSheets(relMap: Map<string, string>): Promise<SheetInfo[]> {
    const workbookXml = await XmlParser.readZipText(this.zip, 'xl/workbook.xml');
    const workbookDoc = XmlParser.parseXml(workbookXml);

    return Array.from(workbookDoc.getElementsByTagName('sheet'))
      .map((sheet) => {
        const name = sheet.getAttribute('name') || 'Sheet';
        const relId = sheet.getAttribute('r:id') || '';
        const target = relMap.get(relId);
        return {
          name,
          relId,
          target: target ? XmlParser.normalizeTargetPath(target) : null,
        };
      })
      .filter((s): s is SheetInfo & { target: string } => s.target !== null);
  }

  /**
   * Load shared strings from the workbook.
   */
  async loadSharedStrings(): Promise<string[]> {
    const shared = this.zip.file('xl/sharedStrings.xml');
    if (!shared) return [];
    const sharedXml = await shared.async('text');
    const sharedDoc = XmlParser.parseXml(sharedXml);
    const items = Array.from(sharedDoc.getElementsByTagName('si'));
    return items.map((item) => {
      const texts = Array.from(item.getElementsByTagName('t'));
      return texts.map((t) => t.textContent || '').join('');
    });
  }
}
