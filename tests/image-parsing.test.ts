// @vitest-environment jsdom
import { describe, it, expect } from 'vitest';
import { XmlParser } from '../src/ts/core/XmlParser';
import { SheetParser } from '../src/ts/workbook/SheetParser';

const sampleDrawingXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:twoCellAnchor>
    <xdr:from>
      <xdr:col>0</xdr:col>
      <xdr:colOff>0</xdr:colOff>
      <xdr:row>0</xdr:row>
      <xdr:rowOff>0</xdr:rowOff>
    </xdr:from>
    <xdr:to>
      <xdr:col>4</xdr:col>
      <xdr:colOff>0</xdr:colOff>
      <xdr:row>9</xdr:row>
      <xdr:rowOff>0</xdr:rowOff>
    </xdr:to>
    <xdr:pic>
      <xdr:nvPicPr>
        <xdr:cNvPr id="1" name="Picture 1"/>
        <xdr:cNvPicPr>
          <a:picLocks noChangeAspect="1"/>
        </xdr:cNvPicPr>
      </xdr:nvPicPr>
      <xdr:blipFill>
        <a:blip r:embed="rId1">
          <a:extLst>
            <a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}">
              <a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" val="0"/>
            </a:ext>
          </a:extLst>
        </a:blip>
        <a:stretch>
          <a:fillRect/>
        </a:stretch>
      </xdr:blipFill>
      <xdr:spPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="3657600" cy="2743200"/>
        </a:xfrm>
        <a:prstGeom prst="rect">
          <a:avLst/>
        </a:prstGeom>
      </xdr:spPr>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>`;

describe('Image Parsing', () => {
  it('parses a twoCellAnchor drawing with one image', () => {
    const drawingDoc = XmlParser.parseXml(sampleDrawingXml);
    // parseDrawing is private, access it for testing:
    const images = (SheetParser as any).parseDrawing(drawingDoc);

    expect(images).toHaveLength(1);
    const img = images[0];
    expect(img.embed).toBe('rId1');
    expect(img.from).toEqual({ col: 1, row: 1 });
    expect(img.to).toEqual({ col: 5, row: 10 });
    expect(img.type).toBe('twoCellAnchor');
  });
});
