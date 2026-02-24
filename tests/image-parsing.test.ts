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

const twoImageDrawingXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:twoCellAnchor editAs="oneCell">
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>50800</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>165100</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>17</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>53</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:pic>
      <xdr:nvPicPr><xdr:cNvPr id="3" name="Picture 2"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>
      <xdr:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>
      <xdr:spPr><a:xfrm><a:off x="50800" y="165100"/><a:ext cx="14160500" cy="10756900"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
  <xdr:twoCellAnchor editAs="oneCell">
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>800100</xdr:colOff><xdr:row>56</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>15</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>78</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:pic>
      <xdr:nvPicPr><xdr:cNvPr id="9" name="Afbeelding 1"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>
      <xdr:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>
      <xdr:spPr><a:xfrm><a:off x="800100" y="11296650"/><a:ext cx="12601575" cy="4391025"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>
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

  it('parses two twoCellAnchor images (System model sheet)', () => {
    const drawingDoc = XmlParser.parseXml(twoImageDrawingXml);
    const images = (SheetParser as any).parseDrawing(drawingDoc);

    expect(images).toHaveLength(2);

    // First image: rows 0-53, cols 0-17 (0-based) → 1-based {1,1} to {18,54}
    expect(images[0].embed).toBe('rId1');
    expect(images[0].from).toEqual({ col: 1, row: 1 });
    expect(images[0].to).toEqual({ col: 18, row: 54 });
    expect(images[0].type).toBe('twoCellAnchor');

    // Second image: rows 56-78, cols 0-15
    expect(images[1].embed).toBe('rId2');
    expect(images[1].from).toEqual({ col: 1, row: 57 });
    expect(images[1].to).toEqual({ col: 16, row: 79 });
    expect(images[1].type).toBe('twoCellAnchor');
  });

  it('resolves embed IDs through drawing rels to get media paths', () => {
    // Simulate what loadSheetAssets now does: resolve embed via drawing rels
    const drawingRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
      <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image2.png"/>
    </Relationships>`;
    const relsDoc = XmlParser.parseXml(drawingRelsXml);
    const rels = XmlParser.buildRelationshipMap(relsDoc);

    expect(rels.get('rId1')).toBe('../media/image1.png');
    expect(rels.get('rId2')).toBe('../media/image2.png');

    // normalizeTargetPath resolves "../media/image1.png" → "xl/media/image1.png"
    expect(XmlParser.normalizeTargetPath('../media/image1.png')).toMatch(/media\/image1\.png$/);
  });
});
