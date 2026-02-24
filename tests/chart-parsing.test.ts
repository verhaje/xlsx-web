// @vitest-environment jsdom
import { describe, it, expect } from 'vitest';
import { XmlParser } from '../src/ts/core/XmlParser';
import { ChartParser } from '../src/ts/workbook/ChartParser';

// ---- Sample Chart XMLs ----

const barChartXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart>
    <c:title>
      <c:tx>
        <c:rich>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:r>
              <a:rPr lang="en-US"/>
              <a:t>Sales by Region</a:t>
            </a:r>
          </a:p>
        </c:rich>
      </c:tx>
    </c:title>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:ser>
          <c:idx val="0"/>
          <c:tx>
            <c:strRef>
              <c:f>Sheet1!$B$1</c:f>
              <c:strCache>
                <c:ptCount val="1"/>
                <c:pt idx="0"><c:v>Revenue</c:v></c:pt>
              </c:strCache>
            </c:strRef>
          </c:tx>
          <c:cat>
            <c:strRef>
              <c:f>Sheet1!$A$2:$A$5</c:f>
              <c:strCache>
                <c:ptCount val="4"/>
                <c:pt idx="0"><c:v>North</c:v></c:pt>
                <c:pt idx="1"><c:v>South</c:v></c:pt>
                <c:pt idx="2"><c:v>East</c:v></c:pt>
                <c:pt idx="3"><c:v>West</c:v></c:pt>
              </c:strCache>
            </c:strRef>
          </c:cat>
          <c:val>
            <c:numRef>
              <c:f>Sheet1!$B$2:$B$5</c:f>
              <c:numCache>
                <c:formatCode>General</c:formatCode>
                <c:ptCount val="4"/>
                <c:pt idx="0"><c:v>100</c:v></c:pt>
                <c:pt idx="1"><c:v>200</c:v></c:pt>
                <c:pt idx="2"><c:v>150</c:v></c:pt>
                <c:pt idx="3"><c:v>250</c:v></c:pt>
              </c:numCache>
            </c:numRef>
          </c:val>
        </c:ser>
        <c:ser>
          <c:idx val="1"/>
          <c:tx>
            <c:strRef>
              <c:f>Sheet1!$C$1</c:f>
              <c:strCache>
                <c:ptCount val="1"/>
                <c:pt idx="0"><c:v>Costs</c:v></c:pt>
              </c:strCache>
            </c:strRef>
          </c:tx>
          <c:cat>
            <c:strRef>
              <c:f>Sheet1!$A$2:$A$5</c:f>
              <c:strCache>
                <c:ptCount val="4"/>
                <c:pt idx="0"><c:v>North</c:v></c:pt>
                <c:pt idx="1"><c:v>South</c:v></c:pt>
                <c:pt idx="2"><c:v>East</c:v></c:pt>
                <c:pt idx="3"><c:v>West</c:v></c:pt>
              </c:strCache>
            </c:strRef>
          </c:cat>
          <c:val>
            <c:numRef>
              <c:f>Sheet1!$C$2:$C$5</c:f>
              <c:numCache>
                <c:formatCode>General</c:formatCode>
                <c:ptCount val="4"/>
                <c:pt idx="0"><c:v>80</c:v></c:pt>
                <c:pt idx="1"><c:v>120</c:v></c:pt>
                <c:pt idx="2"><c:v>90</c:v></c:pt>
                <c:pt idx="3"><c:v>180</c:v></c:pt>
              </c:numCache>
            </c:numRef>
          </c:val>
        </c:ser>
      </c:barChart>
      <c:catAx>
        <c:axId val="1"/>
        <c:title>
          <c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Region</a:t></a:r></a:p></c:rich></c:tx>
        </c:title>
      </c:catAx>
      <c:valAx>
        <c:axId val="2"/>
        <c:title>
          <c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Amount</a:t></a:r></a:p></c:rich></c:tx>
        </c:title>
      </c:valAx>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;

const pieChartXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:title>
      <c:tx>
        <c:rich>
          <a:bodyPr/><a:lstStyle/>
          <a:p><a:r><a:t>Market Share</a:t></a:r></a:p>
        </c:rich>
      </c:tx>
    </c:title>
    <c:plotArea>
      <c:pieChart>
        <c:ser>
          <c:idx val="0"/>
          <c:tx>
            <c:strRef>
              <c:strCache>
                <c:ptCount val="1"/>
                <c:pt idx="0"><c:v>Share</c:v></c:pt>
              </c:strCache>
            </c:strRef>
          </c:tx>
          <c:cat>
            <c:strRef>
              <c:strCache>
                <c:ptCount val="3"/>
                <c:pt idx="0"><c:v>Product A</c:v></c:pt>
                <c:pt idx="1"><c:v>Product B</c:v></c:pt>
                <c:pt idx="2"><c:v>Product C</c:v></c:pt>
              </c:strCache>
            </c:strRef>
          </c:cat>
          <c:val>
            <c:numRef>
              <c:numCache>
                <c:ptCount val="3"/>
                <c:pt idx="0"><c:v>45</c:v></c:pt>
                <c:pt idx="1"><c:v>30</c:v></c:pt>
                <c:pt idx="2"><c:v>25</c:v></c:pt>
              </c:numCache>
            </c:numRef>
          </c:val>
        </c:ser>
      </c:pieChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;

const lineChartXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:plotArea>
      <c:lineChart>
        <c:ser>
          <c:idx val="0"/>
          <c:tx>
            <c:strRef>
              <c:strCache>
                <c:ptCount val="1"/>
                <c:pt idx="0"><c:v>Temperature</c:v></c:pt>
              </c:strCache>
            </c:strRef>
          </c:tx>
          <c:cat>
            <c:strRef>
              <c:strCache>
                <c:ptCount val="5"/>
                <c:pt idx="0"><c:v>Jan</c:v></c:pt>
                <c:pt idx="1"><c:v>Feb</c:v></c:pt>
                <c:pt idx="2"><c:v>Mar</c:v></c:pt>
                <c:pt idx="3"><c:v>Apr</c:v></c:pt>
                <c:pt idx="4"><c:v>May</c:v></c:pt>
              </c:strCache>
            </c:strRef>
          </c:cat>
          <c:val>
            <c:numRef>
              <c:numCache>
                <c:ptCount val="5"/>
                <c:pt idx="0"><c:v>5</c:v></c:pt>
                <c:pt idx="1"><c:v>8</c:v></c:pt>
                <c:pt idx="2"><c:v>12</c:v></c:pt>
                <c:pt idx="3"><c:v>18</c:v></c:pt>
                <c:pt idx="4"><c:v>22</c:v></c:pt>
              </c:numCache>
            </c:numRef>
          </c:val>
        </c:ser>
      </c:lineChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;

const scatterChartXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:title>
      <c:tx>
        <c:rich><a:bodyPr/><a:lstStyle/>
          <a:p><a:r><a:t>XY Plot</a:t></a:r></a:p>
        </c:rich>
      </c:tx>
    </c:title>
    <c:plotArea>
      <c:scatterChart>
        <c:ser>
          <c:idx val="0"/>
          <c:tx>
            <c:strRef>
              <c:strCache>
                <c:ptCount val="1"/>
                <c:pt idx="0"><c:v>Data</c:v></c:pt>
              </c:strCache>
            </c:strRef>
          </c:tx>
          <c:xVal>
            <c:numRef>
              <c:numCache>
                <c:ptCount val="4"/>
                <c:pt idx="0"><c:v>1</c:v></c:pt>
                <c:pt idx="1"><c:v>2</c:v></c:pt>
                <c:pt idx="2"><c:v>3</c:v></c:pt>
                <c:pt idx="3"><c:v>4</c:v></c:pt>
              </c:numCache>
            </c:numRef>
          </c:xVal>
          <c:yVal>
            <c:numRef>
              <c:numCache>
                <c:ptCount val="4"/>
                <c:pt idx="0"><c:v>10</c:v></c:pt>
                <c:pt idx="1"><c:v>20</c:v></c:pt>
                <c:pt idx="2"><c:v>15</c:v></c:pt>
                <c:pt idx="3"><c:v>25</c:v></c:pt>
              </c:numCache>
            </c:numRef>
          </c:yVal>
        </c:ser>
      </c:scatterChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;

const doughnutChartXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:plotArea>
      <c:doughnutChart>
        <c:ser>
          <c:idx val="0"/>
          <c:cat>
            <c:strLit>
              <c:ptCount val="2"/>
              <c:pt idx="0"><c:v>Used</c:v></c:pt>
              <c:pt idx="1"><c:v>Free</c:v></c:pt>
            </c:strLit>
          </c:cat>
          <c:val>
            <c:numLit>
              <c:ptCount val="2"/>
              <c:pt idx="0"><c:v>70</c:v></c:pt>
              <c:pt idx="1"><c:v>30</c:v></c:pt>
            </c:numLit>
          </c:val>
        </c:ser>
      </c:doughnutChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;

const horizontalBarChartXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="bar"/>
        <c:ser>
          <c:idx val="0"/>
          <c:cat>
            <c:strRef>
              <c:strCache>
                <c:ptCount val="3"/>
                <c:pt idx="0"><c:v>A</c:v></c:pt>
                <c:pt idx="1"><c:v>B</c:v></c:pt>
                <c:pt idx="2"><c:v>C</c:v></c:pt>
              </c:strCache>
            </c:strRef>
          </c:cat>
          <c:val>
            <c:numRef>
              <c:numCache>
                <c:ptCount val="3"/>
                <c:pt idx="0"><c:v>10</c:v></c:pt>
                <c:pt idx="1"><c:v>20</c:v></c:pt>
                <c:pt idx="2"><c:v>30</c:v></c:pt>
              </c:numCache>
            </c:numRef>
          </c:val>
        </c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;

const areaChartXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:plotArea>
      <c:areaChart>
        <c:ser>
          <c:idx val="0"/>
          <c:cat>
            <c:numRef>
              <c:numCache>
                <c:ptCount val="3"/>
                <c:pt idx="0"><c:v>2020</c:v></c:pt>
                <c:pt idx="1"><c:v>2021</c:v></c:pt>
                <c:pt idx="2"><c:v>2022</c:v></c:pt>
              </c:numCache>
            </c:numRef>
          </c:cat>
          <c:val>
            <c:numRef>
              <c:numCache>
                <c:ptCount val="3"/>
                <c:pt idx="0"><c:v>50</c:v></c:pt>
                <c:pt idx="1"><c:v>75</c:v></c:pt>
                <c:pt idx="2"><c:v>100</c:v></c:pt>
              </c:numCache>
            </c:numRef>
          </c:val>
        </c:ser>
      </c:areaChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;

const drawingWithChartXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
          xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <xdr:twoCellAnchor>
    <xdr:from>
      <xdr:col>3</xdr:col>
      <xdr:colOff>0</xdr:colOff>
      <xdr:row>1</xdr:row>
      <xdr:rowOff>0</xdr:rowOff>
    </xdr:from>
    <xdr:to>
      <xdr:col>12</xdr:col>
      <xdr:colOff>0</xdr:colOff>
      <xdr:row>18</xdr:row>
      <xdr:rowOff>0</xdr:rowOff>
    </xdr:to>
    <xdr:graphicFrame macro="">
      <xdr:nvGraphicFramePr>
        <xdr:cNvPr id="2" name="Chart 1"/>
        <xdr:cNvGraphicFramePr/>
      </xdr:nvGraphicFramePr>
      <xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart r:id="rId1"/>
        </a:graphicData>
      </a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
  <xdr:twoCellAnchor>
    <xdr:from>
      <xdr:col>0</xdr:col>
      <xdr:colOff>0</xdr:colOff>
      <xdr:row>0</xdr:row>
      <xdr:rowOff>0</xdr:rowOff>
    </xdr:from>
    <xdr:to>
      <xdr:col>2</xdr:col>
      <xdr:colOff>0</xdr:colOff>
      <xdr:row>5</xdr:row>
      <xdr:rowOff>0</xdr:rowOff>
    </xdr:to>
    <xdr:pic>
      <xdr:nvPicPr>
        <xdr:cNvPr id="1" name="Picture 1"/>
        <xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr>
      </xdr:nvPicPr>
      <xdr:blipFill>
        <a:blip r:embed="rId2"/>
        <a:stretch><a:fillRect/></a:stretch>
      </xdr:blipFill>
      <xdr:spPr>
        <a:xfrm><a:off x="0" y="0"/><a:ext cx="1" cy="1"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      </xdr:spPr>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>`;

const emptyChartXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;

const seriesWithColorXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:ser>
          <c:idx val="0"/>
          <c:tx><c:v>Colored</c:v></c:tx>
          <c:spPr>
            <a:solidFill>
              <a:srgbClr val="FF0000"/>
            </a:solidFill>
          </c:spPr>
          <c:cat>
            <c:strLit>
              <c:ptCount val="2"/>
              <c:pt idx="0"><c:v>X</c:v></c:pt>
              <c:pt idx="1"><c:v>Y</c:v></c:pt>
            </c:strLit>
          </c:cat>
          <c:val>
            <c:numLit>
              <c:ptCount val="2"/>
              <c:pt idx="0"><c:v>5</c:v></c:pt>
              <c:pt idx="1"><c:v>10</c:v></c:pt>
            </c:numLit>
          </c:val>
        </c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`;

// ---- Tests ----

describe('ChartParser', () => {
  describe('parseChartDocument', () => {
    it('parses a bar chart with title, two series, categories, and values', () => {
      const doc = XmlParser.parseXml(barChartXml);
      const chart = ChartParser.parseChartDocument(doc);

      expect(chart.type).toBe('col');
      expect(chart.title).toBe('Sales by Region');
      expect(chart.series).toHaveLength(2);

      expect(chart.series[0].name).toBe('Revenue');
      expect(chart.series[0].points).toHaveLength(4);
      expect(chart.series[0].points[0]).toEqual({ category: 'North', value: 100 });
      expect(chart.series[0].points[1]).toEqual({ category: 'South', value: 200 });
      expect(chart.series[0].points[2]).toEqual({ category: 'East', value: 150 });
      expect(chart.series[0].points[3]).toEqual({ category: 'West', value: 250 });

      expect(chart.series[1].name).toBe('Costs');
      expect(chart.series[1].points[0]).toEqual({ category: 'North', value: 80 });
      expect(chart.series[1].points[3]).toEqual({ category: 'West', value: 180 });
    });

    it('parses axis info from a bar chart', () => {
      const doc = XmlParser.parseXml(barChartXml);
      const chart = ChartParser.parseChartDocument(doc);

      expect(chart.categoryAxis).toBeDefined();
      expect(chart.categoryAxis!.title).toBe('Region');
      expect(chart.valueAxis).toBeDefined();
      expect(chart.valueAxis!.title).toBe('Amount');
    });

    it('parses a pie chart with cached string/numeric values', () => {
      const doc = XmlParser.parseXml(pieChartXml);
      const chart = ChartParser.parseChartDocument(doc);

      expect(chart.type).toBe('pie');
      expect(chart.title).toBe('Market Share');
      expect(chart.series).toHaveLength(1);
      expect(chart.series[0].name).toBe('Share');
      expect(chart.series[0].points).toHaveLength(3);
      expect(chart.series[0].points[0]).toEqual({ category: 'Product A', value: 45 });
      expect(chart.series[0].points[1]).toEqual({ category: 'Product B', value: 30 });
      expect(chart.series[0].points[2]).toEqual({ category: 'Product C', value: 25 });
    });

    it('parses a line chart with 5 data points', () => {
      const doc = XmlParser.parseXml(lineChartXml);
      const chart = ChartParser.parseChartDocument(doc);

      expect(chart.type).toBe('line');
      expect(chart.title).toBe('');
      expect(chart.series).toHaveLength(1);
      expect(chart.series[0].name).toBe('Temperature');
      expect(chart.series[0].points).toHaveLength(5);
      expect(chart.series[0].points[0]).toEqual({ category: 'Jan', value: 5 });
      expect(chart.series[0].points[4]).toEqual({ category: 'May', value: 22 });
    });

    it('parses a scatter chart with xVal / yVal', () => {
      const doc = XmlParser.parseXml(scatterChartXml);
      const chart = ChartParser.parseChartDocument(doc);

      expect(chart.type).toBe('scatter');
      expect(chart.title).toBe('XY Plot');
      expect(chart.series).toHaveLength(1);
      expect(chart.series[0].name).toBe('Data');
      expect(chart.series[0].points).toHaveLength(4);
      expect(chart.series[0].points[0]).toEqual({ category: '1', value: 10 });
      expect(chart.series[0].points[3]).toEqual({ category: '4', value: 25 });
    });

    it('parses a doughnut chart with literal values', () => {
      const doc = XmlParser.parseXml(doughnutChartXml);
      const chart = ChartParser.parseChartDocument(doc);

      expect(chart.type).toBe('doughnut');
      expect(chart.series).toHaveLength(1);
      expect(chart.series[0].points).toHaveLength(2);
      expect(chart.series[0].points[0]).toEqual({ category: 'Used', value: 70 });
      expect(chart.series[0].points[1]).toEqual({ category: 'Free', value: 30 });
    });

    it('parses a horizontal bar chart and sets type to bar', () => {
      const doc = XmlParser.parseXml(horizontalBarChartXml);
      const chart = ChartParser.parseChartDocument(doc);

      expect(chart.type).toBe('bar');
      expect(chart.series).toHaveLength(1);
      expect(chart.series[0].points).toHaveLength(3);
      expect(chart.series[0].points[2]).toEqual({ category: 'C', value: 30 });
    });

    it('parses an area chart with numeric categories', () => {
      const doc = XmlParser.parseXml(areaChartXml);
      const chart = ChartParser.parseChartDocument(doc);

      expect(chart.type).toBe('area');
      expect(chart.series).toHaveLength(1);
      expect(chart.series[0].points).toHaveLength(3);
      expect(chart.series[0].points[0]).toEqual({ category: '2020', value: 50 });
      expect(chart.series[0].points[2]).toEqual({ category: '2022', value: 100 });
    });

    it('handles an empty chart with no series', () => {
      const doc = XmlParser.parseXml(emptyChartXml);
      const chart = ChartParser.parseChartDocument(doc);

      expect(chart.type).toBe('col');
      expect(chart.series).toHaveLength(0);
      expect(chart.title).toBe('');
    });

    it('parses series color from spPr solidFill', () => {
      const doc = XmlParser.parseXml(seriesWithColorXml);
      const chart = ChartParser.parseChartDocument(doc);

      expect(chart.series).toHaveLength(1);
      expect(chart.series[0].color).toBe('#FF0000');
      expect(chart.series[0].points[0]).toEqual({ category: 'X', value: 5 });
    });

    it('sets usesCache to true when series have points', () => {
      const doc = XmlParser.parseXml(barChartXml);
      const chart = ChartParser.parseChartDocument(doc);
      expect(chart.usesCache).toBe(true);
    });

    it('sets usesCache to false when no series are present', () => {
      const doc = XmlParser.parseXml(emptyChartXml);
      const chart = ChartParser.parseChartDocument(doc);
      expect(chart.usesCache).toBe(false);
    });
  });

  describe('parseChartAnchors', () => {
    it('extracts chart anchors from a drawing with graphicFrame elements', () => {
      const doc = XmlParser.parseXml(drawingWithChartXml);
      const anchors = ChartParser.parseChartAnchors(doc);

      expect(anchors).toHaveLength(1);
      expect(anchors[0].relId).toBe('rId1');
      expect(anchors[0].from).toEqual({ col: 4, row: 2 });
      expect(anchors[0].to).toEqual({ col: 13, row: 19 });
    });

    it('does not include image-only anchors as chart anchors', () => {
      const doc = XmlParser.parseXml(drawingWithChartXml);
      const anchors = ChartParser.parseChartAnchors(doc);

      // Should only have 1 chart, not the image
      expect(anchors).toHaveLength(1);
      expect(anchors[0].relId).toBe('rId1');
    });

    it('returns empty array for drawing with no charts', () => {
      const noChartsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>0</xdr:col><xdr:row>0</xdr:row></xdr:from>
    <xdr:to><xdr:col>2</xdr:col><xdr:row>5</xdr:row></xdr:to>
    <xdr:pic>
      <xdr:nvPicPr>
        <xdr:cNvPr id="1" name="Picture 1"/>
        <xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr>
      </xdr:nvPicPr>
      <xdr:blipFill>
        <a:blip r:embed="rId1"/>
        <a:stretch><a:fillRect/></a:stretch>
      </xdr:blipFill>
      <xdr:spPr>
        <a:xfrm><a:off x="0" y="0"/><a:ext cx="1" cy="1"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      </xdr:spPr>
    </xdr:pic>
  </xdr:twoCellAnchor>
</xdr:wsDr>`;
      const doc = XmlParser.parseXml(noChartsXml);
      const anchors = ChartParser.parseChartAnchors(doc);
      expect(anchors).toHaveLength(0);
    });
  });
});
