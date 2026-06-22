import { describe, expect, test } from 'bun:test';
import { parseDocumentBody } from '../documentParser';
import type { MediaFile, Relationship, RelationshipMap } from '../../types/document';

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

const DIAGRAM_DRAWING_REL =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramDrawing';

function documentWithDiagramDrawing(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:document ${NS}>
      <w:body>
        <w:p>
          <w:r>
            <w:drawing>
              <wp:anchor behindDoc="0" layoutInCell="1" allowOverlap="1">
                <wp:simplePos x="0" y="0"/>
                <wp:positionH relativeFrom="margin"><wp:posOffset>1000</wp:posOffset></wp:positionH>
                <wp:positionV relativeFrom="paragraph"><wp:posOffset>2000</wp:posOffset></wp:positionV>
                <wp:extent cx="900000" cy="100000"/>
                <wp:wrapNone/>
                <wp:docPr id="9" name="Diagram"/>
                <a:graphic>
                  <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">
                    <dgm:relIds r:dm="rIdData" r:lo="rIdLayout" r:qs="rIdStyle" r:cs="rIdColors"/>
                  </a:graphicData>
                </a:graphic>
              </wp:anchor>
            </w:drawing>
          </w:r>
        </w:p>
      </w:body>
    </w:document>`;
}

const diagramXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <dsp:drawing
    xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram"
    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
    <dsp:spTree>
      <dsp:sp>
        <dsp:spPr>
          <a:xfrm><a:off x="3000" y="4000"/><a:ext cx="100000" cy="200000"/></a:xfrm>
          <a:prstGeom prst="chevron"><a:avLst/></a:prstGeom>
        </dsp:spPr>
        <dsp:txBody>
          <a:bodyPr lIns="1000" tIns="2000" rIns="3000" bIns="4000"/>
          <a:p>
            <a:r>
              <a:rPr sz="650">
                <a:solidFill><a:srgbClr val="32D700"/></a:solidFill>
                <a:latin typeface="Arial"/>
              </a:rPr>
              <a:t>Section Label</a:t>
            </a:r>
          </a:p>
        </dsp:txBody>
        <dsp:txXfrm><a:off x="13000" y="14000"/><a:ext cx="70000" cy="120000"/></dsp:txXfrm>
      </dsp:sp>
    </dsp:spTree>
  </dsp:drawing>`;

const rels: RelationshipMap = new Map<string, Relationship>([
  [
    'rIdDrawing',
    {
      id: 'rIdDrawing',
      type: DIAGRAM_DRAWING_REL as Relationship['type'],
      target: 'diagrams/drawing1.xml',
      targetMode: 'Internal',
    },
  ],
]);

const media: Map<string, MediaFile> = new Map([
  [
    'word/diagrams/drawing1.xml',
    {
      path: 'word/diagrams/drawing1.xml',
      filename: 'drawing1.xml',
      mimeType: 'application/xml',
      data: new ArrayBuffer(0),
      text: diagramXml,
    },
  ],
  [
    'diagrams/drawing1.xml',
    {
      path: 'word/diagrams/drawing1.xml',
      filename: 'drawing1.xml',
      mimeType: 'application/xml',
      data: new ArrayBuffer(0),
      text: diagramXml,
    },
  ],
]);

describe('diagram drawing fallback', () => {
  test('projects cached diagram drawing labels into anchored shape text boxes', () => {
    const body = parseDocumentBody(documentWithDiagramDrawing(), null, null, null, rels, media);
    const paragraph = body.content[0];
    if (paragraph.type !== 'paragraph') throw new Error('expected paragraph');

    const shapeContent = paragraph.content
      .flatMap((c) => (c.type === 'run' ? c.content : []))
      .find((rc) => rc.type === 'shape');
    if (!shapeContent || shapeContent.type !== 'shape') throw new Error('expected shape');

    expect(shapeContent.shape.shapeType).toBe('chevron');
    expect(shapeContent.shape.size).toEqual({ width: 70000, height: 120000 });
    expect(shapeContent.shape.position?.horizontal.posOffset).toBe(14000);
    expect(shapeContent.shape.position?.vertical.posOffset).toBe(16000);
    expect(shapeContent.shape.wrap?.type).toBe('inFront');
    expect(shapeContent.shape.textBody?.margins).toEqual({
      top: 2000,
      right: 3000,
      bottom: 4000,
      left: 1000,
    });

    const innerPara = shapeContent.shape.textBody?.content[0];
    if (!innerPara || innerPara.type !== 'paragraph') throw new Error('expected inner paragraph');
    const innerRun = innerPara.content[0];
    if (innerRun.type !== 'run') throw new Error('expected run');
    expect(innerRun.formatting?.fontSize).toBe(13);
    expect(innerRun.formatting?.fontFamily?.ascii).toBe('Arial');
    expect(innerRun.formatting?.color?.rgb).toBe('32D700');
    expect(innerRun.content[0]).toEqual({ type: 'text', text: 'Section Label' });
  });
});
