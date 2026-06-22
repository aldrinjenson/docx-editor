/**
 * Regression — anchored wps:wsp text boxes wrapped in
 * <mc:AlternateContent><mc:Choice Requires="wps">...</mc:Choice></mc:AlternateContent>
 * must reach the text-box pipeline (not just direct <w:r> children).
 */

import { describe, expect, test } from 'bun:test';
import { parseDocumentBody } from '../documentParser';
import type { MediaFile, Relationship, RelationshipMap } from '../../types/document';

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"';

const TINY_PNG_DATA_URL =
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII=';

function buildDocumentWithAlternateContentTextBox(roleName: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:document ${NS}>
      <w:body>
        <w:p>
          <w:r>
            <mc:AlternateContent>
              <mc:Choice Requires="wps">
                <w:drawing>
                  <wp:anchor distT="45720" distB="45720" distL="114300" distR="114300"
                    simplePos="0" relativeHeight="251695104" behindDoc="0" locked="0"
                    layoutInCell="1" allowOverlap="1">
                    <wp:simplePos x="0" y="0"/>
                    <wp:positionH relativeFrom="margin">
                      <wp:align>right</wp:align>
                    </wp:positionH>
                    <wp:positionV relativeFrom="paragraph">
                      <wp:posOffset>0</wp:posOffset>
                    </wp:positionV>
                    <wp:extent cx="1390650" cy="278130"/>
                    <wp:effectExtent l="0" t="0" r="0" b="0"/>
                    <wp:wrapSquare wrapText="bothSides"/>
                    <wp:docPr id="1" name="Text Box 1"/>
                    <wp:cNvGraphicFramePr/>
                    <a:graphic>
                      <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                        <wps:wsp>
                          <wps:cNvSpPr txBox="1"/>
                          <wps:spPr>
                            <a:xfrm>
                              <a:off x="0" y="0"/>
                              <a:ext cx="1390650" cy="278130"/>
                            </a:xfrm>
                            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                          </wps:spPr>
                          <wps:txbx>
                            <w:txbxContent>
                              <w:p>
                                <w:r>
                                  <w:t>${roleName}</w:t>
                                </w:r>
                              </w:p>
                            </w:txbxContent>
                          </wps:txbx>
                          <wps:bodyPr/>
                        </wps:wsp>
                      </a:graphicData>
                    </a:graphic>
                  </wp:anchor>
                </w:drawing>
              </mc:Choice>
              <mc:Fallback>
                <w:pict/>
              </mc:Fallback>
            </mc:AlternateContent>
          </w:r>
        </w:p>
      </w:body>
    </w:document>`;
}

/**
 * Build a paragraph with N `<mc:AlternateContent>`-wrapped shapes, each in
 * its own `<w:r>` and each carrying only the AC wrapper (no text). Exercises
 * the runIndex clamp: when the parser collapses empty AC-only runs, every
 * shape past the first would otherwise hit `runIndex >= paragraph.content.length`
 * and be dropped.
 */
function buildDocumentWithMultipleAlternateContentShapes(roleNames: string[]): string {
  const runs = roleNames
    .map(
      (name) => `
          <w:r>
            <mc:AlternateContent>
              <mc:Choice Requires="wps">
                <w:drawing>
                  <wp:anchor distT="45720" distB="45720" distL="114300" distR="114300"
                    simplePos="0" relativeHeight="251695104" behindDoc="0" locked="0"
                    layoutInCell="1" allowOverlap="1">
                    <wp:simplePos x="0" y="0"/>
                    <wp:positionH relativeFrom="margin"><wp:align>left</wp:align></wp:positionH>
                    <wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>
                    <wp:extent cx="1390650" cy="278130"/>
                    <wp:effectExtent l="0" t="0" r="0" b="0"/>
                    <wp:wrapSquare wrapText="bothSides"/>
                    <wp:docPr id="1" name="Card"/>
                    <wp:cNvGraphicFramePr/>
                    <a:graphic>
                      <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                        <wps:wsp>
                          <wps:cNvSpPr txBox="1"/>
                          <wps:spPr>
                            <a:xfrm><a:off x="0" y="0"/><a:ext cx="1390650" cy="278130"/></a:xfrm>
                            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                          </wps:spPr>
                          <wps:txbx>
                            <w:txbxContent>
                              <w:p><w:r><w:t>${name}</w:t></w:r></w:p>
                            </w:txbxContent>
                          </wps:txbx>
                          <wps:bodyPr/>
                        </wps:wsp>
                      </a:graphicData>
                    </a:graphic>
                  </wp:anchor>
                </w:drawing>
              </mc:Choice>
              <mc:Fallback><w:pict/></mc:Fallback>
            </mc:AlternateContent>
          </w:r>`
    )
    .join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:document ${NS}>
      <w:body>
        <w:p>${runs}<w:r><w:t>Body text anchoring the cards.</w:t></w:r></w:p>
      </w:body>
    </w:document>`;
}

function buildDocumentWithTextBoxPicture(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:document ${NS}>
      <w:body>
        <w:p>
          <w:r>
            <mc:AlternateContent>
              <mc:Choice Requires="wps">
                <w:drawing>
                  <wp:anchor distT="0" distB="0" distL="0" distR="0"
                    simplePos="0" relativeHeight="251664384" behindDoc="0"
                    locked="0" layoutInCell="1" allowOverlap="1">
                    <wp:simplePos x="0" y="0"/>
                    <wp:positionH relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionH>
                    <wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>
                    <wp:extent cx="914400" cy="457200"/>
                    <wp:wrapNone/>
                    <wp:docPr id="1" name="Text Box 1"/>
                    <a:graphic>
                      <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                        <wps:wsp>
                          <wps:cNvSpPr txBox="1"/>
                          <wps:spPr>
                            <a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm>
                            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                          </wps:spPr>
                          <wps:txbx>
                            <w:txbxContent>
                              <w:p>
                                <w:r>
                                  <w:drawing>
                                    <wp:inline distT="0" distB="0" distL="0" distR="0">
                                      <wp:extent cx="914400" cy="457200"/>
                                      <wp:docPr id="2" name="Picture 1" descr="Nested picture"/>
                                      <a:graphic>
                                        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                                          <pic:pic>
                                            <pic:nvPicPr>
                                              <pic:cNvPr id="2" name="picture.png"/>
                                              <pic:cNvPicPr/>
                                            </pic:nvPicPr>
                                            <pic:blipFill>
                                              <a:blip r:embed="rIdPicture"/>
                                              <a:stretch><a:fillRect/></a:stretch>
                                            </pic:blipFill>
                                            <pic:spPr>
                                              <a:xfrm><a:ext cx="914400" cy="457200"/></a:xfrm>
                                              <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                                            </pic:spPr>
                                          </pic:pic>
                                        </a:graphicData>
                                      </a:graphic>
                                    </wp:inline>
                                  </w:drawing>
                                </w:r>
                              </w:p>
                            </w:txbxContent>
                          </wps:txbx>
                          <wps:bodyPr/>
                        </wps:wsp>
                      </a:graphicData>
                    </a:graphic>
                  </wp:anchor>
                </w:drawing>
              </mc:Choice>
              <mc:Fallback><w:pict/></mc:Fallback>
            </mc:AlternateContent>
          </w:r>
        </w:p>
      </w:body>
    </w:document>`;
}

function buildDocumentWithAlternateContentVisualShape(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:document ${NS}>
      <w:body>
        <w:p>
          <w:r>
            <mc:AlternateContent>
              <mc:Choice Requires="wps">
                <w:drawing>
                  <wp:anchor distT="0" distB="0" distL="0" distR="0"
                    simplePos="0" relativeHeight="251664384" behindDoc="1"
                    locked="0" layoutInCell="1" allowOverlap="1">
                    <wp:simplePos x="0" y="0"/>
                    <wp:positionH relativeFrom="page"><wp:posOffset>914400</wp:posOffset></wp:positionH>
                    <wp:positionV relativeFrom="paragraph"><wp:posOffset>457200</wp:posOffset></wp:positionV>
                    <wp:extent cx="1828800" cy="228600"/>
                    <wp:wrapNone/>
                    <wp:docPr id="8" name="Navigation Background"/>
                    <a:graphic>
                      <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                        <wps:wsp>
                          <wps:cNvSpPr/>
                          <wps:spPr>
                            <a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="228600"/></a:xfrm>
                            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                            <a:solidFill><a:srgbClr val="808080"/></a:solidFill>
                            <a:ln><a:noFill/></a:ln>
                          </wps:spPr>
                          <wps:bodyPr/>
                        </wps:wsp>
                      </a:graphicData>
                    </a:graphic>
                  </wp:anchor>
                </w:drawing>
              </mc:Choice>
              <mc:Fallback><w:pict/></mc:Fallback>
            </mc:AlternateContent>
          </w:r>
        </w:p>
      </w:body>
    </w:document>`;
}

const pictureRels: RelationshipMap = new Map<string, Relationship>([
  [
    'rIdPicture',
    { id: 'rIdPicture', type: 'image', target: 'media/picture.png', targetMode: 'Internal' },
  ],
]);

const pictureMedia: Map<string, MediaFile> = new Map([
  [
    'word/media/picture.png',
    {
      path: 'word/media/picture.png',
      filename: 'picture.png',
      mimeType: 'image/png',
      data: new ArrayBuffer(0),
      dataUrl: TINY_PNG_DATA_URL,
    },
  ],
]);

describe('enrichParagraphTextBoxes — mc:AlternateContent traversal', () => {
  test('extracts a wps:wsp text box wrapped in mc:Choice', () => {
    const body = parseDocumentBody(buildDocumentWithAlternateContentTextBox('Operations Manager'));
    expect(body.content).toHaveLength(1);
    const paragraph = body.content[0];
    expect(paragraph.type).toBe('paragraph');
    if (paragraph.type !== 'paragraph') return;

    // Walk the parsed runs and find any ShapeContent — should have exactly one
    const shapes = paragraph.content.flatMap((c) =>
      c.type === 'run' ? c.content.filter((rc) => rc.type === 'shape') : []
    );
    expect(shapes).toHaveLength(1);
    const shape = shapes[0];
    if (shape.type !== 'shape') return;

    expect(shape.shape.textBody?.content).toBeDefined();
    expect(shape.shape.textBody!.content).toHaveLength(1);
    const innerPara = shape.shape.textBody!.content[0];
    if (innerPara.type !== 'paragraph') throw new Error('expected paragraph');
    const innerRun = innerPara.content[0];
    if (innerRun.type !== 'run') throw new Error('expected run');
    const innerText = innerRun.content[0];
    if (innerText.type !== 'text') throw new Error('expected text');
    expect(innerText.text).toBe('Operations Manager');
  });

  test('runIndex clamp: every shape in a multi-shape paragraph survives', () => {
    // Org-chart pattern: three AC-only runs followed by a text run. Without the
    // clamp, the second and third shapes get dropped because their XML runIndex
    // (1, 2) outruns the collapsed paragraph.content.
    const roles = ['CEO', 'CTO', 'CFO'];
    const body = parseDocumentBody(buildDocumentWithMultipleAlternateContentShapes(roles));
    const paragraph = body.content[0];
    if (paragraph.type !== 'paragraph') throw new Error('expected paragraph');

    const shapes = paragraph.content.flatMap((c) =>
      c.type === 'run' ? c.content.filter((rc) => rc.type === 'shape') : []
    );

    expect(shapes).toHaveLength(roles.length);

    const innerTexts = shapes.map((s) => {
      if (s.type !== 'shape') throw new Error('expected shape');
      const innerPara = s.shape.textBody!.content[0];
      if (innerPara.type !== 'paragraph') throw new Error('expected paragraph');
      const innerRun = innerPara.content[0];
      if (innerRun.type !== 'run') throw new Error('expected run');
      const innerText = innerRun.content[0];
      if (innerText.type !== 'text') throw new Error('expected text');
      return innerText.text;
    });
    expect(innerTexts).toEqual(roles);
  });

  test('resolves picture media inside text-box paragraphs', () => {
    const body = parseDocumentBody(
      buildDocumentWithTextBoxPicture(),
      null,
      null,
      null,
      pictureRels,
      pictureMedia
    );
    const paragraph = body.content[0];
    if (paragraph.type !== 'paragraph') throw new Error('expected paragraph');

    const shapeContent = paragraph.content
      .flatMap((c) => (c.type === 'run' ? c.content : []))
      .find((rc) => rc.type === 'shape');
    if (!shapeContent || shapeContent.type !== 'shape') throw new Error('expected shape');

    const innerPara = shapeContent.shape.textBody?.content[0];
    if (!innerPara || innerPara.type !== 'paragraph') throw new Error('expected inner paragraph');
    const innerRun = innerPara.content[0];
    if (innerRun.type !== 'run') throw new Error('expected run');
    const innerDrawing = innerRun.content[0];
    if (innerDrawing.type !== 'drawing') throw new Error('expected drawing');

    expect(innerDrawing.image.rId).toBe('rIdPicture');
    expect(innerDrawing.image.src).toBe(TINY_PNG_DATA_URL);
  });

  test('preserves visual-only wps shapes wrapped in mc:Choice', () => {
    const body = parseDocumentBody(buildDocumentWithAlternateContentVisualShape());
    const paragraph = body.content[0];
    if (paragraph.type !== 'paragraph') throw new Error('expected paragraph');

    const shapes = paragraph.content.flatMap((c) =>
      c.type === 'run' ? c.content.filter((rc) => rc.type === 'shape') : []
    );

    expect(shapes).toHaveLength(1);
    const shapeContent = shapes[0];
    if (shapeContent.type !== 'shape') throw new Error('expected shape');

    expect(shapeContent.shape.name).toBe('Navigation Background');
    expect(shapeContent.shape.shapeType).toBe('rect');
    expect(shapeContent.shape.fill?.type).toBe('solid');
    expect(shapeContent.shape.fill?.color?.rgb).toBe('808080');
    expect(shapeContent.shape.textBody).toBeUndefined();
    expect(shapeContent.shape.position?.vertical.posOffset).toBe(457200);
    expect(shapeContent.shape.wrap?.type).toBe('behind');
  });
});
