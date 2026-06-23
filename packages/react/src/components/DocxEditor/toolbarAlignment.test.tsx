import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, beforeAll } from 'bun:test';

// Register happy-dom for this file and unregister after, matching the rest of
// the suite. A load-time register that never unregisters leaks the global
// registration across files and collides with other files' setup.
beforeAll(() => GlobalRegistrator.register());
afterAll(() => GlobalRegistrator.unregister());
import { afterEach, describe, expect, test } from 'bun:test';
import { cleanup, render } from '@testing-library/react';
import { DocxEditorToolbar } from './DocxEditorToolbar';

afterEach(() => {
  cleanup();
});

function renderToolbar(toolbarAlignment: 'start' | 'center' | 'end') {
  const noop = () => {};
  return render(
    <DocxEditorToolbar
      toolbarRefCallback={noop}
      agentPanelOpen={false}
      setAgentPanelOpen={noop}
      document={null}
      theme={null}
      pmState={null}
      selectionFormatting={{}}
      tableContext={null}
      imageContext={null}
      readOnly={false}
      editingMode="editing"
      setEditingMode={noop}
      setShowCommentsSidebar={noop}
      setExpandedSidebarItem={noop}
      showCommentsSidebar={false}
      agentPanel={undefined}
      renderLogo={undefined}
      documentName={undefined}
      onDocumentNameChange={undefined}
      documentNameEditable={true}
      renderTitleBarRight={undefined}
      toolbarExtra={null}
      fontFamilies={undefined}
      zoom={1}
      showZoomControl={false}
      toolbarAlignment={toolbarAlignment}
      onFormat={noop}
      onUndo={noop}
      onRedo={noop}
      onPrint={noop}
      showFileOpen={true}
      showHelpMenu={true}
      onOpen={noop}
      onSave={noop}
      onZoomChange={noop}
      onRefocusEditor={noop}
      onInsertTable={noop}
      onInsertImage={noop}
      onInsertPageBreak={noop}
      onInsertSectionBreakNextPage={noop}
      onInsertSectionBreakContinuous={noop}
      onInsertTOC={noop}
      onImageWrapType={noop}
      onImageTransform={noop}
      onOpenImageProperties={noop}
      onPageSetup={noop}
      onWatermark={noop}
      onTableAction={noop}
    />
  );
}

describe('Formatting toolbar alignment', () => {
  test('defaults to start — no alignment margins on the formatting bar', () => {
    const { getByTestId } = renderToolbar('start');
    const bar = getByTestId('formatting-bar');
    expect(bar.className).not.toContain('ms-auto');
    expect(bar.className).not.toContain('me-auto');
  });

  test('center auto-margins both ends so the contents sit in the middle', () => {
    const { getByTestId } = renderToolbar('center');
    const bar = getByTestId('formatting-bar');
    expect(bar.className).toContain('[&>*:first-child]:ms-auto');
    expect(bar.className).toContain('[&>*:last-child]:me-auto');
  });

  test('end auto-margins the leading edge so the contents sit on the right', () => {
    const { getByTestId } = renderToolbar('end');
    const bar = getByTestId('formatting-bar');
    expect(bar.className).toContain('[&>*:first-child]:ms-auto');
    expect(bar.className).not.toContain('me-auto');
  });
});
