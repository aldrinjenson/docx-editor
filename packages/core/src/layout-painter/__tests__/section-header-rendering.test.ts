import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, beforeAll, describe, expect, test } from 'bun:test';
import { renderPages, type HeaderFooterContent } from '../renderPage';
import type { Page, ParagraphBlock, ParagraphMeasure } from '../../layout-engine/types';

beforeAll(() => GlobalRegistrator.register());
afterAll(() => GlobalRegistrator.unregister());

function page(number: number, sectionIndex: number): Page {
  return {
    number,
    sectionIndex,
    sectionPageNumber: 1,
    fragments: [],
    margins: { top: 72, right: 72, bottom: 72, left: 72, header: 36, footer: 36 },
    size: { w: 612, h: 792 },
  };
}

function header(label: string): HeaderFooterContent {
  const block: ParagraphBlock = {
    kind: 'paragraph',
    id: `header-${label}`,
    runs: [{ kind: 'text', text: label }],
  };
  const measure: ParagraphMeasure = {
    kind: 'paragraph',
    lines: [
      {
        fromRun: 0,
        fromChar: 0,
        toRun: 0,
        toChar: label.length,
        width: 80,
        ascent: 10,
        descent: 3,
        lineHeight: 14,
      },
    ],
    totalHeight: 14,
  };
  return {
    blocks: [block],
    measures: [measure],
    height: 14,
    flowHeight: 14,
    visualTop: 0,
    visualBottom: 14,
  };
}

describe('section-aware header rendering', () => {
  test('renders the header for each page section', () => {
    const container = document.createElement('div');
    renderPages([page(1, 0), page(2, 1)], container, {
      document,
      sectionHeaderFooterContent: {
        0: { headerContent: header('Section One') },
        1: { headerContent: header('Section Two') },
      },
    });

    const pages = Array.from(container.querySelectorAll<HTMLElement>('.layout-page'));
    expect(pages).toHaveLength(2);
    expect(pages[0].querySelector('.layout-page-header')?.textContent).toContain('Section One');
    expect(pages[1].querySelector('.layout-page-header')?.textContent).toContain('Section Two');
  });
});
