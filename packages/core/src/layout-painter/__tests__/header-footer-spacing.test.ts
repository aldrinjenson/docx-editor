import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, beforeAll, describe, expect, test } from 'bun:test';
import { PAGE_CLASS_NAMES, renderPage, type HeaderFooterContent } from '../renderPage';
import type {
  Page,
  ParagraphBlock,
  ParagraphMeasure,
  TextBoxBlock,
  TextBoxMeasure,
} from '../../layout-engine/types';

beforeAll(() => GlobalRegistrator.register());
afterAll(() => GlobalRegistrator.unregister());

function makePage(): Page {
  return {
    number: 1,
    fragments: [],
    margins: {
      top: 96,
      right: 96,
      bottom: 96,
      left: 96,
      header: 48,
      footer: 48,
    },
    size: { w: 816, h: 1056 },
  };
}

function makeSeparatorContent(): HeaderFooterContent {
  const block: ParagraphBlock = {
    kind: 'paragraph',
    id: 'separator',
    runs: [],
    attrs: {
      spacing: { before: 8 },
      spacingExplicit: { before: true },
      borders: {
        bottom: {
          style: 'solid',
          width: 1,
          color: '#7A1F2B',
          space: 1.3333333333333333,
        },
      },
      defaultFontSize: 11,
      defaultFontFamily: 'Calibri',
    },
  };

  const measure: ParagraphMeasure = {
    kind: 'paragraph',
    lines: [
      {
        fromRun: 0,
        fromChar: 0,
        toRun: 0,
        toChar: 0,
        width: 0,
        ascent: 11,
        descent: 4,
        lineHeight: 17.9,
      },
    ],
    totalHeight: 25.9,
  };

  return {
    blocks: [block],
    measures: [measure],
    height: 25.9,
    visualTop: 0,
    visualBottom: 25.9,
  };
}

describe('renderPage header/footer paragraph spacing', () => {
  test('stacks body content above header and footer bands', () => {
    const page = makePage();
    const headerContent = makeSeparatorContent();
    const footerContent = makeSeparatorContent();

    const el = renderPage(
      page,
      {
        pageNumber: 1,
        totalPages: 1,
        section: 'body',
      },
      {
        document,
        headerContent,
        footerContent,
      }
    );

    const contentEl = el.querySelector<HTMLElement>(`.${PAGE_CLASS_NAMES.content}`);
    const headerEl = el.querySelector<HTMLElement>(`.${PAGE_CLASS_NAMES.header}`);
    const footerEl = el.querySelector<HTMLElement>(`.${PAGE_CLASS_NAMES.footer}`);

    expect(contentEl?.style.zIndex).toBe('1');
    expect(headerEl?.style.zIndex).toBe('0');
    expect(footerEl?.style.zIndex).toBe('0');
  });

  test('header paragraph honors explicit spacing.before in fragment positioning', () => {
    const page = makePage();
    const headerContent = makeSeparatorContent();

    const el = renderPage(
      page,
      {
        pageNumber: 1,
        totalPages: 1,
        section: 'body',
      },
      {
        document,
        headerContent,
      }
    );

    const headerEl = el.querySelector(`.${PAGE_CLASS_NAMES.header}`);
    expect(headerEl).toBeTruthy();

    const paragraphEl = headerEl?.querySelector('.layout-paragraph') as HTMLElement | null;
    expect(paragraphEl).toBeTruthy();
    expect(paragraphEl?.style.top).toBe('8px');
  });

  test('footer paragraph-relative floats stay anchored to the footer band', () => {
    const page = makePage();
    const floatingBox: TextBoxBlock = {
      kind: 'textBox',
      id: 'footer-float',
      width: 160,
      height: 120,
      content: [],
      displayMode: 'float',
      position: {
        vertical: { relativeTo: 'paragraph', posOffset: 0 },
      },
    };
    const paragraph: ParagraphBlock = {
      kind: 'paragraph',
      id: 'footer-flow',
      runs: [],
    };
    const floatingMeasure: TextBoxMeasure = {
      kind: 'textBox',
      width: 160,
      height: 120,
      innerMeasures: [],
    };
    const paragraphMeasure: ParagraphMeasure = {
      kind: 'paragraph',
      lines: [],
      totalHeight: 24,
    };
    const footerContent: HeaderFooterContent = {
      blocks: [floatingBox, paragraph],
      measures: [floatingMeasure, paragraphMeasure],
      height: 144,
      flowHeight: 24,
      visualTop: 0,
      visualBottom: 144,
    };

    const el = renderPage(
      page,
      {
        pageNumber: 1,
        totalPages: 1,
        section: 'body',
      },
      {
        document,
        footerContent,
      }
    );

    const footerEl = el.querySelector<HTMLElement>(`.${PAGE_CLASS_NAMES.footer}`);
    const footerContentEl = footerEl?.firstElementChild as HTMLElement | null;
    const textBoxEl = footerEl?.querySelector<HTMLElement>('.layout-textbox');

    expect(footerEl?.style.top).toBe('984px');
    expect(footerContentEl?.style.top).toBe('0px');
    expect(textBoxEl?.style.top).toBe('0px');
  });
});
