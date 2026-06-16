import type { Table } from '../../types/document';
import { intAttr } from './xmlUtils';

const DEFAULT_TABLE_WIDTH_DXA = 9360;

function positiveIntegerOrUndefined(value: number | undefined): number | undefined {
  if (!Number.isFinite(value) || value == null || value <= 0) return undefined;
  return Math.max(1, Math.round(value));
}

function distributeWidth(totalWidth: number, columnCount: number): number[] {
  const baseWidth = Math.floor(totalWidth / columnCount);
  let remainder = totalWidth - baseWidth * columnCount;

  return Array.from({ length: columnCount }, () => {
    const width = baseWidth + (remainder > 0 ? 1 : 0);
    remainder = Math.max(0, remainder - 1);
    return Math.max(1, width);
  });
}

function getGridColumnCount(table: Table): number {
  if (table.columnWidths && table.columnWidths.length > 0) {
    return table.columnWidths.length;
  }

  if (table.rows.length === 0) return 0;

  return table.rows[0].cells.reduce((count, cell) => {
    return count + (cell.formatting?.gridSpan ?? 1);
  }, 0);
}

function inferTableGridWidths(table: Table): number[] {
  const explicitWidths = table.columnWidths
    ?.map((width) => positiveIntegerOrUndefined(width))
    .filter((width): width is number => width != null);

  if (explicitWidths && explicitWidths.length > 0) {
    return explicitWidths;
  }

  const columnCount = getGridColumnCount(table);
  if (columnCount <= 0) return [];

  const firstRowCellWidths = table.rows[0]?.cells
    .flatMap((cell) => {
      const width = positiveIntegerOrUndefined(cell.formatting?.width?.value);
      const gridSpan = Math.max(1, Math.round(cell.formatting?.gridSpan ?? 1));
      return width == null ? [] : distributeWidth(width, gridSpan);
    })
    .filter((width): width is number => width != null);

  if (firstRowCellWidths && firstRowCellWidths.length === columnCount) {
    return firstRowCellWidths;
  }

  const tableWidth =
    table.formatting?.width?.type === 'dxa'
      ? positiveIntegerOrUndefined(table.formatting.width.value)
      : undefined;
  return distributeWidth(tableWidth ?? DEFAULT_TABLE_WIDTH_DXA, columnCount);
}

export function serializeTableGridForTable(table: Table): string {
  const columnWidths = inferTableGridWidths(table);
  if (columnWidths.length === 0) return '';

  const cols = columnWidths.map((w) => `<w:gridCol w:w="${intAttr(w)}"/>`);
  return `<w:tblGrid>${cols.join('')}</w:tblGrid>`;
}
