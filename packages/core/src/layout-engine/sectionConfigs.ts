import type { ColumnLayout, FlowBlock, PageMargins, SectionBreakBlock } from './types';

/**
 * Page-flow geometry resolved from a single section's properties.
 * Exported so adapters can reuse the same shape when measuring blocks per
 * section width, keeping pagination and measurement consistent.
 */
export type SectionLayoutConfig = {
  pageSize: { w: number; h: number };
  margins: PageMargins;
  /** Optional. Sections without explicit columns inherit `{ count: 1 }`. */
  columns?: ColumnLayout;
};

/**
 * Walk `blocks` once and collect per-section geometry. `configs` has one
 * entry per section break plus a trailing `finalConfig`. `breakIndices` is
 * 1-to-1 with the inner break entries (same length as `configs.length - 1`).
 * Callers that need the break `type` can read it from
 * `(blocks[breakIndices[i]] as SectionBreakBlock).type`.
 *
 * @internal
 */
export function collectSectionConfigs(
  blocks: FlowBlock[],
  initialConfig: SectionLayoutConfig,
  finalConfig: SectionLayoutConfig
): {
  configs: SectionLayoutConfig[];
  breakIndices: number[];
} {
  const configs: SectionLayoutConfig[] = [];
  const breakIndices: number[] = [];
  let previousConfig = initialConfig;
  for (let i = 0; i < blocks.length; i++) {
    if (blocks[i].kind !== 'sectionBreak') continue;
    const sb = blocks[i] as SectionBreakBlock;
    const config: SectionLayoutConfig = {
      pageSize: sb.pageSize ?? previousConfig.pageSize,
      margins: sb.margins ?? previousConfig.margins,
      columns: sb.columns,
    };
    configs.push(config);
    breakIndices.push(i);
    previousConfig = config;
  }
  configs.push(finalConfig);
  return { configs, breakIndices };
}
