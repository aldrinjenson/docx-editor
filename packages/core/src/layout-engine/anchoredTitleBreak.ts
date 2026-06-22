import type { FlowBlock, ParagraphBlock, TextBoxBlock } from './types';
import { isFloatingTextBoxBlock } from './textBoxFlow';

export function isFollowingBlockAnchorTextBox(block: FlowBlock): block is TextBoxBlock {
  return (
    block.kind === 'textBox' &&
    block.anchorTarget === 'followingBlock' &&
    isFloatingTextBoxBlock(block)
  );
}

export function continuousBreakPrecedesAnchoredTitle(blocks: FlowBlock[], index: number): boolean {
  for (let i = index + 1; i < blocks.length; i++) {
    const block = blocks[i];
    if (isEmptyParagraphBlock(block)) continue;
    if (isEmptyTextBoxBlock(block)) continue;
    return isPageCenteredFollowingAnchorTextBox(block);
  }
  return false;
}

function isEmptyParagraphBlock(block: FlowBlock): block is ParagraphBlock {
  return block.kind === 'paragraph' && block.runs.length === 0;
}

function isEmptyTextBoxBlock(block: FlowBlock): block is TextBoxBlock {
  return (
    block.kind === 'textBox' &&
    block.content.every((paragraph) =>
      paragraph.runs.every((run) => run.kind !== 'text' || run.text.length === 0)
    )
  );
}

function isPageCenteredFollowingAnchorTextBox(block: FlowBlock): block is TextBoxBlock {
  if (!isFollowingBlockAnchorTextBox(block)) return false;
  const vertical = block.position?.vertical;
  return (
    (block.wrapType === 'inFront' || block.wrapType === 'behind') &&
    (vertical?.relativeTo === 'page' || vertical?.relativeTo === 'margin') &&
    vertical.align === 'center'
  );
}
