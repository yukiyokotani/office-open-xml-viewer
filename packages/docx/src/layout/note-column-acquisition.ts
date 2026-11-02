import { projectBodyOccurrence } from './occurrence-projection.js';
import type { NoteLayout } from './types.js';
import type { BodyLayoutSession, NoteLayoutAcquisitionInput } from './body-layout-kernel.js';
import type { ContinuationNoticeProvider, FootnoteCursor, FootnotePartition,
  FootnotePartitionExtent, partitionFootnote } from './footnote-fragmentation.js';

/** Reuse a shaped note epoch only at its original width. Every different
 * column width starts the unpainted source suffix through canonical paragraph
 * acquisition; a line index must never stand in for a source cut. */
export function acquireNoteColumnPartition(
  session: Pick<BodyLayoutSession, 'layoutNotes'>,
  request: NoteLayoutAcquisitionInput,
  cursor: FootnoteCursor | null,
  capacityPt: number,
  partition: typeof partitionFootnote,
  notice?: ContinuationNoticeProvider,
  extent?: FootnotePartitionExtent,
  inspectAcquisition?: (note: NoteLayout) => void,
): FootnotePartition | null {
  if (!session.layoutNotes) throw new Error('Note continuation requires note acquisition');
  // A row/cell cursor owns a shaped epoch. Complete-block boundaries may
  // reflow the suffix; an in-row cut must keep its source geometry.
  let reflow = !cursor?.tableCursor && cursor?.sourceContinuation !== undefined
    && cursor.inlineExtentPt !== request.container.bounds.widthPt;
  let sourceStart = reflow ? cursor!.sourceContinuation : cursor?.acquisitionStart;
  const acquire = () => session.layoutNotes!({ ...request,
    ...(sourceStart ? { sourceStart } : {}),
  })[0];
  let acquired = acquire();
  if (!cursor?.tableCursor && cursor?.sourceContinuation && acquired?.story.bandDependent && request.bandTranslationPt && !reflow) {
    // A page-owned host must see the translation of the retained suffix,
    // excluding source lines discarded by a prior shaped-epoch line cursor.
    sourceStart = cursor.sourceContinuation;
    reflow = true;
    acquired = acquire();
  }
  if (acquired) inspectAcquisition?.(acquired);
  if (extent !== 'overflow') {
    const result = acquired ? partition(acquired, reflow ? null : cursor, capacityPt, notice, extent) : null;
    return result ? Object.freeze({ ...result, fragment: destinationOwnedNote(result.fragment) }) : null;
  }
  const result = acquired ? partition(acquired, reflow ? null : cursor, capacityPt, notice)
    ?? partition(acquired, reflow ? null : cursor, capacityPt, notice, 'overflow') : null;
  return result ? Object.freeze({ ...result, fragment: destinationOwnedNote(result.fragment) }) : null;
}

/** Cache reuse shares source geometry, never a prior destination's ownership. */
export function destinationOwnedNote(note: NoteLayout): NoteLayout {
  const domain = `${note.flowDomainId}:${note.source.story}:${note.source.storyInstance}`;
  if (note.story.blocks.every(block => block.flowDomainId === domain)) return note;
  return Object.freeze({ ...note, story: Object.freeze({ ...note.story,
    blocks: Object.freeze(note.story.blocks.map((block, index) => {
      if (block.kind !== 'paragraph' && block.kind !== 'table') throw new Error('Unsupported retained note block');
      return projectBodyOccurrence(block, { occurrenceId: `${note.id}:block:${index}`,
        destination: { coordinateSpace: 'logical-page-points', flowDomainId: domain,
          translation: { xPt: 0, yPt: 0 } } });
    })),
  }) });
}
