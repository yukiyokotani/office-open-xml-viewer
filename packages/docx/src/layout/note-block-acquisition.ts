import type { NoteSourceContinuation, StoryLayoutAcquisitionInput } from './body-layout-kernel.js';
import type { NoteLayout, ParagraphLayout, StoryLayout, TableLayout } from './types.js';

type Block = ParagraphLayout | TableLayout;
type Acquire = (block: Block, destination: StoryLayoutAcquisitionInput, start?: NoteSourceContinuation, capacityPt?: number, maximumLines?: number) => Block;

// Acquisition capabilities stay outside the immutable paint graph. An epoch
// shares page-independent blocks; partitioning resolves only visited field
// blocks, before fitting them. No full-note copy is made per destination.
const acquisitions = new WeakMap<StoryLayout, Acquire>();
const destinations = new WeakMap<NoteLayout, StoryLayoutAcquisitionInput>();

export function registerNoteBlockAcquisition(story: StoryLayout, acquire: Acquire): void {
  acquisitions.set(story, acquire);
}

export function registerNoteBlockDestination(note: NoteLayout, destination: StoryLayoutAcquisitionInput): void {
  if (acquisitions.has(note.story)) destinations.set(note, destination);
}

export function noteBlockAcquirer(note: NoteLayout): ((block: Block, start?: NoteSourceContinuation, capacityPt?: number, maximumLines?: number) => Block) | undefined {
  const acquire = acquisitions.get(note.story);
  const destination = destinations.get(note);
  return acquire && destination ? (block, start, capacityPt, maximumLines) => acquire(block, destination, start, capacityPt, maximumLines) : undefined;
}
