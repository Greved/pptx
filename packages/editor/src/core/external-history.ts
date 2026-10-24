import type { PresentationData } from '@office-kit/pptx';
import type { ChangeSource } from './change-source.ts';

/** A host-owned, synchronous history replacing the editor's PPTX snapshots. */
export interface ExternalHistoryAdapter {
  /** Replace the canonical document and clear its history. Must be atomic on failure. */
  reset(presentation: PresentationData): void;
  /**
   * Accept a complete committed projection, or throw without changing canonical
   * state or history. Live gesture frames are not reported until the gesture ends.
   */
  commit(
    presentation: PresentationData,
    context: { readonly label: string; readonly source: ChangeSource; readonly slideIndex: number },
  ): void;
  /**
   * Return a usable projection of canonical state, without retaining the mutable
   * input as canonical state. Cancel/rollback must not move the history cursor.
   * Undo/redo must leave canonical state/history unchanged if they throw.
   */
  restore(
    action: 'undo' | 'redo' | 'cancel' | 'rollback',
    presentation: PresentationData,
  ): PresentationData;
  readonly canUndo: boolean;
  readonly canRedo: boolean;
  readonly undoLabel: string | null;
  readonly redoLabel: string | null;
}
