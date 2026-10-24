// The single source of truth for the editor: the in-memory presentation, the
// current selection, undo/redo history, and derived rendering.
//
// The @office-kit/pptx model is a mutable object graph that the library's
// authoring functions edit in place. Rather than deep-proxy that graph (which
// fights the library's reference-based mutation), we hold it in `$state.raw`
// and drive reactivity with an explicit `version` counter bumped on every
// mutation. Derived values (the rendered SVG, the slide list) read `version`,
// so any command re-renders the canvas without the store needing to understand
// the shape of every mutation.
//
// By default undo/redo snapshots the presentation by *serializing it to .pptx bytes*, not
// by structuredClone: the model's real state hangs off a symbol-keyed
// `OpcPackage` instance that structuredClone silently drops, producing a
// corrupt copy. save→load is the library's guaranteed round-trip, so it is the
// only sound snapshot. Edits stay instant (the mutation runs synchronously);
// savePresentation captures bytes synchronously and returns a promise. Store
// that promise immediately so rapid edits remain distinct. Because each discrete
// gesture (a drag, a click) commits exactly one transaction, this is one
// serialization per user action — not per pointer move.
// A host-supplied external history replaces that mechanism entirely, including
// the initial snapshot, and restores projections synchronously from its model.

import {
  createPresentation,
  setGridSpacing,
  setSnapToGrid,
  addTitleSlide,
  getSlides,
  getCustomShows,
  findShapeById,
  loadPresentation,
  savePresentation,
} from '@office-kit/pptx';
import type { RememberedFill } from './remembered-fill.ts';
import { renderSlidePaintCached } from './slide-render-cache.ts';
import type { SlideSvgResult } from '@office-kit/pptx-preview';
import type { CustomShow, PresentationData, SlideData, SlideShapeData } from '@office-kit/pptx';
import type { ChangeSource } from './change-source.ts';
import type { ExternalHistoryAdapter } from './external-history.ts';
import { RegroupHistory } from './regroup-history.ts';
import { selectedSlideIndices, type Selection } from './selection.ts';

interface Snapshot {
  readonly bytes: Promise<Uint8Array>;
  readonly selection: Selection;
  readonly label: string;
  readonly formerGroups: RegroupHistory['records'];
}

const HISTORY_MAX = 60;

export class EditorDocument {
  readonly rememberedFills = new Map<string, RememberedFill>();
  /** The live presentation. Mutated in place by library commands. */
  pres = $state.raw<PresentationData>(createInitial());
  /** Bumped on every mutation to invalidate derived rendering. */
  version = $state(0);
  committedVersion = $state(0);
  liveEditing = $state(false);
  selection = $state.raw<Selection>({ kind: 'none', slideIndex: 0 });
  fileName = $state<string>('Untitled.pptx');
  dirty = $state<boolean>(false);
  /** Slide Master view's selected master and layout (see `CommandDoc.layoutTarget`). */
  layoutTarget = $state.raw<{
    readonly master: string | null;
    readonly partName: string | null;
  } | null>(null);

  // History of committed states as serialized snapshots. `#cursor` is the index
  // of the current state within `#history`.
  #history = $state.raw<Snapshot[]>([]);
  #cursor = $state(-1);
  // A requested restore may still be loading while another undo is requested.
  #requestedCursor = $state(-1);
  #operation = 0;
  // Set while a failed edit's partial mutation is being replaced by the last
  // committed snapshot. Only a completed whole-document replacement clears it,
  // so no edit can snapshot (and thereby commit) the partial state meanwhile.
  #rollbackBlock: string | null = null;
  // What the pending rollback restores. Kept until a replacement completes, so
  // a newer Open/Undo/Redo that wins the race but then fails can resume it.
  #rollbackState: { selection: Selection; dirty: boolean } | null = null;
  #rollback: Promise<void> = Promise.resolve();
  readonly #externalHistory: ExternalHistoryAdapter | undefined;
  #liveDirty: boolean | undefined;
  readonly regroupHistory = new RegroupHistory();
  /**
   * Called after the committed document changes: an edit or gesture is
   * committed, a document setting changes, Undo/Redo, New or Open. Not for live
   * gesture frames, nor for a failed edit's rollback (which restores the state
   * last reported).
   */
  onChange: ((source: ChangeSource) => void) | undefined;

  constructor(externalHistory?: ExternalHistoryAdapter) {
    this.#externalHistory = externalHistory;
    // Seed the initial state so the first undo returns to the blank deck.
    if (externalHistory) externalHistory.reset(this.pres);
    else this.#snapshot('Initial');
  }

  // --- Derived views -----------------------------------------------------
  slides = $derived.by<ReadonlyArray<SlideData>>(() => {
    this.version;
    return getSlides(this.pres);
  });

  customShows = $derived.by<ReadonlyArray<CustomShow>>(() => {
    this.version;
    return getCustomShows(this.pres);
  });

  currentSlide = $derived.by<SlideData | null>(() => {
    const list = this.slides;
    const idx = this.selection.slideIndex;
    return list[idx] ?? list[0] ?? null;
  });

  currentPaint = $derived.by<SlideSvgResult | string>(() => {
    this.version;
    const slide = this.currentSlide;
    if (!slide) return '';
    try {
      return renderSlidePaintCached(this.pres, slide);
    } catch (err) {
      return `<!-- render error: ${(err as Error).message} -->`;
    }
  });

  currentSvg = $derived(
    typeof this.currentPaint === 'string' ? this.currentPaint : this.currentPaint.svg,
  );

  canUndo = $derived.by(() => {
    void this.committedVersion;
    return this.#externalHistory ? this.#externalHistory.canUndo : this.#requestedCursor > 0;
  });
  canRedo = $derived.by(() => {
    void this.committedVersion;
    return this.#externalHistory
      ? this.#externalHistory.canRedo
      : this.#requestedCursor >= 0 && this.#requestedCursor < this.#history.length - 1;
  });
  /** The name of the edit Undo reverts (Edit ▸ Undo <name>), or null. */
  undoLabel = $derived.by(() => {
    void this.committedVersion;
    return this.#externalHistory
      ? this.#externalHistory.undoLabel
      : this.#requestedCursor > 0
        ? (this.#history[this.#requestedCursor]?.label ?? null)
        : null;
  });
  /** The name of the edit Redo restores, or null. */
  redoLabel = $derived.by(() => {
    void this.committedVersion;
    return this.#externalHistory
      ? this.#externalHistory.redoLabel
      : this.canRedo
        ? (this.#history[this.#requestedCursor + 1]?.label ?? null)
        : null;
  });

  // --- Resolvers ---------------------------------------------------------
  slideAt(index: number): SlideData | null {
    return this.slides[index] ?? null;
  }

  shapeById(slideIndex: number, id: number): SlideShapeData | null {
    const slide = this.slideAt(slideIndex);
    return slide ? findShapeById(slide, id) : null;
  }

  // --- Mutation core -----------------------------------------------------
  /**
   * Run `fn` as one undoable transaction. The mutation runs synchronously so
   * the UI updates immediately; a byte snapshot is captured asynchronously for
   * undo. Returns whatever `fn` returns (e.g. a newly created shape/slide).
   * `dirty: false` is for a step that leaves the deck as its host has saved it.
   */
  transact<T>(label: string, fn: () => T, source: ChangeSource = 'user', dirty = true): T {
    if (this.#externalHistory && this.liveEditing)
      throw new Error('Finish the current gesture before starting another transaction.');
    const result = this.#atomic(() => {
      this.#invalidateRestore();
      const current = this.#history[this.#cursor];
      if (current)
        this.#history = this.#history.map((snapshot, index) =>
          index === this.#cursor ? { ...snapshot, selection: this.selection } : snapshot,
        );
      const result = fn();
      this.version++;
      this.dirty = dirty;
      this.#commitHistory(label, source);
      return result;
    });
    // Outside #atomic: a throwing listener must not roll back a committed edit.
    this.onChange?.(source);
    return result;
  }

  /** Persist a document setting without adding an undo step (as the reference desktop app does for aspect locks). */
  setDocumentSetting(fn: () => void): void {
    if (this.#externalHistory && this.liveEditing)
      throw new Error('Finish the current gesture before changing document settings.');
    this.#atomic(() => {
      this.#invalidateRestore();
      fn();
      this.version++;
      if (this.#externalHistory) {
        this.#commitHistory('Document setting', 'user');
        this.dirty = true;
        return;
      }
      this.committedVersion = this.version;
      this.dirty = true;
      const bytes = savePresentation(this.pres);
      this.#history = this.#history.map((snapshot, index) =>
        index === this.#cursor ? { ...snapshot, bytes } : snapshot,
      );
    });
    this.onChange?.('user');
  }

  /**
   * Apply a live, high-frequency mutation (a drag/resize frame). Re-renders
   * immediately but does NOT snapshot — so a whole gesture stays one undo step.
   * Call `commit()` once when the gesture ends.
   */
  applyLive<T>(fn: () => T): T {
    return this.#atomic(() => {
      if (!this.liveEditing) this.#liveDirty = this.dirty;
      this.liveEditing = true;
      this.#invalidateRestore();
      const result = fn();
      this.version++;
      this.dirty = true;
      return result;
    });
  }

  /** Close a live gesture by taking a single undo snapshot. */
  commit(label: string): void {
    this.#atomic(() => {
      this.liveEditing = false;
      this.#commitHistory(label, 'user');
      this.#liveDirty = undefined;
    });
    this.onChange?.('user');
  }

  /**
   * Settles once a failed edit has been replaced by the last committed state.
   * Edits stay synchronous, so the replacement (an async load) runs after the
   * failing call has already rethrown.
   */
  get rollback(): Promise<void> {
    return this.#rollback;
  }

  /**
   * Library mutations edit the package in place and are not transactional, so
   * an exception can leave a half-applied change (for example a media part
   * without its picture). Discard it by reloading the committed snapshot,
   * which also ends any live gesture.
   */
  #atomic<T>(fn: () => T): T {
    if (this.#rollbackBlock !== null) throw new Error(this.#rollbackBlock);
    const state = { selection: this.selection, dirty: this.dirty };
    try {
      return fn();
    } catch (error) {
      if (this.#externalHistory) {
        const dirty = this.#liveDirty ?? state.dirty;
        try {
          this.#restoreExternal('rollback');
        } catch (restoreError) {
          this.liveEditing = false;
          this.#rollbackBlock =
            'The previous edit failed and could not be undone. Reopen the presentation.';
          throw new AggregateError([error, restoreError], this.#rollbackBlock);
        }
        this.dirty = dirty;
        this.#liveDirty = undefined;
        throw error;
      }
      this.#startRollback(state);
      throw error;
    }
  }

  #startRollback(state: { selection: Selection; dirty: boolean }): void {
    this.#rollbackState = state;
    this.#rollbackBlock = 'The previous edit failed and is being undone. Try again.';
    this.#rollback = this.#restore(this.#cursor, state.selection, true).then(
      (applied) => {
        if (applied) this.dirty = state.dirty;
      },
      (cause: unknown) => {
        // A newer New/Open/restore already replaced the partial document.
        if (this.#rollbackBlock === null) return;
        this.#rollbackBlock = `The previous edit failed and could not be undone (${
          cause instanceof Error ? cause.message : String(cause)
        }). Reopen the presentation.`;
      },
    );
  }

  /** A newer replacement superseded the rollback and then failed: finish the rollback. */
  #resumeRollback(): void {
    if (this.#rollbackState !== null && this.#rollbackBlock !== null)
      this.#startRollback(this.#rollbackState);
  }

  #endRollback(): void {
    this.#rollbackBlock = null;
    this.#rollbackState = null;
  }

  #invalidateRestore(): void {
    this.#operation++;
    this.#requestedCursor = this.#cursor;
  }

  #commitHistory(label: string, source: ChangeSource): void {
    if (this.#externalHistory) {
      this.#externalHistory.commit(this.pres, {
        label,
        source,
        slideIndex: this.selection.slideIndex,
      });
      this.regroupHistory.prune(getSlides(this.pres));
      this.committedVersion = this.version;
    } else this.#snapshot(label);
  }

  #resetExternal(pres: PresentationData): void {
    const history = this.#externalHistory;
    if (!history) return;
    // New/Open must remain usable even if a host projection restore failed.
    if (this.#rollbackBlock !== null) history.reset(pres);
    else this.#atomic(() => history.reset(pres));
  }

  #restoreExternal(action: Parameters<ExternalHistoryAdapter['restore']>[0]): void {
    if (!this.#externalHistory) return;
    this.pres = this.#externalHistory.restore(action, this.pres);
    this.#invalidateRestore();
    this.#endRollback();
    // The host may rebuild shape IDs, so stale shape/cell selections and regroup
    // references must not survive a projection replacement.
    this.selection = {
      kind: 'none',
      slideIndex: Math.max(0, Math.min(this.selection.slideIndex, getSlides(this.pres).length - 1)),
    };
    this.regroupHistory.records = [];
    this.rememberedFills.clear();
    this.liveEditing = false;
    this.version++;
    this.committedVersion = this.version;
  }

  /** Capture before yielding: neither the document nor selection may drift. */
  #snapshot(label: string): void {
    this.committedVersion = this.version;
    const bytes = savePresentation(this.pres);
    const kept = this.#history.slice(0, this.#cursor + 1);
    this.regroupHistory.prune(getSlides(this.pres));
    kept.push({
      bytes,
      selection: this.selection,
      label,
      formerGroups: this.regroupHistory.records,
    });
    this.#history = kept.slice(-HISTORY_MAX);
    this.#cursor = this.#history.length - 1;
    this.#requestedCursor = this.#cursor;
  }

  /** Resolves to whether this restore replaced the document (a newer operation may win). */
  async #restore(index: number, selection?: Selection, rollback = false): Promise<boolean> {
    const snap = this.#history[index];
    if (!snap) return false;
    const operation = ++this.#operation;
    this.#requestedCursor = index;
    try {
      const pres = await loadPresentation(await snap.bytes);
      // New/Open, another restore, or an edit takes precedence over stale work.
      if (operation !== this.#operation) return false;
      this.pres = pres;
      this.#endRollback();
      this.regroupHistory.records = snap.formerGroups;
      this.selection = selection ?? snap.selection;
      this.liveEditing = false;
      this.#cursor = index;
      this.version++;
      this.committedVersion = this.version;
      this.dirty = true;
      return true;
    } catch (error) {
      if (operation === this.#operation) {
        this.#requestedCursor = this.#cursor;
        if (!rollback) this.#resumeRollback();
      }
      throw error;
    }
  }

  /** Discard the current gesture without adding an undo step or dropping redo. */
  async cancelLive(): Promise<void> {
    if (this.#externalHistory) {
      if (this.liveEditing) {
        this.#atomic(() => this.#restoreExternal('cancel'));
        this.dirty = this.#liveDirty ?? this.dirty;
        this.#liveDirty = undefined;
      }
      return;
    }
    if (this.liveEditing) await this.#restore(this.#cursor, this.selection);
  }

  async undo(): Promise<void> {
    if (this.#externalHistory) {
      if (this.canUndo) {
        this.#atomic(() => this.#restoreExternal('undo'));
        this.#liveDirty = undefined;
        this.dirty = true;
        this.onChange?.('user');
      }
      return;
    }
    if (this.#requestedCursor > 0 && (await this.#restore(this.#requestedCursor - 1)))
      this.onChange?.('user');
  }

  async redo(): Promise<void> {
    if (this.#externalHistory) {
      if (this.canRedo) {
        this.#atomic(() => this.#restoreExternal('redo'));
        this.#liveDirty = undefined;
        this.dirty = true;
        this.onChange?.('user');
      }
      return;
    }
    if (
      this.#requestedCursor < this.#history.length - 1 &&
      (await this.#restore(this.#requestedCursor + 1))
    )
      this.onChange?.('user');
  }

  /** Reproject a canonical change already accepted by the external host. */
  refreshFromExternalHistory(): void {
    if (!this.#externalHistory) throw new Error('External history is not configured.');
    if (this.liveEditing)
      throw new Error('Finish the current gesture before refreshing external history.');
    this.#atomic(() => this.#restoreExternal('rollback'));
    this.dirty = true;
    this.onChange?.('agent');
  }

  // --- Selection ---------------------------------------------------------
  select(sel: Selection): void {
    this.selection = sel;
  }

  selectSlide(index: number, options: { additive?: boolean; range?: boolean } = {}): void {
    const clamped = Math.max(0, Math.min(index, getSlides(this.pres).length - 1));
    const current = this.selection;
    const anchor = current.kind === 'slide' ? (current.anchorIndex ?? current.slideIndex) : clamped;
    let indices = [clamped];
    if (options.range) {
      indices = Array.from(
        { length: Math.abs(clamped - anchor) + 1 },
        (_, i) => Math.min(clamped, anchor) + i,
      );
    } else if (options.additive && current.kind === 'slide') {
      const selected = new Set(selectedSlideIndices(current));
      if (selected.has(clamped) && selected.size > 1) selected.delete(clamped);
      else selected.add(clamped);
      indices = [...selected].sort((a, b) => a - b);
    }
    this.selection = {
      kind: 'slide',
      slideIndex: indices.includes(clamped) ? clamped : indices[indices.length - 1]!,
      slideIndices: indices,
      anchorIndex: options.range ? anchor : clamped,
    };
  }

  selectShape(slideIndex: number, shapeId: number, additive = false): void {
    if (additive && this.selection.kind === 'shape' && this.selection.slideIndex === slideIndex) {
      const set = new Set(this.selection.shapeIds);
      if (set.has(shapeId)) set.delete(shapeId);
      else set.add(shapeId);
      this.selection = { kind: 'shape', slideIndex, shapeIds: [...set] };
    } else {
      this.selection = { kind: 'shape', slideIndex, shapeIds: [shapeId] };
    }
  }

  selectCell(slideIndex: number, shapeId: number, row: number, col: number, extend = false): void {
    const current = this.selection;
    if (
      extend &&
      current.kind === 'cell' &&
      current.slideIndex === slideIndex &&
      current.shapeId === shapeId
    ) {
      this.selection = { ...current, end: { row, col } };
      return;
    }
    this.selection = { kind: 'cell', slideIndex, shapeId, row, col };
  }

  clearShapeSelection(): void {
    this.selection = { kind: 'none', slideIndex: this.selection.slideIndex };
  }

  // --- IO ----------------------------------------------------------------
  /** `dirty`: the opened deck is not what the host has saved (a recovered copy). */
  async loadBytes(bytes: Uint8Array, name: string, dirty = false): Promise<void> {
    this.#invalidateRestore();
    const operation = this.#operation;
    let pres: PresentationData;
    try {
      pres = await loadPresentation(bytes);
    } catch (error) {
      if (operation === this.#operation) this.#resumeRollback();
      throw error;
    }
    if (operation !== this.#operation) return;
    this.#resetExternal(pres);
    this.pres = pres;
    this.#endRollback();
    this.fileName = name;
    this.regroupHistory.records = [];
    this.rememberedFills.clear();
    this.#history = [];
    this.#cursor = -1;
    this.selection = { kind: 'none', slideIndex: 0 };
    this.version++;
    this.dirty = dirty;
    this.liveEditing = false;
    this.#liveDirty = undefined;
    if (this.#externalHistory) this.committedVersion = this.version;
    else this.#snapshot('Open');
    this.onChange?.('user');
  }

  async toBytes(): Promise<Uint8Array> {
    // Never serialize a half-applied edit; save the state the rollback restores.
    if (this.#rollbackBlock !== null) {
      await this.#rollback;
      if (this.#rollbackBlock !== null) throw new Error(this.#rollbackBlock);
    }
    return savePresentation(this.pres);
  }

  markSaved(version: number): void {
    if (version === this.version) this.dirty = false;
  }

  resetBlank(): void {
    const pres = createInitial();
    this.#resetExternal(pres);
    this.#invalidateRestore();
    this.pres = pres;
    this.#endRollback();
    this.fileName = 'Untitled.pptx';
    this.regroupHistory.records = [];
    this.rememberedFills.clear();
    this.#history = [];
    this.#cursor = -1;
    this.selection = { kind: 'none', slideIndex: 0 };
    this.version++;
    this.dirty = false;
    this.liveEditing = false;
    this.#liveDirty = undefined;
    if (this.#externalHistory) this.committedVersion = this.version;
    else this.#snapshot('New');
    this.onChange?.('user');
  }
}

function createInitial(): PresentationData {
  const pres = createPresentation();
  addTitleSlide(pres, 'Untitled presentation');
  try {
    const defaults = JSON.parse(localStorage.getItem('office-grid-defaults') ?? 'null');
    if (
      defaults &&
      Number.isSafeInteger(defaults.x) &&
      defaults.x > 0 &&
      Number.isSafeInteger(defaults.y) &&
      defaults.y > 0 &&
      typeof defaults.snap === 'boolean'
    ) {
      setGridSpacing(pres, { x: defaults.x, y: defaults.y });
      setSnapToGrid(pres, defaults.snap);
    }
  } catch {
    /* Missing or invalid local defaults do not prevent creating a deck. */
  }
  return pres;
}
