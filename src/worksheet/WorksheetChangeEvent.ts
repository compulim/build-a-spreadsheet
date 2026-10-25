import type { Worksheet, WorksheetChange, WorksheetUndo } from './Worksheet.ts';

/** Dispatched once per `Worksheet#edit()` transaction, with all of its changes in order. */
export class WorksheetChangeEvent extends Event {
  readonly changes: readonly WorksheetChange[];
  /** `true` if the transaction loaded CSV, which is not meant to be undoable. */
  readonly load: boolean;
  /** Undoes the transaction, `undefined` if the event is from an undo or redo. */
  readonly undo: WorksheetUndo | undefined;
  readonly worksheet: Worksheet;

  constructor(
    type: string,
    {
      changes,
      load,
      undo,
      worksheet
    }: { changes: readonly WorksheetChange[]; load: boolean; undo?: WorksheetUndo | undefined; worksheet: Worksheet }
  ) {
    super(type);

    this.changes = changes;
    this.load = load;
    this.undo = undo;
    this.worksheet = worksheet;
  }
}
