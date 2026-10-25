import type { Worksheet } from './Worksheet.ts';

export class WorksheetEvent extends Event {
  readonly worksheet: Worksheet;
  readonly index: number;
  readonly previousName: string | undefined;

  constructor(type: string, { worksheet, index, previousName }: { worksheet: Worksheet; index: number; previousName?: string }) {
    super(type);

    this.worksheet = worksheet;
    this.index = index;
    this.previousName = previousName;
  }
}
