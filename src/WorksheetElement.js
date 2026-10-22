import { SpreadsheetGrid } from './grid.js';
import { Worksheet } from './workbook/Worksheet.js';

/** @type {(element: WorksheetElement) => SpreadsheetGrid} */
let getGrid;

/**
 * `<compulim-worksheet>`, a grid showing a `Worksheet` it owns.
 *
 * The only public member is `worksheet`, a newly constructed `Worksheet` owned by the element. Number of lockable leading
 * rows and columns are read from the `locked-row-count` and `locked-column-count` attributes when first connected.
 */
export class WorksheetElement extends HTMLElement {
  #worksheet = new Worksheet();

  /** @type {SpreadsheetGrid | undefined} */
  #grid;

  static {
    getGrid = element => element.#getGrid();
  }

  /** @returns {Worksheet} */
  get worksheet() {
    return this.#worksheet;
  }

  connectedCallback() {
    this.#getGrid();
  }

  #getGrid() {
    if (!this.#grid) {
      this.#grid = new SpreadsheetGrid(this, {
        lockedColumnCount: Number(this.getAttribute('locked-column-count')) || 0,
        lockedRowCount: Number(this.getAttribute('locked-row-count')) || 0
      });
      this.#grid.setWorksheet(this.#worksheet);
    }

    return this.#grid;
  }
}

customElements.get('compulim-worksheet') || customElements.define('compulim-worksheet', WorksheetElement);

/**
 * Returns the grid controller of a `<compulim-worksheet>`. For use by the app only, not part of the element's public
 * interface.
 *
 * @param {WorksheetElement} element
 * @returns {SpreadsheetGrid}
 */
export function getWorksheetGrid(element) {
  return getGrid(element);
}
