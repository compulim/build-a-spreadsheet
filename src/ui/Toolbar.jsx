import {
  BoldIcon,
  CircleSlashIcon,
  DownloadIcon,
  FileDirectoryOpenFillIcon,
  FileZipIcon,
  ImageIcon,
  ItalicIcon,
  NumberIcon,
  RedoIcon,
  SquareIcon,
  StrikethroughIcon,
  TableIcon,
  TypographyIcon,
  UndoIcon
} from '@primer/octicons-react';
import { ActionBar, ActionList, ActionMenu, SegmentedControl, Text, TextInput, ToggleSwitch } from '@primer/react';
import { useSyncExternalStore } from 'react';
import { HIGHLIGHTS } from '../highlights.js';
import { MarkerIcon, UnderlineIcon } from './icons.jsx';

/** @typedef {import('../worksheet/Worksheet.js').CellAlign} CellAlign */
/** @typedef {import('../worksheet/Worksheet.js').CellNumberFormat} CellNumberFormat */
/** @typedef {'bold' | 'italic' | 'underline' | 'strikethrough'} CellStyle */
/** @typedef {'left' | 'right' | 'top' | 'bottom' | 'outside' | 'all' | 'none'} BorderPreset */

/**
 * @typedef {{
 *   align: CellAlign | undefined;
 *   canRedo: boolean;
 *   canUndo: boolean;
 *   format: CellNumberFormat | undefined;
 *   highlight: string | undefined;
 *   locked: boolean;
 *   merged: boolean;
 *   nameBox: string;
 *   status: string;
 *   styles: Record<CellStyle, boolean>;
 *   valueBox: string;
 * }} ToolbarState
 */

/**
 * @typedef {{
 *   merge(): void;
 *   open(): void;
 *   print(): void;
 *   redo(): void;
 *   saveCSV(): void;
 *   saveZip(): void;
 *   setAlign(align: CellAlign | undefined): void;
 *   setBorder(preset: BorderPreset): void;
 *   setFormat(format: CellNumberFormat | undefined): void;
 *   setHighlight(name: string | undefined): void;
 *   setLocked(locked: boolean): void;
 *   toggleStyle(style: CellStyle): void;
 *   undo(): void;
 *   unformat(): void;
 * }} ToolbarActions
 */

/**
 * @typedef {{
 *   getState(): ToolbarState;
 *   subscribe(listener: () => void): () => void;
 * }} ToolbarStore
 */

/** Segments of the alignment control, the first segment is the default (type-based) alignment. */
const ALIGNS = /** @type {const} */ ([
  { align: undefined, label: 'Auto', description: 'Align by value type' },
  { align: 'left', label: 'Left', description: 'Align left (Ctrl+L)' },
  { align: 'center', label: 'Center', description: 'Align center (Ctrl+E)' },
  { align: 'right', label: 'Right', description: 'Align right (Ctrl+R)' }
]);

/** @type {readonly { format: CellNumberFormat | undefined; label: string; description?: string }[]} */
const FORMATS = [
  { format: undefined, label: 'General' },
  { format: 'currency', label: 'Currency', description: 'Accounting' },
  { format: 'percent', label: 'Percentage', description: '1 is 100%' },
  { format: 'date', label: 'Date', description: 'Unix epoch in milliseconds' },
  { format: 'datetime', label: 'Date and time', description: 'Unix epoch in milliseconds' },
  { format: 'duration', label: 'Duration', description: 'Milliseconds' }
];

/** Segments of the marker control, the first segment removes the highlight. */
const MARKERS = [
  { name: undefined, label: 'No highlight', icon: CircleSlashIcon },
  ...HIGHLIGHTS.map(({ background, color, label, name }) => ({
    name,
    label,
    icon: <MarkerIcon background={background} color={color} />
  }))
];

/** @type {readonly { preset: BorderPreset; label: string }[]} */
const BORDERS = [
  { preset: 'left', label: 'Left' },
  { preset: 'right', label: 'Right' },
  { preset: 'top', label: 'Top' },
  { preset: 'bottom', label: 'Bottom' },
  { preset: 'outside', label: 'Surround' },
  { preset: 'all', label: 'Surround and inside' },
  { preset: 'none', label: 'Unset surround and inside' }
];

/**
 * `returnFocusRef` is where focus goes after a menu closes, i.e. the grid.
 *
 * @param {{ actions: ToolbarActions; returnFocusRef: import('react').RefObject<HTMLElement>; store: ToolbarStore }} props
 */
export function Toolbar({ actions, returnFocusRef, store }) {
  const state = useSyncExternalStore(store.subscribe, store.getState);
  const formatLabel = FORMATS.find(({ format }) => format === state.format)?.label ?? 'General';

  return (
    <div className="toolbar">
      <div className="toolbar__row">
        <ActionBar aria-label="Commands" className="toolbar__action-bar" flush size="small">
          <ActionBar.Group>
            <ActionBar.IconButton
              aria-label="Open…"
              description="Open a CSV file (as a new sheet) or a ZIP file of CSV files (as a workbook)"
              icon={FileDirectoryOpenFillIcon}
              onClick={actions.open}
            />
            <ActionBar.IconButton
              aria-label="Save sheet (CSV)"
              description="Save the current sheet as a CSV file"
              icon={DownloadIcon}
              onClick={actions.saveCSV}
            />
            <ActionBar.IconButton
              aria-label="Save workbook (ZIP)"
              description="Save all sheets as CSV files in a ZIP file"
              icon={FileZipIcon}
              onClick={actions.saveZip}
            />
            <ActionBar.IconButton
              aria-label="Print selected cells"
              description="Save the selected cells as a PNG image for printing"
              icon={ImageIcon}
              onClick={actions.print}
            />
          </ActionBar.Group>
          <ActionBar.Divider />
          <ActionBar.Group>
            <ActionBar.IconButton
              aria-label="Undo"
              description="Undo (Ctrl+Z)"
              disabled={!state.canUndo}
              icon={UndoIcon}
              onClick={actions.undo}
            />
            <ActionBar.IconButton
              aria-label="Redo"
              description="Redo (Ctrl+Y)"
              disabled={!state.canRedo}
              icon={RedoIcon}
              onClick={actions.redo}
            />
          </ActionBar.Group>
          <ActionBar.Divider />
          <ActionBar.Group>
            <ActionBar.IconButton
              aria-label="Bold"
              aria-pressed={state.styles.bold}
              description="Bold (Ctrl+B)"
              icon={BoldIcon}
              onClick={() => actions.toggleStyle('bold')}
            />
            <ActionBar.IconButton
              aria-label="Italic"
              aria-pressed={state.styles.italic}
              description="Italic (Ctrl+I)"
              icon={ItalicIcon}
              onClick={() => actions.toggleStyle('italic')}
            />
            <ActionBar.IconButton
              aria-label="Underline"
              aria-pressed={state.styles.underline}
              description="Underline (Ctrl+U)"
              icon={UnderlineIcon}
              onClick={() => actions.toggleStyle('underline')}
            />
            <ActionBar.IconButton
              aria-label="Strikethrough"
              aria-pressed={state.styles.strikethrough}
              icon={StrikethroughIcon}
              onClick={() => actions.toggleStyle('strikethrough')}
            />
          </ActionBar.Group>
          <ActionBar.Divider />
          <ActionBar.Group>
            <ActionBar.Menu
              aria-label="Border"
              icon={SquareIcon}
              items={BORDERS.map(({ preset, label }) => ({ label, onClick: () => actions.setBorder(preset) }))}
              returnFocusRef={returnFocusRef}
              // ActionBar.Menu does not inherit the size of ActionBar, a taller button makes every item overflow.
              size="small"
            />
            <ActionBar.IconButton
              aria-label="Merge"
              aria-pressed={state.merged}
              description="Merge the selected cells of each row, click again to unmerge (Ctrl+M)"
              icon={TableIcon}
              onClick={actions.merge}
            />
            <ActionBar.IconButton
              aria-label="Unformat"
              description="Remove all formatting"
              icon={TypographyIcon}
              onClick={actions.unformat}
            />
          </ActionBar.Group>
        </ActionBar>
        <SegmentedControl aria-label="Alignment" onChange={index => actions.setAlign(ALIGNS[index].align)} size="small">
          {ALIGNS.map(({ align, description, label }) => (
            <SegmentedControl.Button key={label} selected={align === state.align} title={description}>
              {label}
            </SegmentedControl.Button>
          ))}
        </SegmentedControl>
        <SegmentedControl
          aria-label="Highlight"
          onChange={index => actions.setHighlight(MARKERS[index].name)}
          size="small"
        >
          {MARKERS.map(({ icon, label, name }) => (
            <SegmentedControl.IconButton
              aria-label={label}
              icon={icon}
              key={label}
              selected={name === state.highlight}
            />
          ))}
        </SegmentedControl>
        <ActionMenu>
          <ActionMenu.Button aria-label={`Number format: ${formatLabel}`} leadingVisual={NumberIcon} size="small">
            {formatLabel}
          </ActionMenu.Button>
          <ActionMenu.Overlay returnFocusRef={returnFocusRef} width="auto">
            <ActionList selectionVariant="single">
              {FORMATS.map(({ description, format, label }) => (
                <ActionList.Item
                  key={label}
                  onSelect={() => actions.setFormat(format)}
                  selected={format === state.format}
                >
                  {label}
                  {description && <ActionList.Description variant="block">{description}</ActionList.Description>}
                </ActionList.Item>
              ))}
            </ActionList>
          </ActionMenu.Overlay>
        </ActionMenu>
        <div className="toolbar__lock" title="Lock editing of protected rows and columns">
          <Text id="lock-label" size="small" weight="semibold">
            Lock
          </Text>
          <ToggleSwitch aria-labelledby="lock-label" checked={state.locked} onChange={actions.setLocked} size="small" />
        </div>
      </div>
      <div className="toolbar__row">
        <TextInput
          aria-label="Selected cells"
          className="toolbar__name-box"
          readOnly
          size="small"
          value={state.nameBox}
        />
        <TextInput
          aria-label="Value of active cell"
          block
          className="toolbar__value-box"
          readOnly
          size="small"
          value={state.valueBox}
        />
        <Text className="toolbar__status" role="status" size="small">
          {state.status}
        </Text>
      </div>
    </div>
  );
}
