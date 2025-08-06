# FormulaEngine Events & React Integration

This document describes the event system and React hooks integration added to FormulaEngine.

## Events System

FormulaEngine now includes a comprehensive event system that allows you to subscribe to various changes happening within the engine.

### Available Events

- `cell-changed`: Emitted when a cell value changes
- `sheet-added`: Emitted when a new sheet is added
- `sheet-removed`: Emitted when a sheet is removed
- `sheet-renamed`: Emitted when a sheet is renamed
- `formula-calculated`: Emitted when a formula is evaluated
- `dependency-updated`: Emitted when dependencies are updated
- `named-expression-added`: Emitted when named expressions are added
- `named-expression-changed`: Emitted when named expressions are modified
- `named-expression-removed`: Emitted when named expressions are removed

### Basic Usage

```typescript
import { FormulaEngine } from 'formula-engine';

const engine = FormulaEngine.buildEmpty();
const sheetName = engine.addSheet('MySheet');
const sheetId = engine.getSheetId(sheetName);

// Subscribe to cell changes
const unsubscribe = engine.on('cell-changed', (event) => {
  console.log(`Cell ${event.address} changed from ${event.oldValue} to ${event.newValue}`);
});

// Make some changes
engine.setCellContents({ sheet: sheetId, col: 0, row: 0 }, 42);
// Logs: Cell {sheet: 0, col: 0, row: 0} changed from undefined to 42

// Clean up
unsubscribe();
```

### Event Properties

#### cell-changed
```typescript
{
  address: SimpleCellAddress;
  oldValue: CellValue;
  newValue: CellValue;
}
```

#### sheet-added
```typescript
{
  sheetId: number;
  sheetName: string;
}
```

#### sheet-removed
```typescript
{
  sheetId: number;
  sheetName: string;
}
```

#### sheet-renamed
```typescript
{
  sheetId: number;
  oldName: string;
  newName: string;
}
```

## React Integration

FormulaEngine provides several React hooks for seamless integration with React applications.

### useSpreadsheet Hook

Subscribe to an entire spreadsheet sheet and get reactive updates.

```tsx
import { FormulaEngine, useSpreadsheet } from 'formula-engine';

function SpreadsheetComponent({ engine, sheetName }: { engine: FormulaEngine, sheetName: string }) {
  const { spreadsheet, isLoading, error } = useSpreadsheet(engine, sheetName);

  if (isLoading) return <div>Loading...</div>;
  if (error) return <div>Error: {error.message}</div>;

  return (
    <div>
      {Array.from(spreadsheet.entries()).map(([address, value]) => (
        <div key={address}>
          {address}: {JSON.stringify(value)}
        </div>
      ))}
    </div>
  );
}
```

### useCell Hook

Subscribe to a single cell value with automatic updates.

```tsx
import { FormulaEngine, useCell } from 'formula-engine';

function CellDisplay({ engine, sheetName, cellAddress }: { 
  engine: FormulaEngine, 
  sheetName: string, 
  cellAddress: string 
}) {
  const { value, isLoading, error } = useCell(engine, sheetName, cellAddress);

  if (isLoading) return <div>Loading...</div>;
  if (error) return <div>Error: {error.message}</div>;

  return <div>Cell {cellAddress}: {JSON.stringify(value)}</div>;
}
```

### useSpreadsheetRange Hook

Subscribe to a specific range of cells.

```tsx
import { FormulaEngine, useSpreadsheetRange } from 'formula-engine';

function RangeDisplay({ engine, sheetName, range }: { 
  engine: FormulaEngine, 
  sheetName: string, 
  range: string 
}) {
  const { rangeData, isLoading, error } = useSpreadsheetRange(engine, sheetName, range);

  if (isLoading) return <div>Loading...</div>;
  if (error) return <div>Error: {error.message}</div>;

  return (
    <div>
      <h3>Range {range}:</h3>
      {Array.from(rangeData.entries()).map(([address, value]) => (
        <div key={address}>
          {address}: {JSON.stringify(value)}
        </div>
      ))}
    </div>
  );
}
```

### useFormulaEngineEvents Hook

Subscribe to FormulaEngine events for custom event handling.

```tsx
import { FormulaEngine, useFormulaEngineEvents } from 'formula-engine';

function EventLogger({ engine }: { engine: FormulaEngine }) {
  const [events, setEvents] = useState<any[]>([]);

  useFormulaEngineEvents(engine, {
    onCellChanged: (event) => {
      setEvents(prev => [...prev, { type: 'cell-changed', ...event }]);
    },
    onSheetAdded: (event) => {
      setEvents(prev => [...prev, { type: 'sheet-added', ...event }]);
    },
    // ... other event handlers
  });

  return (
    <div>
      <h3>Event Log:</h3>
      {events.map((event, index) => (
        <div key={index}>
          {event.type}: {JSON.stringify(event)}
        </div>
      ))}
    </div>
  );
}
```

### Hook Options

All hooks support options for customization:

```tsx
const options = {
  autoUpdate: true,    // Enable automatic re-rendering (default: true)
  debounceMs: 100      // Debounce updates by 100ms (default: 0)
};

const { spreadsheet } = useSpreadsheet(engine, sheetName, options);
const { value } = useCell(engine, sheetName, 'A1', options);
const { rangeData } = useSpreadsheetRange(engine, sheetName, 'A1:C10', options);
```

## Demo

Run the demo to see the events and React integration in action:

```bash
bun run dev
```

Navigate to the "Events & Hooks Demo" tab to see:
- Live spreadsheet updates using React hooks
- Real-time event logging
- Interactive controls to trigger various events
- Dependency tracking and formula recalculation

## Features

### Event Features
- ✅ Comprehensive event coverage
- ✅ Synchronous event emission
- ✅ Proper event ordering for dependent cells
- ✅ Unsubscribe functionality
- ✅ Multiple listeners per event
- ✅ Event data integrity

### React Hook Features
- ✅ Automatic re-rendering on changes
- ✅ Error handling and loading states
- ✅ Debouncing support
- ✅ Memory leak prevention
- ✅ Proper cleanup on unmount
- ✅ TypeScript support

### Testing
- ✅ 20 comprehensive event system tests
- ✅ 18 React hooks integration tests
- ✅ All tests passing
- ✅ Edge case coverage

## Architecture Notes

- Events are emitted synchronously to ensure consistency
- React hooks use React's built-in state management and effects
- Dependencies are properly tracked and cleaned up
- Event ordering ensures parent cell changes emit before dependent cell changes
- Hooks are designed to be efficient and avoid unnecessary re-renders

This implementation provides a solid foundation for building reactive spreadsheet applications with FormulaEngine.