# FormulaEngine Demo

This interactive demo showcases the capabilities of the FormulaEngine library with two comprehensive examples.

## Running the Demo

```bash
bun run dev
```

Then open your browser to `http://localhost:3000` (or the URL shown in the terminal).

## Demo Features

### 1. Events & Hooks Demo

An interactive demonstration of the events system and React hooks integration:

- **Live Spreadsheet Updates**: Watch how React hooks automatically update the UI when spreadsheet data changes
- **Real-time Event Logging**: See all events (cell changes, sheet operations) as they happen
- **Interactive Controls**: Add values, formulas, and new sheets to trigger various events
- **Dependency Tracking**: Observe how changing one cell automatically recalculates dependent cells

**Try these examples:**
- Set A1 to 100 and watch C1 update automatically (it's calculated as =A1+B1)
- Create formulas like `=A1*3` or `=SUM(A1:B2)`
- Add new sheets to see sheet-added events
- Clear cells to see how dependencies update

### 2. Full Spreadsheet Demo

A complete spreadsheet interface with rich example data:

**Example Data Includes:**
- **Product Inventory**: Laptops, mice, keyboards, monitors with prices and quantities
- **Automatic Calculations**: Total costs, tax calculations, final prices
- **Summary Analytics**: Total items, subtotals, grand totals, averages
- **Advanced Formulas**: SUM, AVERAGE, MAX, MIN, COUNT, IF, COUNTIF functions
- **Live Updates**: All calculations update automatically when you change values

**How to Use:**
1. Click any cell to select it
2. View the current formula/value in the formula bar at the top
3. Edit the formula or value and press Enter to update
4. Watch dependent cells automatically recalculate
5. Explore the live calculations display at the bottom

**Interactive Features:**
- Formula bar shows the selected cell address and current formula
- Type new formulas (start with =) or plain values
- Press Enter to confirm changes
- Real-time calculation display
- Comprehensive example data to explore

## Example Formulas to Try

### Basic Arithmetic
- `=B2*C2` (multiply price by quantity)
- `=D2*(1+E2)` (add tax to total)
- `=B2+C2+D2` (sum multiple cells)

### Functions
- `=SUM(B2:B5)` (sum a range)
- `=AVERAGE(B2:B5)` (average of range)
- `=MAX(F2:F5)` (maximum value)
- `=MIN(F2:F5)` (minimum value)
- `=COUNT(B2:B5)` (count numbers)

### Logical Functions
- `=IF(B2>100,"Expensive","Affordable")` (conditional logic)
- `=IF(COUNTIF(A2:A5,"Laptop")>0,"Has Laptop","No Laptop")` (conditional counting)

### Text Functions
- `=CONCATENATE("Product: ",A2)` (combine text)
- `=UPPER(A2)` (convert to uppercase)
- `=LEN(A2)` (text length)

## Architecture Highlights

### Events System
- **Real-time Updates**: All changes emit events that React hooks can subscribe to
- **Type Safety**: Full TypeScript support for all event types
- **Memory Management**: Automatic cleanup prevents memory leaks
- **Event Ordering**: Ensures parent cell changes emit before dependent changes

### React Integration
- **useSpreadsheet**: Subscribe to entire sheet changes
- **useCell**: Subscribe to individual cell changes
- **useSpreadsheetRange**: Subscribe to range changes
- **useFormulaEngineEvents**: Custom event handling

### Formula Engine Features
- **150+ Functions**: Math, logical, text, lookup, info, and array functions
- **Dependency Tracking**: Automatic recalculation of dependent cells
- **Array Formulas**: Support for array operations and spilling
- **Error Handling**: Comprehensive error types (#DIV/0!, #REF!, etc.)
- **Named Expressions**: Support for named variables and formulas
- **Multiple Sheets**: Full multi-sheet support

## Development

The demo is built with:
- **React 18** with TypeScript
- **TailwindCSS** for styling
- **Shadcn/ui** for UI components
- **@anocca-pub/components** for the spreadsheet grid
- **Bun** for development and building

### File Structure
```
demo/
├── App.tsx                 # Main app with tab navigation
├── EventsDemo.tsx          # Events and React hooks demonstration
├── FullSpreadsheetDemo.tsx # Complete spreadsheet with example data
├── components/ui/          # Reusable UI components
├── lib/                    # Utility functions
└── styles/                 # CSS styles
```

### Key Features Demonstrated
1. **Real-time Reactivity**: Changes to the engine immediately update the UI
2. **Formula Editing**: Live formula bar with Enter-to-commit functionality
3. **Rich Example Data**: Realistic business data with complex formulas
4. **Event Monitoring**: Live event log showing all engine operations
5. **Error Handling**: Graceful handling of invalid formulas and references
6. **Performance**: Efficient updates even with many dependent cells

This demo provides a comprehensive showcase of FormulaEngine's capabilities and serves as a practical example for building spreadsheet applications.