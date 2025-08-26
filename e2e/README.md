# End-to-End Tests

This directory contains Playwright tests for the Formula Engine Excel Demo.

## Prerequisites

1. Start the development server:
   ```bash
   bun run dev
   ```
   (Should be running on http://localhost:3000)

2. Install Playwright browsers (if not already done):
   ```bash
   npx playwright install
   ```

## Running Tests

### Run all tests
```bash
bun run test:e2e
```

### Run tests with UI mode (interactive)
```bash
bun run test:e2e:ui
```

### Run tests in debug mode
```bash
bun run test:e2e:debug
```

### Run specific test file
```bash
npx playwright test excel-demo.spec.ts
```

### Run specific test
```bash
npx playwright test -g "should add new sheets"
```

## Test Coverage

The `excel-demo.spec.ts` file covers:

### Core Functionality
- ✅ Initial state display
- ✅ Sheet management (add, rename, delete, switch)
- ✅ Named expressions (add global/sheet-scoped, delete)
- ✅ Save/load functionality with localStorage
- ✅ Unsaved changes tracking

### Advanced Features
- ✅ Formula entry and evaluation
- ✅ Formula bar integration
- ✅ Named expressions in formulas
- ✅ State persistence across page reloads

### UI Interactions
- ✅ Button states and visibility
- ✅ Panel expand/collapse
- ✅ Form validation
- ✅ Hover interactions

## Test Data

Tests use `data-testid` attributes for reliable element selection:
- `[data-testid="spreadsheet-cell"]` - Individual spreadsheet cells
- Form elements use placeholder text and ARIA labels

## Debugging

### Console Log Debugging

For debugging failing tests, run the dev server manually to see console output:

**Terminal 1 (Dev Server):**
```bash
bun run dev
```

**Terminal 2 (Tests):**
```bash
bun run test:e2e
# or run specific test:
npx playwright test -g "should enter and evaluate basic formulas"
```

Any `console.log()` statements in your tests or the application will appear in the dev server output (Terminal 1), making it easy to debug what's happening during test execution.

### Other Debugging Options

1. **View test results**: After running tests, open `playwright-report/index.html`
2. **Screenshots**: Failed tests automatically capture screenshots
3. **Videos**: Failed tests record videos for debugging
4. **Traces**: Use `npx playwright show-trace` to view detailed execution traces
5. **Headed mode**: Run `bun run test:e2e:headed` to see browser actions
6. **Debug mode**: Run `bun run test:e2e:debug` for step-by-step debugging
