import { test, expect } from '@playwright/test';

test.describe('Excel Demo', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/excel');
    // Wait for the page to fully load
    await expect(page.locator('h1')).toContainText('FormulaEngine Multi-Workbook Demo');
  });

  test('should display initial state correctly', async ({ page }) => {
    // Check main title
    await expect(page.locator('h1')).toContainText('FormulaEngine Multi-Workbook Demo');
    
    // Check that we have one initial workbook
    await expect(page.locator('[data-testid="total-workbooks-count"]')).toContainText('Workbooks: 1');
    
    // Check save button is in "Saved" state initially
    await expect(page.locator('button:has-text("Saved")')).toBeVisible();
    
    // Check sheet tab is present in the workbook
    await expect(page.locator('[data-testid="sheet-tab-Sheet1"]')).toBeVisible();
  });

  test('should add new sheets', async ({ page }) => {
    // Click the add sheet button for Workbook1 using data-testid
    await page.locator('[data-testid="add-sheet-Workbook1"]').click();
    
    // Wait a moment for the UI to update
    await page.waitForTimeout(100);
    
    // Check that Sheet2 tab is present and active
    await expect(page.locator('[data-testid="sheet-tab-Sheet2"]')).toBeVisible();
    
    // Check that save button shows unsaved changes
    await expect(page.locator('[data-testid="save-button"]')).toContainText('Save Changes');
  });

  test('should rename sheets', async ({ page }) => {
    // Hover over the sheet tab to reveal edit button
    await page.locator('[data-testid="sheet-tab-Sheet1"]').hover();
    
    // Click the edit button
    await page.locator('[data-testid="rename-sheet-Sheet1"]').click();
    
    // Wait for edit mode to activate
    await page.waitForTimeout(100);
    
    // Type new name in the input field that appears
    await page.locator('[data-testid="rename-sheet-input-Sheet1"]').fill('My Sheet');
    
    // Press Enter to confirm
    await page.locator('[data-testid="rename-sheet-input-Sheet1"]').press('Enter');
    
    // Check that sheet was renamed
    await expect(page.locator('[data-testid="sheet-tab-My Sheet"]')).toBeVisible();
    
    // Check that save button shows unsaved changes
    await expect(page.locator('[data-testid="save-button"]')).toContainText('Save Changes');
  });

  test('should delete sheets', async ({ page }) => {
    // Set up data on Sheet1
    await page.locator('[data-testid="spreadsheet-cell-A1"]').dblclick();
    await page.locator('[data-testid="spreadsheet-cell-input-A1"]').fill('100');
    await page.keyboard.press('Enter');
    
    // Add a second sheet
    await page.locator('[data-testid="add-sheet-Workbook1"]').click();
    
    // On Sheet2, create a formula that references Sheet1
    await page.locator('[data-testid="spreadsheet-cell-B1"]').dblclick();
    await page.locator('[data-testid="spreadsheet-cell-input-B1"]').fill('=Sheet1!A1*2');
    await page.keyboard.press('Enter');
    
    // Verify cross-sheet formula works
    await expect(page.locator('[data-testid="spreadsheet-cell-B1"]')).toContainText('200');
    
    // Hover over Sheet1 to reveal delete button
    await page.locator('[data-testid="sheet-tab-Sheet1"]').hover();
    
    // Handle the confirmation dialog
    page.on('dialog', dialog => dialog.accept());
    
    // Click delete button
    await page.locator('[data-testid="delete-sheet-Sheet1"]').click();
    
    // Check that only Sheet2 remains and is active
    await expect(page.locator('[data-testid="sheet-tab-Sheet2"]')).toBeVisible();
    
    // Verify that formulas referencing deleted sheet show error
    await expect(page.locator('[data-testid="spreadsheet-cell-B1"]')).toContainText('#REF!');
    
    // Check that save button shows unsaved changes
    await expect(page.locator('button:has-text("Save Changes")')).toBeVisible();
  });

  test('should remove tables when their sheet is deleted', async ({ page }) => {
    // Create a table on Sheet1
    await page.locator('[data-testid="spreadsheet-cell-A1"]').dblclick();
    await page.locator('[data-testid="spreadsheet-cell-input-A1"]').fill('Product');
    await page.keyboard.press('Enter');
    
    await page.locator('[data-testid="spreadsheet-cell-B1"]').dblclick();
    await page.locator('[data-testid="spreadsheet-cell-input-B1"]').fill('Price');
    await page.keyboard.press('Enter');
    
    // Select range A1:B1 to create table
    await page.locator('[data-testid="spreadsheet-cell-A1"]').click();
    await page.keyboard.down('Shift');
    await page.locator('[data-testid="spreadsheet-cell-B1"]').click();
    await page.keyboard.up('Shift');
    
    // Create table
    await page.locator('[data-testid="create-table-button"]').click();
    
    // Open Expressions & Tables panel to verify table exists
    await page.locator('[data-testid="named-expressions-toggle"]').click();
    
    // Verify table appears in the Expressions & Tables panel list (be specific)
    await expect(page.locator('[data-testid="table-Table1"]')).toBeVisible();
    
    // Add a second sheet
    await page.locator('[data-testid="add-sheet-Workbook1"]').click();
    
    // Delete Sheet1 (which contains the table)
    await page.locator('[data-testid="sheet-tab-Sheet1"]').hover();
    
    // Handle the confirmation dialog
    page.on('dialog', dialog => dialog.accept());
    
    await page.locator('[data-testid="delete-sheet-Sheet1"]').click();
    
    // Wait a moment for the deletion to process and UI to update
    await page.waitForTimeout(500);
    

    
    // Verify table is removed from the UI list in the Expressions & Tables panel
    await expect(page.locator('[data-testid="table-Table1"]')).not.toBeVisible();
    
    // Verify no tables are shown in the panel
    await expect(page.locator('[data-testid="tables-list"]').locator('text=No tables')).toBeVisible();
  });

  test('should open and close Named Expressions panel', async ({ page }) => {
    // Click Named Expressions button
    await page.locator('[data-testid="named-expressions-toggle"]').click();
    
    // Check that the panel is visible
    await expect(page.locator('text=Add Named Expression')).toBeVisible();
    await expect(page.locator('[data-testid="named-expressions-section"]')).toBeVisible();
    await expect(page.locator('[data-testid="tables-section"]')).toBeVisible();
    
    // Click again to close
    await page.locator('[data-testid="named-expressions-toggle"]').click();
    
    // Check that panel is hidden
    await expect(page.locator('text=Add Named Expression')).not.toBeVisible();
  });

  test('should add global named expressions', async ({ page }) => {
    // Open Named Expressions panel
    await page.locator('[data-testid="named-expressions-toggle"]').click();
    
    // Fill in the form (scope defaults to global)
    await page.locator('[data-testid="expression-name-input"]').fill('TAX_RATE');
    await page.locator('[data-testid="expression-formula-input"]').fill('0.08');
    await page.locator('[data-testid="expression-scope-select"]').selectOption('global');
    
    // Click Add button
    await page.locator('[data-testid="add-expression-button"]').click();
    
    // Check that the named expression appears in the list with Global badge
    await expect(page.locator('[data-testid="expression-name-TAX_RATE"]')).toBeVisible();
    await expect(page.locator('[data-testid="expression-formula-TAX_RATE"]')).toContainText('0.08');
    await expect(page.locator('[data-testid="expression-scope-TAX_RATE"]')).toContainText('Global');
    
    // Check that save button shows unsaved changes
    await expect(page.locator('button:has-text("Save Changes")')).toBeVisible();
  });

  test('should add sheet-scoped named expressions', async ({ page }) => {
    // Open Named Expressions panel
    await page.locator('[data-testid="named-expressions-toggle"]').click();
    
    // Fill in the form for sheet-scoped expression
    await page.locator('[data-testid="expression-name-input"]').fill('LOCAL_RATE');
    await page.locator('[data-testid="expression-formula-input"]').fill('0.05');
    await page.locator('[data-testid="expression-scope-select"]').selectOption('sheet');
    
    // Click Add button
    await page.locator('[data-testid="add-expression-button"]').click();
    
    // Check that the named expression appears in the list with sheet scope badge
    await expect(page.locator('[data-testid="expression-name-LOCAL_RATE"]')).toBeVisible();
    await expect(page.locator('[data-testid="expression-formula-LOCAL_RATE"]')).toContainText('0.05');
    await expect(page.locator('[data-testid="expression-scope-LOCAL_RATE"]')).toContainText('Workbook1 → Sheet1');
    
    // Check that save button shows unsaved changes
    await expect(page.locator('button:has-text("Save Changes")')).toBeVisible();
  });

  test('should delete named expressions', async ({ page }) => {
    // First add a global named expression
    await page.locator('[data-testid="named-expressions-toggle"]').click();
    await page.locator('[data-testid="expression-name-input"]').fill('TEST_RATE');
    await page.locator('[data-testid="expression-formula-input"]').fill('0.1');
    await page.locator('[data-testid="expression-scope-select"]').selectOption('global');
    await page.locator('[data-testid="add-expression-button"]').click();
    
    // Verify it was added
    await expect(page.locator('[data-testid="expression-name-TEST_RATE"]')).toBeVisible();
    
    // Click the delete button for the TEST_RATE named expression
    await page.locator('[data-testid="delete-global-named-expression-TEST_RATE"]').click();
    
    // Verify it was deleted
    await expect(page.locator('[data-testid="expression-name-TEST_RATE"]')).not.toBeVisible();
  });

  test('should save and maintain state', async ({ page }) => {
    // Add a sheet
    await page.locator('[data-testid="add-sheet-Workbook1"]').click();
    
    // Add a named expression
    await page.locator('[data-testid="named-expressions-toggle"]').click();
    await page.locator('[data-testid="expression-name-input"]').fill('SAVE_TEST');
    await page.locator('[data-testid="expression-formula-input"]').fill('42');
    await page.locator('[data-testid="add-expression-button"]').click();
    
    // Save changes
    await page.locator('[data-testid="save-button"]').click();
    
    // Check that save button shows "Saved"
    await expect(page.locator('[data-testid="save-button"]')).toContainText('Saved');
    
    // Reload the page
    await page.reload();
    await expect(page.locator('h1')).toContainText('FormulaEngine Multi-Workbook Demo');
    
    // Check that state was preserved - should have 2 sheets in Workbook1
    await expect(page.locator('[data-testid="sheet-count-Workbook1"]')).toContainText('Sheets: 2');
    
    // Check named expression was preserved
    await page.locator('[data-testid="named-expressions-toggle"]').click();
    await expect(page.locator('[data-testid="expression-name-SAVE_TEST"]')).toBeVisible();
    await expect(page.locator('[data-testid="expression-formula-SAVE_TEST"]')).toContainText('42');
  });

  test('should show unsaved changes indicator', async ({ page }) => {
    // Initially should show "Saved"
    await expect(page.locator('[data-testid="save-button"]')).toContainText('Saved');
    await expect(page.locator('[data-testid="unsaved-changes-indicator"]')).not.toBeVisible();
    
    // Make a change (add sheet)
    await page.locator('[data-testid="add-sheet-Workbook1"]').click();
    
    // Should now show unsaved changes
    await expect(page.locator('[data-testid="save-button"]')).toContainText('Save Changes');
    await expect(page.locator('[data-testid="unsaved-changes-indicator"]')).toBeVisible();
    
    // Save changes
    await page.locator('[data-testid="save-button"]').click();
    
    // Should be back to saved state
    await expect(page.locator('[data-testid="save-button"]')).toContainText('Saved');
    await expect(page.locator('[data-testid="unsaved-changes-indicator"]')).not.toBeVisible();
  });

  test('should prevent deleting the last sheet', async ({ page }) => {
    // Verify we start with 1 sheet in Workbook1
    await expect(page.locator('[data-testid="sheet-count-Workbook1"]')).toContainText('Sheets: 1');
    
    // Hover over the sheet tab
    await page.locator('[data-testid="sheet-tab-Sheet1"]').hover();
    
    // The delete button should not be present for the last sheet
    await expect(page.locator('[data-testid="delete-sheet-Sheet1"]')).not.toBeVisible();
  });

  test('should switch between sheets', async ({ page }) => {
    // Set up different content on Sheet1
    await page.locator('[data-testid="spreadsheet-cell-A1"]').dblclick();
    await page.locator('[data-testid="spreadsheet-cell-input-A1"]').fill('Sheet1 Data');
    await page.keyboard.press('Enter');
    
    // Add a second sheet
    await page.locator('[data-testid="add-sheet-Workbook1"]').click();
    
    // Set up different content on Sheet2
    await page.locator('[data-testid="spreadsheet-cell-A1"]').dblclick();
    await page.locator('[data-testid="spreadsheet-cell-input-A1"]').fill('Sheet2 Data');
    await page.keyboard.press('Enter');
    
    // Verify Sheet2 content
    await expect(page.locator('[data-testid="spreadsheet-cell-A1"]')).toContainText('Sheet2 Data');
    
    // Click on Sheet1 tab text specifically to avoid nested buttons
    await page.locator('[data-testid="sheet-tab-Sheet1"] span').click();
    
    // Wait a moment for the UI to update
    await page.waitForTimeout(100);
    
    // Should switch to Sheet1 and show Sheet1 content
    await expect(page.locator('[data-testid="spreadsheet-cell-A1"]')).toContainText('Sheet1 Data');
    
    // Click on Sheet2 tab text specifically to avoid nested buttons
    await page.locator('[data-testid="sheet-tab-Sheet2"] span').click();
    
    // Should switch back to Sheet2 and show Sheet2 content
    await expect(page.locator('[data-testid="spreadsheet-cell-A1"]')).toContainText('Sheet2 Data');
  });

  test('should enter and evaluate basic formulas', async ({ page }) => {
    // Wait for spreadsheet to be ready
    await expect(page.locator('[data-testid="spreadsheet-container"]')).toBeVisible();
    
    // Click on cell A1 to select it
    await page.locator('[data-testid="spreadsheet-cell-A1"]').click();
    
    // Double-click to enter edit mode and fill with value
    await page.locator('[data-testid="spreadsheet-cell-A1"]').dblclick();
    await page.locator('[data-testid="spreadsheet-cell-input-A1"]').fill('42');
    await page.keyboard.press('Enter');
    
    // Click on cell B1 to select it
    await page.locator('[data-testid="spreadsheet-cell-B1"]').click();
    
    // Double-click to enter edit mode and fill with formula
    await page.locator('[data-testid="spreadsheet-cell-B1"]').dblclick();
    await page.locator('[data-testid="spreadsheet-cell-input-B1"]').fill('=A1*2');
    await page.keyboard.press('Enter');
    
    // Verify A1 has the expected value
    await expect(page.locator('[data-testid="spreadsheet-cell-A1"]')).toContainText('42');
    
    // Verify the result - should show calculated value, not formula
    await expect(page.locator('[data-testid="spreadsheet-cell-B1"]')).toContainText('84');
  });

  test('should evaluate formulas correctly', async ({ page }) => {
    // Wait for spreadsheet to be ready
    await expect(page.locator('[data-testid="spreadsheet-container"]')).toBeVisible();
    
    // Click on cell A1 to select it
    await page.locator('[data-testid="spreadsheet-cell-A1"]').click();
    
    // Double-click to enter edit mode and fill with formula
    await page.locator('[data-testid="spreadsheet-cell-A1"]').dblclick();
    await page.locator('[data-testid="spreadsheet-cell-input-A1"]').fill('=1+2+3');
    await page.keyboard.press('Enter');
    
    // Verify that the cell shows the calculated result
    await expect(page.locator('[data-testid="spreadsheet-cell-A1"]')).toContainText('6');
    
    // Test another formula
    await page.locator('[data-testid="spreadsheet-cell-B1"]').click();
    await page.locator('[data-testid="spreadsheet-cell-B1"]').dblclick();
    await page.locator('[data-testid="spreadsheet-cell-input-B1"]').fill('=A1*2');
    await page.keyboard.press('Enter');
    
    // Verify the second formula result
    await expect(page.locator('[data-testid="spreadsheet-cell-B1"]')).toContainText('12');
  });

  test('should use named expressions in formulas', async ({ page }) => {
    // Add a named expression
    await page.locator('[data-testid="named-expressions-toggle"]').click();
    await page.locator('[data-testid="expression-name-input"]').fill('DISCOUNT');
    await page.locator('[data-testid="expression-formula-input"]').fill('0.1');
    await page.locator('[data-testid="add-expression-button"]').click();
    
    // Close the named expressions panel
    await page.locator('[data-testid="named-expressions-toggle"]').click();
    
    // Wait for spreadsheet to be ready
    await expect(page.locator('[data-testid="spreadsheet-container"]')).toBeVisible();
    
    // Enter a price in A1
    await page.locator('[data-testid="spreadsheet-cell-A1"]').click();
    await page.locator('[data-testid="spreadsheet-cell-A1"]').dblclick();
    await page.locator('[data-testid="spreadsheet-cell-input-A1"]').fill('100');
    await page.keyboard.press('Enter');
    
    // Use the named expression in B1
    await page.locator('[data-testid="spreadsheet-cell-B1"]').click();
    await page.locator('[data-testid="spreadsheet-cell-B1"]').dblclick();
    await page.locator('[data-testid="spreadsheet-cell-input-B1"]').fill('=A1*DISCOUNT');
    await page.keyboard.press('Enter');
    
    // Should calculate discount (100 * 0.1 = 10)
    await expect(page.locator('[data-testid="spreadsheet-cell-B1"]')).toContainText('10');
  });

  test('should add, edit, and remove global named expressions with formula updates', async ({ page }) => {
    // Open named expressions panel
    await page.locator('[data-testid="named-expressions-toggle"]').click();
    
    // Add a global named expression
    await page.locator('[data-testid="expression-name-input"]').fill('TAX_RATE');
    await page.locator('[data-testid="expression-formula-input"]').fill('0.08');
    await page.locator('[data-testid="add-expression-button"]').click();
    
    // Verify it appears in the global section
    await expect(page.locator('[data-testid="expression-name-TAX_RATE"]')).toBeVisible();
    await expect(page.locator('[data-testid="expression-formula-TAX_RATE"]')).toContainText('0.08');
    
    // Close panel and use the named expression in a formula
    await page.locator('[data-testid="named-expressions-toggle"]').click();
    await page.locator('[data-testid="spreadsheet-cell-A1"]').dblclick();
    await page.locator('[data-testid="spreadsheet-cell-input-A1"]').fill('1000');
    await page.keyboard.press('Enter');
    
    await page.locator('[data-testid="spreadsheet-cell-B1"]').dblclick();
    await page.locator('[data-testid="spreadsheet-cell-input-B1"]').fill('=A1*TAX_RATE');
    await page.keyboard.press('Enter');
    
    // Should calculate tax (1000 * 0.08 = 80)
    await expect(page.locator('[data-testid="spreadsheet-cell-B1"]')).toContainText('80');
    
    // Now delete the named expression
    await page.locator('[data-testid="named-expressions-toggle"]').click();
    await page.locator('[data-testid="delete-global-named-expression-TAX_RATE"]').click();
    
    // Verify it's removed
    await expect(page.locator('[data-testid="expression-name-TAX_RATE"]')).not.toBeVisible();
    
    // Close panel and verify formula shows error
    await page.locator('[data-testid="named-expressions-toggle"]').click();
    
    // Formula should now show error since TAX_RATE is undefined
    await expect(page.locator('[data-testid="spreadsheet-cell-B1"]')).toContainText('#NAME?');
  });

  test('should add, edit, and remove sheet-scoped named expressions', async ({ page }) => {
    // Open named expressions panel
    await page.locator('[data-testid="named-expressions-toggle"]').click();
    
    // Change scope to "Current Sheet" to add sheet-scoped expression
    await page.locator('[data-testid="expression-scope-select"]').selectOption('sheet');
    
    // Add sheet-scoped named expression (placeholders will now show sheet examples)
    await page.locator('[data-testid="expression-name-input"]').fill('COMMISSION');
    await page.locator('[data-testid="expression-formula-input"]').fill('0.05');
    await page.locator('[data-testid="add-expression-button"]').click();
    
    // Verify it appears in the sheet section
    await expect(page.locator('text=COMMISSION')).toBeVisible();
    await expect(page.locator('text=0.05')).toBeVisible();
    
    // Close panel and use in formula
    await page.locator('[data-testid="named-expressions-toggle"]').click();
    await page.locator('[data-testid="spreadsheet-cell-A1"]').dblclick();
    await page.locator('[data-testid="spreadsheet-cell-input-A1"]').fill('2000');
    await page.keyboard.press('Enter');
    
    await page.locator('[data-testid="spreadsheet-cell-B1"]').dblclick();
    await page.locator('[data-testid="spreadsheet-cell-input-B1"]').fill('=A1*COMMISSION');
    await page.keyboard.press('Enter');
    
    // Should calculate commission (2000 * 0.05 = 100)
    await expect(page.locator('[data-testid="spreadsheet-cell-B1"]')).toContainText('100');
    
    // Add a second sheet to test sheet-scoped behavior
    await page.locator('[data-testid="add-sheet-Workbook1"]').click();
    
    // On Sheet2, the sheet-scoped COMMISSION from Sheet1 should not be available
    await page.locator('[data-testid="spreadsheet-cell-A1"]').dblclick();
    await page.locator('[data-testid="spreadsheet-cell-input-A1"]').fill('=1000*COMMISSION');
    await page.keyboard.press('Enter');
    
    // Should show error since COMMISSION is scoped to Sheet1
    await expect(page.locator('[data-testid="spreadsheet-cell-A1"]')).toContainText('#NAME?');
    
    // Switch back to Sheet1
    await page.locator('[data-testid="sheet-tab-Sheet1"] span').click();
    
    // Formula should still work on Sheet1
    await expect(page.locator('[data-testid="spreadsheet-cell-B1"]')).toContainText('100');
  });

  test('should handle named expressions across sheet operations', async ({ page }) => {
    // Add a global named expression
    await page.locator('[data-testid="named-expressions-toggle"]').click();
    
    // Ensure scope is set to "Global" (default)
    await page.locator('[data-testid="expression-scope-select"]').selectOption('global');
    await page.locator('[data-testid="expression-name-input"]').fill('GLOBAL_RATE');
    await page.locator('[data-testid="expression-formula-input"]').fill('0.15');
    await page.locator('[data-testid="add-expression-button"]').click();
    
    // Add a sheet-scoped named expression
    await page.locator('[data-testid="expression-scope-select"]').selectOption('sheet');
    await page.locator('[data-testid="expression-name-input"]').fill('LOCAL_RATE');
    await page.locator('[data-testid="expression-formula-input"]').fill('0.05');
    await page.locator('[data-testid="add-expression-button"]').click();
    
    await page.locator('[data-testid="named-expressions-toggle"]').click();
    
    // Use both in formulas
    await page.locator('[data-testid="spreadsheet-cell-A1"]').dblclick();
    await page.locator('[data-testid="spreadsheet-cell-input-A1"]').fill('1000');
    await page.keyboard.press('Enter');
    
    await page.locator('[data-testid="spreadsheet-cell-B1"]').dblclick();
    await page.locator('[data-testid="spreadsheet-cell-input-B1"]').fill('=A1*GLOBAL_RATE');
    await page.keyboard.press('Enter');
    
    await page.locator('[data-testid="spreadsheet-cell-C1"]').dblclick();
    await page.locator('[data-testid="spreadsheet-cell-input-C1"]').fill('=A1*LOCAL_RATE');
    await page.keyboard.press('Enter');
    
    // Verify calculations
    await expect(page.locator('[data-testid="spreadsheet-cell-B1"]')).toContainText('150'); // 1000 * 0.15
    await expect(page.locator('[data-testid="spreadsheet-cell-C1"]')).toContainText('50');  // 1000 * 0.05
    
    // Add a second sheet
    await page.locator('[data-testid="add-sheet-Workbook1"]').click();
    
    // On Sheet2, global should work but local should not
    await page.locator('[data-testid="spreadsheet-cell-A1"]').dblclick();
    await page.locator('[data-testid="spreadsheet-cell-input-A1"]').fill('=500*GLOBAL_RATE');
    await page.keyboard.press('Enter');
    
    await page.locator('[data-testid="spreadsheet-cell-B1"]').dblclick();
    await page.locator('[data-testid="spreadsheet-cell-input-B1"]').fill('=500*LOCAL_RATE');
    await page.keyboard.press('Enter');
    
    // Global should work, local should error
    await expect(page.locator('[data-testid="spreadsheet-cell-A1"]')).toContainText('75'); // 500 * 0.15
    await expect(page.locator('[data-testid="spreadsheet-cell-B1"]')).toContainText('#NAME?');
    
    // Rename Sheet1 and verify formulas still work
    await page.locator('[data-testid="sheet-tab-Sheet1"]').hover();
    await page.locator('[data-testid="rename-sheet-Sheet1"]').click();
    await page.locator('[data-testid="rename-sheet-input-Sheet1"]').fill('Data Sheet');
    await page.keyboard.press('Enter');
    
    // Switch to renamed sheet
    await page.locator('[data-testid="sheet-tab-Data Sheet"] span').click();
    
    // Formulas should still work after rename
    await expect(page.locator('[data-testid="spreadsheet-cell-B1"]')).toContainText('150');
    await expect(page.locator('[data-testid="spreadsheet-cell-C1"]')).toContainText('50');
  });

  test('should create tables from cell selection', async ({ page }) => {
    // Wait for spreadsheet to be ready
    await expect(page.locator('[data-testid="spreadsheet-container"]')).toBeVisible();
    
    // Enter table headers
    await page.locator('[data-testid="spreadsheet-cell-A1"]').click();
    await page.locator('[data-testid="spreadsheet-cell-A1"]').dblclick();
    await page.locator('[data-testid="spreadsheet-cell-input-A1"]').fill('Name');
    await page.keyboard.press('Enter');
    
    await page.locator('[data-testid="spreadsheet-cell-B1"]').click();
    await page.locator('[data-testid="spreadsheet-cell-B1"]').dblclick();
    await page.locator('[data-testid="spreadsheet-cell-input-B1"]').fill('Age');
    await page.keyboard.press('Enter');
    
    // Enter first row of data
    await page.locator('[data-testid="spreadsheet-cell-A2"]').click();
    await page.locator('[data-testid="spreadsheet-cell-A2"]').dblclick();
    await page.locator('[data-testid="spreadsheet-cell-input-A2"]').fill('John');
    await page.keyboard.press('Enter');
    
    await page.locator('[data-testid="spreadsheet-cell-B2"]').click();
    await page.locator('[data-testid="spreadsheet-cell-B2"]').dblclick();
    await page.locator('[data-testid="spreadsheet-cell-input-B2"]').fill('25');
    await page.keyboard.press('Enter');
    
    // Enter second row of data
    await page.locator('[data-testid="spreadsheet-cell-A3"]').click();
    await page.locator('[data-testid="spreadsheet-cell-A3"]').dblclick();
    await page.locator('[data-testid="spreadsheet-cell-input-A3"]').fill('Jane');
    await page.keyboard.press('Enter');
    
    await page.locator('[data-testid="spreadsheet-cell-B3"]').click();
    await page.locator('[data-testid="spreadsheet-cell-B3"]').dblclick();
    await page.locator('[data-testid="spreadsheet-cell-input-B3"]').fill('30');
    await page.keyboard.press('Enter');
    
    // Select the range A1:B3 for table creation
    await page.locator('[data-testid="spreadsheet-cell-A1"]').click();
    await page.locator('[data-testid="spreadsheet-cell-B3"]').click({ modifiers: ['Shift'] });
    
    // Create table using the button we added data-testid to
    await page.locator('[data-testid="create-table-button"]').click();
    
    // Verify table was created by checking if we can click on a table cell and see table management
    await page.locator('[data-testid="spreadsheet-cell-A1"]').click();
    await expect(page.locator('[data-testid="table-name-input"]')).toBeVisible();
    await expect(page.locator('[data-testid="rename-table-button"]')).toBeVisible();
    
    // Now select a different range to create another table and verify it shows Table2
    await page.locator('[data-testid="spreadsheet-cell-D1"]').click();
    await page.locator('[data-testid="spreadsheet-cell-E2"]').click({ modifiers: ['Shift'] });
    await expect(page.locator('[data-testid="table-name-input"]')).toHaveValue('Table2');
  });

  test('should create, rename, and remove tables with formula updates', async ({ page }) => {
    // Create a table with data
    await page.locator('[data-testid="spreadsheet-cell-A1"]').dblclick();
    await page.locator('[data-testid="spreadsheet-cell-input-A1"]').fill('Product');
    await page.keyboard.press('Enter');
    
    await page.locator('[data-testid="spreadsheet-cell-B1"]').dblclick();
    await page.locator('[data-testid="spreadsheet-cell-input-B1"]').fill('Price');
    await page.keyboard.press('Enter');
    
    await page.locator('[data-testid="spreadsheet-cell-A2"]').dblclick();
    await page.locator('[data-testid="spreadsheet-cell-input-A2"]').fill('Widget');
    await page.keyboard.press('Enter');
    
    await page.locator('[data-testid="spreadsheet-cell-B2"]').dblclick();
    await page.locator('[data-testid="spreadsheet-cell-input-B2"]').fill('10');
    await page.keyboard.press('Enter');
    
    await page.locator('[data-testid="spreadsheet-cell-A3"]').dblclick();
    await page.locator('[data-testid="spreadsheet-cell-input-A3"]').fill('Gadget');
    await page.keyboard.press('Enter');
    
    await page.locator('[data-testid="spreadsheet-cell-B3"]').dblclick();
    await page.locator('[data-testid="spreadsheet-cell-input-B3"]').fill('15');
    await page.keyboard.press('Enter');
    
    // Select range and create table
    await page.locator('[data-testid="spreadsheet-cell-A1"]').click();
    await page.locator('[data-testid="spreadsheet-cell-B3"]').click({ modifiers: ['Shift'] });
    
    // Set table name to "Products"
    await page.locator('[data-testid="table-name-input"]').fill('Products');
    await page.locator('[data-testid="create-table-button"]').click();
    
    // Use table in a formula
    await page.locator('[data-testid="spreadsheet-cell-D1"]').dblclick();
    await page.locator('[data-testid="spreadsheet-cell-input-D1"]').fill('=SUM(Products[Price])');
    await page.keyboard.press('Enter');
    
    // Should calculate sum of prices (10 + 15 = 25)
    await expect(page.locator('[data-testid="spreadsheet-cell-D1"]')).toContainText('25');
    
    // Test table column reference
    await page.locator('[data-testid="spreadsheet-cell-D2"]').dblclick();
    await page.locator('[data-testid="spreadsheet-cell-input-D2"]').fill('=AVERAGE(Products[Price])');
    await page.keyboard.press('Enter');
    
    // Should calculate average (25/2 = 12.5)
    await expect(page.locator('[data-testid="spreadsheet-cell-D2"]')).toContainText('12.5');
    
    // Now rename the table - first click on a table cell to show table management UI
    await page.locator('[data-testid="spreadsheet-cell-A1"]').click(); // Click on table cell
    await page.locator('[data-testid="table-name-input"]').fill('Inventory');
    await page.locator('[data-testid="rename-table-button"]').click();
    
    // Formulas should update to use new table name
    await expect(page.locator('[data-testid="spreadsheet-cell-D1"]')).toContainText('25'); // Still works
    await expect(page.locator('[data-testid="spreadsheet-cell-D2"]')).toContainText('12.5'); // Still works
    
    // Remove the table
    await page.locator('[data-testid="remove-table-button"]').click();
    
    // Formulas should now show errors since table no longer exists
    await expect(page.locator('[data-testid="spreadsheet-cell-D1"]')).toContainText('#REF!');
    await expect(page.locator('[data-testid="spreadsheet-cell-D2"]')).toContainText('#REF!');
  });

  test('should handle table operations across sheets', async ({ page }) => {
    // Create a table on Sheet1
    await page.locator('[data-testid="spreadsheet-cell-A1"]').dblclick();
    await page.locator('[data-testid="spreadsheet-cell-input-A1"]').fill('Category');
    await page.keyboard.press('Enter');
    
    await page.locator('[data-testid="spreadsheet-cell-B1"]').dblclick();
    await page.locator('[data-testid="spreadsheet-cell-input-B1"]').fill('Sales');
    await page.keyboard.press('Enter');
    
    await page.locator('[data-testid="spreadsheet-cell-A2"]').dblclick();
    await page.locator('[data-testid="spreadsheet-cell-input-A2"]').fill('Electronics');
    await page.keyboard.press('Enter');
    
    await page.locator('[data-testid="spreadsheet-cell-B2"]').dblclick();
    await page.locator('[data-testid="spreadsheet-cell-input-B2"]').fill('1000');
    await page.keyboard.press('Enter');
    
    // Create table
    await page.locator('[data-testid="spreadsheet-cell-A1"]').click();
    await page.locator('[data-testid="spreadsheet-cell-B2"]').click({ modifiers: ['Shift'] });
    await page.locator('[data-testid="table-name-input"]').fill('SalesData');
    await page.locator('[data-testid="create-table-button"]').click();
    
    // Add second sheet
    await page.locator('[data-testid="add-sheet-Workbook1"]').click();
    
    // On Sheet2, reference the table from Sheet1
    await page.locator('[data-testid="spreadsheet-cell-A1"]').dblclick();
    await page.locator('[data-testid="spreadsheet-cell-input-A1"]').fill('=SUM(Sheet1!SalesData[Sales])');
    await page.keyboard.press('Enter');
    
    // Should work with cross-sheet table reference
    await expect(page.locator('[data-testid="spreadsheet-cell-A1"]')).toContainText('1,000');
    
    // Go back to Sheet1 and rename it
    await page.locator('[data-testid="sheet-tab-Sheet1"]').hover();
    await page.locator('[data-testid="rename-sheet-Sheet1"]').click();
    await page.locator('[data-testid="rename-sheet-input-Sheet1"]').fill('Sales');
    await page.keyboard.press('Enter');
    
    // Switch back to Sheet2
    await page.locator('[data-testid="sheet-tab-Sheet2"] span').click();
    
    // Formula should still work after sheet rename (engine should update references)
    await expect(page.locator('[data-testid="spreadsheet-cell-A1"]')).toContainText('1,000');
  });

  test('should handle cross-sheet formulas with sheet operations', async ({ page }) => {
    // Set up data on Sheet1
    await page.locator('[data-testid="spreadsheet-cell-A1"]').dblclick();
    await page.locator('[data-testid="spreadsheet-cell-input-A1"]').fill('100');
    await page.keyboard.press('Enter');
    
    await page.locator('[data-testid="spreadsheet-cell-B1"]').dblclick();
    await page.locator('[data-testid="spreadsheet-cell-input-B1"]').fill('200');
    await page.keyboard.press('Enter');
    
    // Add second sheet
    await page.locator('[data-testid="add-sheet-Workbook1"]').click();
    
    // On Sheet2, create formulas that reference Sheet1
    await page.locator('[data-testid="spreadsheet-cell-A1"]').dblclick();
    await page.locator('[data-testid="spreadsheet-cell-input-A1"]').fill('=Sheet1!A1+Sheet1!B1');
    await page.keyboard.press('Enter');
    
    await page.locator('[data-testid="spreadsheet-cell-B1"]').dblclick();
    await page.locator('[data-testid="spreadsheet-cell-input-B1"]').fill('=Sheet1!A1*2');
    await page.keyboard.press('Enter');
    
    // Verify cross-sheet formulas work
    await expect(page.locator('[data-testid="spreadsheet-cell-A1"]')).toContainText('300'); // 100 + 200
    await expect(page.locator('[data-testid="spreadsheet-cell-B1"]')).toContainText('200'); // 100 * 2
    
    // Rename Sheet1
    await page.locator('[data-testid="sheet-tab-Sheet1"]').hover();
    await page.locator('[data-testid="rename-sheet-Sheet1"]').click();
    await page.locator('[data-testid="rename-sheet-input-Sheet1"]').fill('Data');
    await page.keyboard.press('Enter');
    
    // Switch back to Sheet2
    await page.locator('[data-testid="sheet-tab-Sheet2"] span').click();
    
    // Formulas should still work after sheet rename
    await expect(page.locator('[data-testid="spreadsheet-cell-A1"]')).toContainText('300');
    await expect(page.locator('[data-testid="spreadsheet-cell-B1"]')).toContainText('200');
    
    // Change data on renamed sheet
    await page.locator('[data-testid="sheet-tab-Data"] span').click();
    await page.locator('[data-testid="spreadsheet-cell-A1"]').dblclick();
    await page.locator('[data-testid="spreadsheet-cell-input-A1"]').fill('150');
    await page.keyboard.press('Enter');
    
    // Switch back to Sheet2 and verify formulas updated
    await page.locator('[data-testid="sheet-tab-Sheet2"] span').click();
    await expect(page.locator('[data-testid="spreadsheet-cell-A1"]')).toContainText('350'); // 150 + 200
    await expect(page.locator('[data-testid="spreadsheet-cell-B1"]')).toContainText('300'); // 150 * 2
    
    // Delete the Data sheet (this should cause formulas to error)
    await page.locator('[data-testid="sheet-tab-Data"]').hover();
    
    // Handle the confirmation dialog
    page.on('dialog', dialog => dialog.accept());
    
    await page.locator('[data-testid="delete-sheet-Data"]').click();
    
    // Formulas should now show errors since referenced sheet is gone
    await expect(page.locator('[data-testid="spreadsheet-cell-A1"]')).toContainText('#REF!');
    await expect(page.locator('[data-testid="spreadsheet-cell-B1"]')).toContainText('#REF!');
  });

});
