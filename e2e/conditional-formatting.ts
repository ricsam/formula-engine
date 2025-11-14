import { test, expect } from '@playwright/test';

test.describe('Conditional Formatting', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/excel');
    // Wait for the page to fully load
    await expect(page.locator('h1')).toContainText('FormulaEngine Multi-Workbook Demo');
    
    // Open the Expressions, Tables & Formatting panel
    await page.locator('[data-testid="named-expressions-toggle"]').click();
    await page.waitForTimeout(200);
    
    // Check that the conditional formatting section is visible
    await expect(page.locator('[data-testid="conditional-formatting-section"]')).toBeVisible();
  });

  test('should display conditional formatting UI', async ({ page }) => {
    // Check title
    await expect(page.locator('[data-testid="conditional-formatting-title"]')).toContainText('Conditional Formatting');
    
    // Check range input is present (scoped to conditional formatting section)
    const conditionalSection = page.locator('[data-testid="conditional-formatting-section"]');
    const rangeInput = conditionalSection.locator('input[placeholder*="Select a range"]');
    await expect(rangeInput).toBeVisible();
    
    // Check type selector
    const typeSelector = conditionalSection.locator('select').filter({ hasText: 'Formula' });
    await expect(typeSelector).toBeVisible();
    
    // Check Add Rule button
    await expect(conditionalSection.locator('button:has-text("Add Rule")')).toBeVisible();
  });

  test('should add formula-based conditional formatting', async ({ page }) => {
    // Fill in the range - select some cells first
    const rangeInput = page.locator('input[placeholder*="Select a range"]').first();
    await rangeInput.fill('[Workbook1]Sheet1!A1:A10');
    
    // Make sure Formula type is selected
    const typeSelector = page.locator('select').filter({ hasText: 'Formula' }).first();
    await typeSelector.selectOption('formula');
    
    // Wait for formula input to appear
    await page.waitForTimeout(100);
    
    // Fill in formula
    const formulaInput = page.locator('input[placeholder*="ROW()"]');
    await formulaInput.fill('ROW() > 5');
    
    // Click Add Rule button
    await page.locator('button:has-text("Add Rule")').click();
    
    // Wait for the style to be added
    await page.waitForTimeout(200);
    
    // Check that the style appears in the list
    const conditionalStylesList = page.locator('[data-testid="conditional-formatting-list"]');
    await expect(conditionalStylesList.locator('[data-testid^="conditional-style-"]')).toHaveCount(1);
    
    // Check the style details
    const firstStyle = conditionalStylesList.locator('[data-testid="conditional-style-0"]');
    await expect(firstStyle).toContainText('Workbook1');
    await expect(firstStyle).toContainText('Sheet1');
    await expect(firstStyle).toContainText('A1:A10');
    await expect(firstStyle).toContainText('Formula: ROW() > 5');
  });

  test('should add gradient-based conditional formatting with lowest/highest values', async ({ page }) => {
    // Fill in the range
    const rangeInput = page.locator('input[placeholder*="Select a range"]').first();
    await rangeInput.fill('[Workbook1]Sheet1!B1:B5');
    
    // Select Gradient type
    const typeSelector = page.locator('select').filter({ hasText: 'Formula' }).first();
    await typeSelector.selectOption('gradient');
    
    // Wait for gradient options to appear
    await page.waitForTimeout(100);
    
    // Check that gradient type selector is visible
    const gradientTypeSelector = page.locator('select').filter({ hasText: 'Lowest to Highest Value' });
    await expect(gradientTypeSelector).toBeVisible();
    
    // Keep default "Lowest to Highest Value"
    
    // Click Add Rule button
    await page.locator('button:has-text("Add Rule")').click();
    
    // Wait for the style to be added
    await page.waitForTimeout(200);
    
    // Check that the style appears in the list
    const conditionalStylesList = page.locator('[data-testid="conditional-formatting-list"]');
    await expect(conditionalStylesList.locator('[data-testid^="conditional-style-"]')).toHaveCount(1);
    
    // Check the style details
    const firstStyle = conditionalStylesList.locator('[data-testid="conditional-style-0"]');
    await expect(firstStyle).toContainText('B1:B5');
    await expect(firstStyle).toContainText('Gradient: Min to Max');
    
    // Check that color preview is visible
    await expect(firstStyle.locator('div[title="Color preview"]')).toBeVisible();
  });

  test('should add gradient with custom min/max formulas', async ({ page }) => {
    // Fill in the range
    const rangeInput = page.locator('input[placeholder*="Select a range"]').first();
    await rangeInput.fill('[Workbook1]Sheet1!C1:C10');
    
    // Select Gradient type
    const typeSelector = page.locator('select').filter({ hasText: 'Formula' }).first();
    await typeSelector.selectOption('gradient');
    
    // Wait for gradient options to appear
    await page.waitForTimeout(100);
    
    // Select Custom Min/Max
    const gradientTypeSelector = page.locator('select').filter({ hasText: 'Lowest to Highest Value' });
    await gradientTypeSelector.selectOption('number');
    
    // Wait for formula inputs to appear
    await page.waitForTimeout(100);
    
    // Fill in custom min/max formulas
    const minFormulaInput = page.locator('input[placeholder*="Min formula"]');
    await minFormulaInput.fill('0');
    
    const maxFormulaInput = page.locator('input[placeholder*="Max formula"]');
    await maxFormulaInput.fill('100');
    
    // Click Add Rule button
    await page.locator('button:has-text("Add Rule")').click();
    
    // Wait for the style to be added
    await page.waitForTimeout(200);
    
    // Check that the style appears in the list
    const conditionalStylesList = page.locator('[data-testid="conditional-formatting-list"]');
    await expect(conditionalStylesList.locator('[data-testid^="conditional-style-"]')).toHaveCount(1);
    
    // Check the style details
    const firstStyle = conditionalStylesList.locator('[data-testid="conditional-style-0"]');
    await expect(firstStyle).toContainText('Gradient: 0 to 100');
  });

  test('should delete conditional formatting rule', async ({ page }) => {
    // Add a rule first
    const rangeInput = page.locator('input[placeholder*="Select a range"]').first();
    await rangeInput.fill('[Workbook1]Sheet1!A1:A5');
    
    const formulaInput = page.locator('input[placeholder*="ROW()"]');
    await formulaInput.fill('TRUE');
    
    await page.locator('button:has-text("Add Rule")').click();
    await page.waitForTimeout(200);
    
    // Verify rule exists
    const conditionalStylesList = page.locator('[data-testid="conditional-formatting-list"]');
    await expect(conditionalStylesList.locator('[data-testid^="conditional-style-"]')).toHaveCount(1);
    
    // Delete the rule
    await page.locator('[data-testid="delete-conditional-style-0"]').click();
    await page.waitForTimeout(200);
    
    // Verify rule is deleted
    await expect(conditionalStylesList.locator('[data-testid^="conditional-style-"]')).toHaveCount(0);
    await expect(conditionalStylesList).toContainText('No conditional formatting rules');
  });

  test('should apply formula-based styling to cells', async ({ page }) => {
    // First, add some numeric data to cells using the same pattern as excel-demo.ts
    await page.locator('[data-testid="spreadsheet-cell-A1"]').dblclick();
    await page.locator('[data-testid="spreadsheet-cell-input-A1"]').fill('1');
    await page.keyboard.press('Enter');
    await page.waitForTimeout(100);
    
    await page.locator('[data-testid="spreadsheet-cell-A2"]').dblclick();
    await page.locator('[data-testid="spreadsheet-cell-input-A2"]').fill('2');
    await page.keyboard.press('Enter');
    await page.waitForTimeout(100);
    
    await page.locator('[data-testid="spreadsheet-cell-A3"]').dblclick();
    await page.locator('[data-testid="spreadsheet-cell-input-A3"]').fill('10');
    await page.keyboard.press('Enter');
    await page.waitForTimeout(100);
    
    // Add conditional formatting
    const rangeInput = page.locator('input[placeholder*="Select a range"]').first();
    await rangeInput.fill('[Workbook1]Sheet1!A1:A10');
    
    const formulaInput = page.locator('input[placeholder*="ROW()"]');
    await formulaInput.fill('ROW() > 2');
    
    await page.locator('button:has-text("Add Rule")').click();
    await page.waitForTimeout(300);
    
    // Close the panel to see the cells better
    await page.locator('[data-testid="named-expressions-toggle"]').click();
    await page.waitForTimeout(200);
    
    // Check that cells have the conditional-background attribute
    // Cell A1 (row 1) should not have styling (ROW() = 1, not > 2)
    const cellA1 = page.locator('[data-testid="spreadsheet-cell-A1"]');
    const cellA1HasStyle = await cellA1.locator('[data-conditional-background]').count();
    expect(cellA1HasStyle).toBe(0); // A1 should not have conditional style
    
    // Cell A2 (row 2) should not have styling (ROW() = 2, not > 2)
    const cellA2 = page.locator('[data-testid="spreadsheet-cell-A2"]');
    const cellA2HasStyle = await cellA2.locator('[data-conditional-background]').count();
    expect(cellA2HasStyle).toBe(0); // A2 should not have conditional style
    
    // Cell A3 (row 3) should have styling (ROW() = 3, > 2)
    const cellA3 = page.locator('[data-testid="spreadsheet-cell-A3"]');
    const cellA3StyleDiv = cellA3.locator('[data-conditional-background]');
    await expect(cellA3StyleDiv).toBeVisible();
    const cellA3Style = await cellA3StyleDiv.getAttribute('data-conditional-background');
    expect(cellA3Style).toBeTruthy(); // A3 should have a hex color
    expect(cellA3Style).toMatch(/^#[0-9A-Fa-f]{6}$/); // Should be a valid hex color
  });

  test('should apply gradient styling to numeric cells', async ({ page }) => {
    // Add numeric data using the same pattern as excel-demo.ts
    await page.locator('[data-testid="spreadsheet-cell-B1"]').dblclick();
    await page.locator('[data-testid="spreadsheet-cell-input-B1"]').fill('10');
    await page.keyboard.press('Enter');
    await page.waitForTimeout(100);
    
    await page.locator('[data-testid="spreadsheet-cell-B2"]').dblclick();
    await page.locator('[data-testid="spreadsheet-cell-input-B2"]').fill('50');
    await page.keyboard.press('Enter');
    await page.waitForTimeout(100);
    
    await page.locator('[data-testid="spreadsheet-cell-B3"]').dblclick();
    await page.locator('[data-testid="spreadsheet-cell-input-B3"]').fill('90');
    await page.keyboard.press('Enter');
    await page.waitForTimeout(100);
    
    // Add gradient conditional formatting
    const rangeInput = page.locator('input[placeholder*="Select a range"]').first();
    await rangeInput.fill('[Workbook1]Sheet1!B1:B3');
    
    const typeSelector = page.locator('select').filter({ hasText: 'Formula' }).first();
    await typeSelector.selectOption('gradient');
    await page.waitForTimeout(100);
    
    await page.locator('button:has-text("Add Rule")').click();
    await page.waitForTimeout(300);
    
    // Close the panel
    await page.locator('[data-testid="named-expressions-toggle"]').click();
    await page.waitForTimeout(200);
    
    // Check that cells have conditional styling applied with different colors
    // B1 (10) should have a color closer to min color
    const cellB1 = page.locator('[data-testid="spreadsheet-cell-B1"]');
    const cellB1StyleDiv = cellB1.locator('[data-conditional-background]');
    await expect(cellB1StyleDiv).toBeVisible();
    const cellB1Style = await cellB1StyleDiv.getAttribute('data-conditional-background');
    expect(cellB1Style).toBeTruthy();
    expect(cellB1Style).toMatch(/^#[0-9A-Fa-f]{6}$/);
    
    // B2 (50) should have an intermediate color
    const cellB2 = page.locator('[data-testid="spreadsheet-cell-B2"]');
    const cellB2StyleDiv = cellB2.locator('[data-conditional-background]');
    await expect(cellB2StyleDiv).toBeVisible();
    const cellB2Style = await cellB2StyleDiv.getAttribute('data-conditional-background');
    expect(cellB2Style).toBeTruthy();
    expect(cellB2Style).toMatch(/^#[0-9A-Fa-f]{6}$/);
    // B2 should have a different color than B1 (interpolated)
    expect(cellB2Style).not.toBe(cellB1Style);
    
    // B3 (90) should have a color closer to max color
    const cellB3 = page.locator('[data-testid="spreadsheet-cell-B3"]');
    const cellB3StyleDiv = cellB3.locator('[data-conditional-background]');
    await expect(cellB3StyleDiv).toBeVisible();
    const cellB3Style = await cellB3StyleDiv.getAttribute('data-conditional-background');
    expect(cellB3Style).toBeTruthy();
    expect(cellB3Style).toMatch(/^#[0-9A-Fa-f]{6}$/);
    // B3 should have a different color than B1 and B2
    expect(cellB3Style).not.toBe(cellB1Style);
    expect(cellB3Style).not.toBe(cellB2Style);
  });

  test('should update range input when cells are selected', async ({ page }) => {
    // Get the range input
    const rangeInput = page.locator('input[placeholder*="Select a range"]').first();
    
    // Clear it first
    await rangeInput.fill('');
    
    // Click on A1 to start selection
    await page.locator('[data-testid="spreadsheet-cell-A1"]').click();
    await page.waitForTimeout(200);
    
    // Click with shift to select a range (similar to excel-demo.ts pattern)
    await page.locator('[data-testid="spreadsheet-cell-C3"]').click({ modifiers: ['Shift'] });
    await page.waitForTimeout(200);
    
    // Check that range input was updated
    const inputValue = await rangeInput.inputValue();
    
    // Should contain workbook and sheet reference
    expect(inputValue).toContain('[Workbook1]');
    expect(inputValue).toContain('Sheet1');
    expect(inputValue).toContain(':');
  });

  test('should support canonical range formats', async ({ page }) => {
    // Test closed rectangle format
    let rangeInput = page.locator('input[placeholder*="Select a range"]').first();
    await rangeInput.fill('[Workbook1]Sheet1!A1:D10');
    
    let formulaInput = page.locator('input[placeholder*="ROW()"]');
    await formulaInput.fill('TRUE');
    
    await page.locator('button:has-text("Add Rule")').click();
    await page.waitForTimeout(200);
    
    // Verify it was added
    let conditionalStylesList = page.locator('[data-testid="conditional-formatting-list"]');
    await expect(conditionalStylesList.locator('[data-testid="conditional-style-0"]')).toContainText('A1:D10');
    
    // Test row-bounded format (col-infinite)
    await rangeInput.fill('[Workbook1]Sheet1!A5:10');
    await page.locator('button:has-text("Add Rule")').click();
    await page.waitForTimeout(200);
    
    await expect(conditionalStylesList.locator('[data-testid="conditional-style-1"]')).toContainText('A5:');
    
    // Test col-bounded format (row-infinite)
    await rangeInput.fill('[Workbook1]Sheet1!A5:D');
    await page.locator('button:has-text("Add Rule")').click();
    await page.waitForTimeout(200);
    
    await expect(conditionalStylesList.locator('[data-testid="conditional-style-2"]')).toContainText('A5:');
    
    // Test open both format
    await rangeInput.fill('[Workbook1]Sheet1!A5:INFINITY');
    await page.locator('button:has-text("Add Rule")').click();
    await page.waitForTimeout(200);
    
    await expect(conditionalStylesList.locator('[data-testid="conditional-style-3"]')).toContainText('A5:');
  });

  test('should serialize and deserialize conditional formatting', async ({ page }) => {
    // Add a conditional formatting rule
    const rangeInput = page.locator('input[placeholder*="Select a range"]').first();
    await rangeInput.fill('[Workbook1]Sheet1!A1:A10');
    
    const formulaInput = page.locator('input[placeholder*="ROW()"]');
    await formulaInput.fill('ROW() > 5');
    
    await page.locator('button:has-text("Add Rule")').click();
    await page.waitForTimeout(200);
    
    // Verify the rule was added
    const conditionalStylesList = page.locator('[data-testid="conditional-formatting-list"]');
    await expect(conditionalStylesList.locator('[data-testid^="conditional-style-"]')).toHaveCount(1);
    
    // Save the file
    await page.locator('[data-testid="save-button"]').click();
    await page.waitForTimeout(500);
    
    // Verify save button shows "Saved"
    await expect(page.locator('[data-testid="save-button"]')).toContainText('Saved');
    
    // Reload the page
    await page.reload();
    
    // Wait for the page to be fully loaded
    await expect(page.locator('h1')).toContainText('FormulaEngine Multi-Workbook Demo');
    
    // Wait for the spreadsheet to be ready (indicates engine is loaded)
    await expect(page.locator('[data-testid="spreadsheet-container"]')).toBeVisible();
    
    // Give more time for OPFS to load and engine to restore state
    await page.waitForTimeout(3000);
    
    // Open the panel again
    await page.locator('[data-testid="named-expressions-toggle"]').click();
    await page.waitForTimeout(500);
    
    // Check that the conditional formatting rule is still there
    const reloadedList = page.locator('[data-testid="conditional-formatting-list"]');
    await expect(reloadedList.locator('[data-testid^="conditional-style-"]')).toHaveCount(1);
    
    const firstStyle = reloadedList.locator('[data-testid="conditional-style-0"]');
    await expect(firstStyle).toContainText('A1:A10');
    await expect(firstStyle).toContainText('Formula: ROW() > 5');
  });

  test('should show color preview for formula-based styles', async ({ page }) => {
    // Add a formula-based style
    const rangeInput = page.locator('input[placeholder*="Select a range"]').first();
    await rangeInput.fill('[Workbook1]Sheet1!A1:A5');
    
    const formulaInput = page.locator('input[placeholder*="ROW()"]');
    await formulaInput.fill('TRUE');
    
    // Click Add Rule button
    await page.locator('button:has-text("Add Rule")').click();
    await page.waitForTimeout(200);
    
    // Check that color preview is visible
    const firstStyle = page.locator('[data-testid="conditional-style-0"]');
    const colorPreview = firstStyle.locator('div[title="Color preview"]');
    
    await expect(colorPreview).toBeVisible();
    
    // Check that it has a background style
    const bgStyle = await colorPreview.getAttribute('style');
    expect(bgStyle).toContain('background');
  });

  test('should show gradient preview for gradient styles', async ({ page }) => {
    // Add a gradient style
    const rangeInput = page.locator('input[placeholder*="Select a range"]').first();
    await rangeInput.fill('[Workbook1]Sheet1!B1:B5');
    
    const typeSelector = page.locator('select').filter({ hasText: 'Formula' }).first();
    await typeSelector.selectOption('gradient');
    await page.waitForTimeout(100);
    
    await page.locator('button:has-text("Add Rule")').click();
    await page.waitForTimeout(200);
    
    // Check that color preview shows a gradient
    const firstStyle = page.locator('[data-testid="conditional-style-0"]');
    const colorPreview = firstStyle.locator('div[title="Color preview"]');
    
    await expect(colorPreview).toBeVisible();
    
    // Check that it has a gradient background
    const bgStyle = await colorPreview.getAttribute('style');
    expect(bgStyle).toContain('gradient');
  });
});

